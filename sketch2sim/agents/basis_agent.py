import json
import os
import re
from typing import Any

from langchain_core.messages import HumanMessage, SystemMessage
from langchain_core.runnables import RunnableLambda
from langchain_ollama import ChatOllama
from langgraph.graph import StateGraph
from pydantic import BaseModel

from ..RAG.vectorless_rag import build_component_retriever
from ..simulators import SimulatorProfile, get_simulator
from ..utils.logging_utils import log_prompt

_STEP4_MARKER = r"# === Step 4: Add Unit Operations \(Agent 2\) ==="
_CLEAN_SUFFIX = {"feed", "fresh", "batch", "raw"}
# Set VECTORLESS_RAG_REASONING=0 to use only the lexical stage of the retriever.
_RAG_REASONING = os.getenv("VECTORLESS_RAG_REASONING", "1") != "0"


class BasisState(BaseModel):
    description: str
    template_path: str = ""  # empty -> the simulator's default template
    simulator: str = "hysys"  # "hysys" | "aspenplus"
    rag_reasoning_model: str | None = None  # default: the Basis model itself
    log_dir: str = "logs"
    structure_json: str | None = None
    materials: list[str] | None = None
    python_code: str | None = None
    summary: str | None = None
    raw_output: str | None = None


def _load_text(path: str) -> str:
    with open(path, encoding="utf-8") as file:
        return file.read()


def _extract_code(text: str) -> str:
    matches = re.findall(
        r"```(?:python|py)?\s*(.*?)```", text, re.DOTALL | re.IGNORECASE
    )
    return (
        "\n\n".join(block.strip() for block in matches if block.strip())
        if matches
        else text.strip()
    )


def _clean_name(name: str) -> str:
    tokens = re.sub(r"\s+", " ", (name or "").strip()).split()

    while tokens and tokens[0].lower() in _CLEAN_SUFFIX:
        tokens.pop(0)

    while tokens and tokens[-1].lower() in _CLEAN_SUFFIX:
        tokens.pop()

    return " ".join(tokens)


def _extract_materials(structure_json: Any) -> list[str]:
    try:
        if isinstance(structure_json, dict):
            data = structure_json
        elif isinstance(structure_json, str):
            text = structure_json.strip()
            if os.path.exists(text):
                with open(text, encoding="utf-8") as file:
                    data = json.load(file)
            else:
                data = json.loads(text)
        else:
            return []
    except Exception:
        return []

    materials = []
    for stream in data.get("feed_streams", []):
        name = stream.get("name") or stream.get("id")
        if name:
            materials.append(_clean_name(name))

    unique_materials = []
    seen = set()
    for material in materials:
        if material and material not in seen:
            seen.add(material)
            unique_materials.append(material)

    return unique_materials


def _build_knowledge_base(
    materials: list[str], reasoning_model: str | None = None
) -> str:
    if not materials:
        return ""

    retriever = build_component_retriever(reasoning_model=reasoning_model)

    chunks = []
    for material in materials:
        documents = retriever.invoke(material)
        chunks.append(f"Material: {material}")

        if not documents:
            chunks.append("No matches found.\n")
            continue

        for document in documents[:2]:
            chunks.append(document.page_content)
        chunks.append("")

    return "\n".join(chunks).strip()


def _apply_safety_belt(template: str, llm_code: str) -> str:
    match = re.search(rf"(?s)(# === Step 1:.*?)(?={_STEP4_MARKER})", llm_code)
    if not match:
        raise ValueError("Could not extract Steps 1–3 from Basis output.")

    step_block = match.group(1).rstrip()
    pattern = rf"(?s)# === Step 1:.*?(?={_STEP4_MARKER})"

    if not re.search(pattern, template):
        raise ValueError("Template missing Step 1/4 markers.")

    return re.sub(pattern, step_block + "\n\n", template, count=1)


def _system_prompt(profile: SimulatorProfile) -> str:
    water = "WATER" if profile.key == "aspenplus" else "H2O"
    return (
        f"You are the Basis Agent for {profile.display_name}. Modify only:\n"
        "- case_name\n"
        f"- {profile.property_package_field} "
        f"(examples: {profile.property_package_examples})\n"
        "- Step 3 components section\n\n"
        f"Use only valid {profile.display_name} names from KB lines:\n"
        f"- '{profile.component_field}:'\n"
        f"- '{profile.recipe_field}:'\n"
        "Ignore the other name fields in the KB.\n\n"
        "If KB says 'No matches found.' then replace that material with:\n"
        f'components.extend(["{water}"]).\n\n'
        "Follow strict Step 3 pattern:\n"
        "- Initial components = [descriptive names]\n"
        "- One if-block per entry replacing it\n"
        "- Keep final loop adding components\n"
        "Output only the full Python file."
    )


def _user_prompt(
    description: str, template: str, materials: list[str], knowledge_base: str
) -> str:
    material_lines = "\n".join(f"- {material}" for material in materials)

    return (
        f"Process description:\n{description}\n\n"
        f"Feed materials:\n{material_lines}\n\n"
        f"KB:\n{knowledge_base}\n\n"
        "Template:\n```python\n"
        f"{template}\n"
        "```\n\n"
        "Return only the full modified Python file."
    )


def generate_basis_script(state: BasisState, model_name: str) -> dict[str, Any]:
    profile = get_simulator(state.simulator)
    template = _load_text(state.template_path or str(profile.template_path))
    materials = state.materials or (
        _extract_materials(state.structure_json) if state.structure_json else []
    )
    reasoning_model = (
        (state.rag_reasoning_model or model_name) if _RAG_REASONING else None
    )
    knowledge_base = _build_knowledge_base(materials, reasoning_model)

    system_prompt = _system_prompt(profile)
    user_prompt = _user_prompt(state.description, template, materials, knowledge_base)

    log_prompt(
        log_dir=state.log_dir,
        filename="basis_prompt_debug.txt",
        system_prompt=system_prompt,
        human_prompt=user_prompt,
    )

    llm = ChatOllama(model=model_name, temperature=0)
    response = llm.invoke(
        [
            SystemMessage(content=system_prompt),
            HumanMessage(content=user_prompt),
        ]
    )

    raw_output = response.content or ""
    python_code = _extract_code(raw_output)
    safe_code = _apply_safety_belt(template, python_code)

    return {
        "python_code": safe_code,
        "raw_output": raw_output,
        "summary": "Basis script generated.",
        "output_path": None,
    }


def build_basis_graph(model_name: str):
    graph = StateGraph(BasisState)
    graph.add_node(
        "basis",
        RunnableLambda(lambda state: generate_basis_script(state, model_name)),
    )
    graph.set_entry_point("basis")
    graph.set_finish_point("basis")
    return graph.compile()
