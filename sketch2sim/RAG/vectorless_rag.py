"""
Vectorless RAG for component lookup.

No embeddings and no vector database. Retrieval works in two stages over a
structured, hierarchical index of the component catalog:

1. Lexical retrieval (deterministic, pure Python)
   - exact match on any key (simulator name, alias, formula, mixture name)
   - BM25 over name/alias/formula tokens
   - fuzzy string similarity (difflib) to absorb typos and spelling variants

2. Reasoning selection (optional, LLM)
   The top candidates are shown to an LLM as a compact table of contents
   (node id + one-line summary). The LLM reasons about which entries match the
   query and returns node ids. If the LLM is unavailable or answers badly, the
   lexical ranking is used unchanged.

The index groups every catalog row that shares a simulator name into one node,
so aliases such as "Water" and "H2O" point at the same component.
"""

from __future__ import annotations

import json
import math
import re
from collections import Counter
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from difflib import SequenceMatcher, get_close_matches
from functools import lru_cache
from pathlib import Path

import pandas as pd
from langchain_core.documents import Document

RAG_DIR = Path(__file__).resolve().parent
DEFAULT_EXCEL_PATH = RAG_DIR / "components_list.xlsx"
DEFAULT_MIXTURE_PATH = RAG_DIR / "mixture_recipes.txt"

_RECIPE_SEPARATOR = "========================================"
_NOISE_WORDS = {"feed", "fresh", "raw", "batch", "stream", "inlet", "make-up", "makeup"}

# Everyday words that the catalog does not list as aliases.
SYNONYMS: dict[str, str] = {
    "steam": "Water",
    "process water": "Water",
    "wash water": "Water",
    "effluent water": "Water",
    "brine": "Water",
    "caustic": "Caustic Solution",
    "caustic soda": "Caustic Solution",
    "crude": "Crude Oil",
    "kerosene": "Jet Fuel",
    "diesel": "Diesel Blend",
    "air": "Air",
}

# Aspen Plus databank names that do not follow the generic naming rule.
ASPEN_NAME_OVERRIDES: dict[str, str] = {
    "H2O": "WATER",
    "CO2": "CARBON-DIOXIDE",
    "CO": "CARBON-MONOXIDE",
    "H2S": "HYDROGEN-SULFIDE",
    "SO2": "SULFUR-DIOXIDE",
    "NH3": "AMMONIA",
    "NACL": "SODIUM-CHLORIDE",
    "NaOH": "SODIUM-HYDROXIDE",
    "Hydrogen": "HYDROGEN",
    "Nitrogen": "NITROGEN",
    "Oxygen": "OXYGEN",
    "Air": "AIR",
}


def normalize(text: str) -> str:
    """Lowercase, unify separators and collapse whitespace."""
    text = str(text or "").replace("_", " ").strip().lower()
    return re.sub(r"\s+", " ", text)


def tokenize(text: str) -> list[str]:
    return [token for token in re.split(r"[^a-z0-9]+", normalize(text)) if token]


def clean_query(text: str) -> str:
    """Drop words like 'fresh' or 'feed' that describe the stream, not the chemical."""
    tokens = normalize(text).split(" ")
    while tokens and tokens[0] in _NOISE_WORDS:
        tokens.pop(0)
    while tokens and tokens[-1] in _NOISE_WORDS:
        tokens.pop()
    return " ".join(tokens)


def aspen_databank_name(simulator_name: str, aliases: Iterable[str] = ()) -> str:
    """
    Best-effort Aspen Plus databank name for a catalog component.

    Rules: explicit override, else a word-like name ("Toluene" -> "TOLUENE",
    "Carbon_Dioxide" -> "CARBON-DIOXIDE", "i-Butane" -> "ISOBUTANE"), else the
    simulator name upper-cased. Verify unusual components in Aspen Plus.
    """
    if simulator_name in ASPEN_NAME_OVERRIDES:
        return ASPEN_NAME_OVERRIDES[simulator_name]

    def word_like(name: str) -> bool:
        letters = re.sub(r"[^A-Za-z]", "", name)
        return (
            len(letters) >= 4
            and re.fullmatch(r"[A-Za-z0-9,\-_ ]+", name) is not None
            and not re.fullmatch(r"(?:[A-Z][a-z]?\d*)+", name)  # not a formula
        )

    for candidate in [simulator_name, *aliases]:
        if word_like(candidate):
            name = re.sub(r"[\s_]+", "-", candidate.strip()).upper()
            name = re.sub(r"^I-", "ISO", name)
            return name
    return simulator_name.upper()


@dataclass
class IndexNode:
    """One retrievable unit: a pure component or a mixture recipe."""

    node_id: str
    kind: str  # "pure_component" | "pseudo_recipe"
    title: str  # simulator (HYSYS) name or mixture name
    keys: list[str]  # everything that should match exactly
    text: str  # what is handed to the LLM / Basis agent
    metadata: dict = field(default_factory=dict)

    def summary(self) -> str:
        if self.kind == "pure_component":
            aliases = ", ".join(self.metadata.get("aliases", [])[:4])
            return f"pure component '{self.title}' (formula {self.metadata.get('formula')}; aliases: {aliases})"
        return f"mixture recipe '{self.title}': {self.metadata.get('description', '')}"

    def to_document(self, score: float, match: str) -> Document:
        return Document(
            page_content=self.text,
            metadata={**self.metadata, "node_id": self.node_id, "kind": self.kind,
                      "score": round(score, 4), "match": match},
        )


# ---------------------------------------------------------------------------
# Index construction
# ---------------------------------------------------------------------------
def _load_pure_component_nodes(excel_path: Path) -> list[IndexNode]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    df = pd.read_excel(excel_path)
    required = {"component_name", "full_name", "formula"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"components_list.xlsx is missing required columns: {missing}")

    nodes: list[IndexNode] = []
    for index, (name, group) in enumerate(df.groupby("component_name", sort=False)):
        name = str(name).strip()
        aliases = list(dict.fromkeys(str(v).strip() for v in group["full_name"] if str(v).strip()))
        formula = str(group["formula"].iloc[0]).strip()
        aspen_name = aspen_databank_name(name, aliases)
        text = (
            "Type: pure_component\n"
            f"HYSYS_name: {name}\n"
            f"Aspen_name: {aspen_name}\n"
            f"Full_name: {aliases[0] if aliases else name}\n"
            f"Formula: {formula}\n"
            f"Aliases: {', '.join([name, *aliases, formula])}\n"
        )
        nodes.append(
            IndexNode(
                node_id=f"C{index:04d}",
                kind="pure_component",
                title=name,
                keys=[name, formula, *aliases],
                text=text,
                metadata={"component_name": name, "aspen_name": aspen_name,
                          "formula": formula, "aliases": aliases},
            )
        )
    return nodes


def _field(block: str, name: str) -> str:
    match = re.search(rf"(?im)^\s*{re.escape(name)}\s*:\s*(.*?)\s*$", block)
    return match.group(1).strip() if match else ""


def _recipe_components(block: str) -> list[tuple[str, str]]:
    """Return [(component, fraction)] listed under HYSYS_components."""
    section = re.search(r"(?is)HYSYS_components:\s*\n(.*?)(?:\n\S|\Z)", block)
    if not section:
        return []
    return re.findall(r"-\s*(.+?)\s*\(([\d.]+)\)", section.group(1))


def _load_recipe_nodes(txt_path: Path, aspen_lookup: dict[str, str]) -> list[IndexNode]:
    if not txt_path.exists():
        raise FileNotFoundError(f"Mixture recipe file not found: {txt_path}")
    blocks = [b.strip() for b in txt_path.read_text(encoding="utf-8").split(_RECIPE_SEPARATOR) if b.strip()]

    nodes: list[IndexNode] = []
    for index, block in enumerate(blocks):
        mixture = _field(block, "Mixture_name") or f"UNKNOWN_MIXTURE_{index}"
        aliases_line = _field(block, "Aliases")
        aliases = [a.strip() for a in aliases_line.split(",") if a.strip()]
        components = _recipe_components(block)
        aspen_lines = "".join(
            f"  - {aspen_lookup.get(name, aspen_databank_name(name))} ({fraction})\n"
            for name, fraction in components
        )
        text = block + ("\nAspen_components:\n" + aspen_lines if aspen_lines else "")
        nodes.append(
            IndexNode(
                node_id=f"M{index:03d}",
                kind="pseudo_recipe",
                title=mixture,
                keys=[mixture, *aliases],
                text=text,
                metadata={"mixture_name": mixture, "aliases": aliases,
                          "description": _field(block, "Description"),
                          "components": [name for name, _ in components]},
            )
        )
    return nodes


class _BM25:
    def __init__(self, documents: list[list[str]], k1: float = 1.5, b: float = 0.75):
        self.k1, self.b = k1, b
        self.docs = [Counter(doc) for doc in documents]
        self.lengths = [len(doc) for doc in documents]
        self.avg_len = sum(self.lengths) / max(len(documents), 1)
        df: Counter = Counter()
        for doc in self.docs:
            df.update(doc.keys())
        n = len(documents)
        self.idf = {t: math.log(1 + (n - f + 0.5) / (f + 0.5)) for t, f in df.items()}
        self.postings: dict[str, list[int]] = {}
        for i, doc in enumerate(self.docs):
            for token in doc:
                self.postings.setdefault(token, []).append(i)

    def scores(self, query: list[str]) -> dict[int, float]:
        out: dict[int, float] = {}
        for token in set(query):
            idf = self.idf.get(token)
            if idf is None:
                continue
            for i in self.postings[token]:
                tf = self.docs[i][token]
                norm = tf + self.k1 * (1 - self.b + self.b * self.lengths[i] / self.avg_len)
                out[i] = out.get(i, 0.0) + idf * tf * (self.k1 + 1) / norm
        return out


class VectorlessIndex:
    """Hierarchical catalog index: root -> {mixture recipes, pure components}."""

    def __init__(self, nodes: list[IndexNode]):
        if not nodes:
            raise ValueError("No documents loaded for the vectorless index.")
        self.nodes = nodes
        self.by_id = {node.node_id: node for node in nodes}
        self.tree = {
            "pseudo_recipe": [n.node_id for n in nodes if n.kind == "pseudo_recipe"],
            "pure_component": [n.node_id for n in nodes if n.kind == "pure_component"],
        }
        self.exact: dict[str, list[int]] = {}
        for i, node in enumerate(nodes):
            for key in node.keys:
                k = normalize(key)
                if k and i not in self.exact.get(k, []):
                    self.exact.setdefault(k, []).append(i)
        self.key_list = list(self.exact)
        self.bm25 = _BM25([tokenize(" ".join(node.keys)) for node in nodes])

    def lexical_search(self, query: str, top_k: int = 8) -> list[tuple[IndexNode, float, str]]:
        """Rank nodes by exact > fuzzy > BM25. Returns (node, score, match_type)."""
        q = clean_query(query)
        if not q:
            return []
        q = normalize(SYNONYMS.get(q, q))
        scored: dict[int, tuple[float, str]] = {}

        def offer(i: int, score: float, match: str) -> None:
            if score > scored.get(i, (-1.0, ""))[0]:
                scored[i] = (score, match)

        for i in self.exact.get(q, []):
            # Mixture recipes win ties: a "crude oil" feed needs the recipe.
            offer(i, 3.0 + (0.1 if self.nodes[i].kind == "pseudo_recipe" else 0.0), "exact")

        for key in get_close_matches(q, self.key_list, n=10, cutoff=0.75):
            ratio = SequenceMatcher(None, q, key).ratio()
            for i in self.exact[key]:
                offer(i, 1.0 + ratio, "fuzzy")

        bm25 = self.bm25.scores(tokenize(q))
        if bm25:
            best = max(bm25.values())
            for i, s in bm25.items():
                offer(i, s / best, "bm25")

        ranked = sorted(scored.items(), key=lambda item: (-item[1][0], item[0]))
        return [(self.nodes[i], score, match) for i, (score, match) in ranked[:top_k]]


@lru_cache(maxsize=4)
def build_index(excel_path: str | Path = DEFAULT_EXCEL_PATH,
                mixture_path: str | Path = DEFAULT_MIXTURE_PATH) -> VectorlessIndex:
    pure = _load_pure_component_nodes(Path(excel_path))
    aspen_lookup = {n.title: n.metadata["aspen_name"] for n in pure}
    recipes = _load_recipe_nodes(Path(mixture_path), aspen_lookup)
    return VectorlessIndex(recipes + pure)


# ---------------------------------------------------------------------------
# Reasoning stage
# ---------------------------------------------------------------------------
ReasoningFn = Callable[[str], str]

_REASONING_PROMPT = """You select catalog entries for a chemical process simulator.

Feed material from a process diagram: "{query}"

Candidate entries (id: summary):
{toc}

Pick the entries that represent this material in a simulation. Prefer a mixture
recipe when the material is a blend (crude oil, fuel, natural gas, solutions).
Prefer the single pure component when the material is one chemical.
Reply with JSON only: {{"selected": ["<id>", ...], "reason": "<one sentence>"}}
Select at most {max_select} ids, best first. Use only ids from the list."""


def ollama_reasoner(model_name: str) -> ReasoningFn:
    """Reasoning function backed by a local/cloud Ollama chat model."""
    from langchain_ollama import ChatOllama

    llm = ChatOllama(model=model_name, temperature=0, format="json", num_predict=300)

    def reason(prompt: str) -> str:
        return str(llm.invoke(prompt).content or "")

    return reason


class VectorlessRetriever:
    """
    Drop-in replacement for the old Chroma retriever: `.invoke(query)` returns a
    list of langchain Documents whose page_content is the catalog entry.
    """

    def __init__(self, index: VectorlessIndex, reasoner: ReasoningFn | None = None,
                 top_k: int = 2, candidates: int = 8):
        self.index = index
        self.reasoner = reasoner
        self.top_k = top_k
        self.candidates = candidates
        self.last_trace: dict = {}

    def _reason(self, query: str, ranked: list[tuple[IndexNode, float, str]]) -> list[str] | None:
        toc = "\n".join(f"{node.node_id}: {node.summary()}" for node, _, _ in ranked)
        prompt = _REASONING_PROMPT.format(query=query, toc=toc, max_select=self.top_k)
        raw = self.reasoner(prompt)
        match = re.search(r"\{.*\}", raw or "", flags=re.DOTALL)
        if not match:
            raise ValueError(f"no JSON in reasoning output: {raw!r}")
        data = json.loads(match.group(0))
        allowed = {node.node_id for node, _, _ in ranked}
        selected = [i for i in data.get("selected", []) if i in allowed]
        self.last_trace["reason"] = data.get("reason", "")
        return selected or None

    def invoke(self, query: str, **_kwargs) -> list[Document]:
        ranked = self.index.lexical_search(query, top_k=self.candidates)
        self.last_trace = {"query": query, "lexical": [(n.node_id, n.title, round(s, 3), m) for n, s, m in ranked]}
        if not ranked:
            return []

        # A single exact hit needs no reasoning step.
        exact = [r for r in ranked if r[2] == "exact"]
        if self.reasoner is not None and len(exact) != 1:
            try:
                selected = self._reason(query, ranked)
            except Exception as exc:  # LLM down or bad output: keep lexical order
                self.last_trace["reasoning_error"] = str(exc)
                selected = None
            if selected:
                by_id = {node.node_id: (node, score, match) for node, score, match in ranked}
                self.last_trace["selected"] = selected
                return [by_id[i][0].to_document(by_id[i][1], "reasoning") for i in selected[: self.top_k]]

        return [node.to_document(score, match) for node, score, match in ranked[: self.top_k]]


def build_component_retriever(
    excel_path: str | Path = DEFAULT_EXCEL_PATH,
    mixture_path: str | Path = DEFAULT_MIXTURE_PATH,
    reasoning_model: str | None = None,
    reasoner: ReasoningFn | None = None,
    top_k: int = 2,
) -> VectorlessRetriever:
    """
    Build the vectorless retriever. Pass `reasoning_model` (an Ollama model name)
    or a custom `reasoner` callable to enable the LLM reasoning stage.
    """
    index = build_index(str(excel_path), str(mixture_path))
    if reasoner is None and reasoning_model:
        reasoner = ollama_reasoner(reasoning_model)
    return VectorlessRetriever(index, reasoner=reasoner, top_k=top_k)
