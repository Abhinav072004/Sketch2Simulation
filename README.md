# Sketch2Simulation

[![Lint](https://github.com/Abhinav072004/Sketch2Simulation/actions/workflows/lint.yml/badge.svg)](https://github.com/Abhinav072004/Sketch2Simulation/actions/workflows/lint.yml)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![arXiv](https://img.shields.io/badge/arXiv-2603.24629-b31b1b.svg)](https://arxiv.org/abs/2603.24629)

Sketch2Simulation converts a process sketch image into an executable **Aspen HYSYS** or **Aspen Plus** Python script through a multi-agent pipeline.

📄 **Paper:** [Sketch2Simulation: Automating Flowsheet Generation via Multi Agent Large Language Models](https://arxiv.org/abs/2603.24629) (arXiv:2603.24629)

> **Prerequisites:** Executing the generated script (Step 9) needs a licensed installation of **Aspen HYSYS** or **Aspen Plus** with COM access enabled. Steps 1–8 and the test suite run without a simulator.

### What's new in this fork

| Change | Summary |
|--------|---------|
| **Aspen Plus support** | Pick the target with `SIMULATOR = "hysys"` or `"aspenplus"`. Both templates share the same Step markers and helper API, so every agent and instruction file works for both simulators. |
| **Vectorless RAG** | Component lookup no longer uses Chroma, embeddings or a vector store. Retrieval runs over a structured catalog index (exact → fuzzy → BM25), followed by an optional LLM reasoning step. `torch`, `chromadb`, `sentence-transformers` and the HuggingFace packages are gone from the requirements. |
| **Instruction files included** | `instantiation_instructions_simple_v5.txt` and `configuration_instructions_simple_v5.txt`, which the notebook requires, now ship with the repo. |
| **Helper layer** | Agents emit one-line calls (`add_unit`, `connect_inlet`, `connect_outlet`) instead of raw COM port assignments. Port mapping lives in tested Python. |
| **Tests + CI** | `pytest` suite with a fake `win32com` that executes the generated HYSYS and Aspen Plus scripts end to end. Runs in GitHub Actions. |
| **Fixes** | Package imports, RAG paths, configuration prompt, executor interpreter, absolute save path, `pywin32` platform marker. |

---

## Pipeline Overview

| Step | Agent / Module | Description |
|------|---------------|-------------|
| 1 | Descriptor | Describes the process from the input diagram image |
| 2 | Description Validator | Checks the description for repetition, inconsistency, and nonsense |
| 3 | Extractor | Converts the description into an intermediate structured representation |
| 4A | Mixing Points Normalizer | Resolves multi-input stream merges |
| 4B | Coupling Normalizer | Resolves stream coupling across unit operations |
| 5 | Basis Agent + Vectorless RAG | Case setup, property package/method, and components for the selected simulator |
| 6 | Instantiation Agent | Adds unit operations and material streams |
| 7 | Unit View Normalization | Builds a unit IN/OUT view for stream connection |
| 8 | Configuration Agent | Connects streams to unit operations |
| 9 | Merge + Execute | Merges the generated script, executes it, and applies runtime-driven fixing retries |

---

## Supported Simulators

| | Aspen HYSYS | Aspen Plus |
|---|---|---|
| Template | `instructions/hysys_template_base.py` | `instructions/aspenplus_template_base.py` |
| COM ProgID | `HYSYS.Application` | `Apwn.Document` |
| Output file | `<case>.hsc` | `<case>.bkp` |
| Component names | catalog `HYSYS_name` | catalog `Aspen_name` (databank name) |
| Property package | e.g. `Peng-Robinson` | e.g. `PENG-ROB` (HYSYS-style names are mapped automatically) |
| Check your install | `python sketch2sim/instructions/verify_hysys_op_types.py` | `python sketch2sim/instructions/verify_aspenplus_setup.py` |

Both templates expose the same helpers, so Steps 4–6 of a generated script look the same for either target:

```python
units["P-101"] = add_unit("P-101", "pump")
streams["S1"] = add_material_stream("S1")
connect_inlet(units["P-101"], streams["S1"])
connect_outlet(units["D-101"], streams["S7"], "overhead")
```

Unit kinds (`pump`, `valve`, `mixer`, `separator`, `three_phase_separator`, `heat_exchanger`, `distillation_column`, …) map to HYSYS operation types (`pumpop`, `flashtank`, …) or Aspen Plus blocks (`PUMP`, `FLASH2`, `RADFRAC`, …). The helpers choose the right port from the unit kind and the stream's `connect_point`. For example, a separator's overhead goes to `VapourProduct` in HYSYS and `V(OUT)` in Aspen Plus.

**Aspen Plus notes**
- Aspen Plus IDs are shortened to 8 characters, uniquely (`MIXER_TO_V-101` → `MIXER_TO`). The mapping is printed when the script runs. Change the limit with `ASPEN_MAX_ID_LEN`.
- A new case starts with `InitNew2()`. To start from your own blank file instead, set `ASPEN_PLUS_TEMPLATE=path\to\blank.apt` (or `.bkp`).
- Set `ASPEN_RUN=1` to run the engine before saving.
- Aspen databank names are derived from the catalog aliases (`Carbon_Dioxide` → `CARBON-DIOXIDE`, `i-Butane` → `ISOBUTANE`). Check unusual components in Aspen Plus.

> The COM type names and port names in both templates are best-effort and version dependent. Run the matching `verify_*` script once on a licensed machine and adjust the tables at the top of the template if needed. Column wiring (`distillation_column`, `absorber`) is reported as `[MANUAL ACTION NEEDED]` instead of failing the run when a port is missing.

---

## Vectorless RAG

`sketch2sim/RAG/vectorless_rag.py` replaces the Chroma vector store used by the Basis agent.

1. **Structured index.** The ~5,700 catalog rows are grouped into ~1,840 component nodes, one per simulator name with all its aliases and formula, plus one node per mixture recipe. Each node carries both `HYSYS_name` and `Aspen_name`. Recipes also list `Aspen_components`.
2. **Lexical retrieval, no embeddings:**
   - exact match on any key
   - fuzzy match (`difflib`) for typos (`Tolune` → Toluene)
   - BM25 over name tokens
   - stream-description words like *fresh*, *feed* and *raw* are stripped
   - a small synonym table (*steam* → water)
3. **Reasoning selection.** When the lexical stage is ambiguous, the LLM sees the top candidates as a compact table of contents and returns the ids that fit. A single exact hit skips this step. If the LLM is unavailable or answers badly, the lexical ranking is used.

```python
from sketch2sim.RAG.vectorless_rag import build_component_retriever

retriever = build_component_retriever(reasoning_model="qwen2.5-coder:latest")  # or None
docs = retriever.invoke("Fresh water")
print(docs[0].page_content)      # HYSYS_name: H2O / Aspen_name: WATER / ...
print(retriever.last_trace)      # lexical candidates, selected ids, reason
```

The Basis agent uses its own model for the reasoning step. Set `VECTORLESS_RAG_REASONING=0` to use only the lexical stage. Nothing has to be built or persisted on disk.

---

## Project Structure

```text
.
├── run_workflow.ipynb              # Main orchestrator notebook (set SIMULATOR here)
├── requirements.txt
├── requirements-dev.txt            # + pytest, ruff
├── images/                         # Input process sketch images (cs1–cs4)
├── tests/
│   ├── fakes/win32com/client.py    # Fake HYSYS + Aspen Plus COM for tests
│   ├── test_simulator_pipeline.py  # End-to-end wiring, both simulators
│   └── test_vectorless_rag.py
└── sketch2sim/
    ├── simulators.py               # Simulator profiles (hysys, aspenplus)
    ├── agents/                     # descriptor, validator, extractor, basis,
    │                               # instantiation, configuration, fixer, executor
    ├── mixing_points_normalizer.py
    ├── coupling_normalizer.py
    ├── merge_simulation_scripts.py
    ├── utils/
    ├── RAG/
    │   ├── vectorless_rag.py       # Vectorless retriever
    │   ├── rag_components.py       # Backward-compatible import shim
    │   ├── components_list.xlsx
    │   └── mixture_recipes.txt
    └── instructions/
        ├── hysys_template_base.py
        ├── aspenplus_template_base.py
        ├── instantiation_instructions_simple_v5.txt
        ├── configuration_instructions_simple_v5.txt
        ├── verify_hysys_op_types.py
        └── verify_aspenplus_setup.py
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt          # runtime
pip install -r requirements-dev.txt      # + tests and lint
```

### 2. Configure environment variables

```bash
export OLLAMA_API_KEY=your_api_key_here
export OLLAMA_CLOUD_HOST=your_cloud_endpoint_here

# Optional
export SKETCH2SIM_SIMULATOR=aspenplus    # default: hysys
export VECTORLESS_RAG_REASONING=0        # lexical retrieval only
export ASPEN_PLUS_TEMPLATE=C:/path/blank.apt
export DESCRIPTOR_MAX_IMAGE_WIDTH=2048
export DESCRIPTOR_NUM_PREDICT=5000
export DESCRIPTOR_NUM_CTX=21000
export DESCRIPTOR_TIMEOUT=900
export EXTRACTOR_TIMEOUT=900
export EXTRACTOR_NUM_CTX=21000
```

### 3. Make the simulator available

Install a licensed Aspen HYSYS or Aspen Plus with COM access enabled. The generated script talks to it via `pywin32`, which is installed on Windows only. Then run the matching `verify_*` script once.

---

## Models Used

| Agent | Model |
|-------|-------|
| Descriptor | `gemini-3-flash-preview:cloud` |
| Description Validator | `qwen3-vl:8b-instruct-q4_K_M` |
| Extractor | `gemini-3-flash-preview:cloud` |
| Basis Agent (+ RAG reasoning) | `qwen2.5-coder:latest` |
| Instantiation Agent | `qwen2.5-coder:latest` |
| Configuration Agent | `qwen3-coder:30b-a3b-q4_K_M` |
| Fixer Agent | configured via orchestrator |

---

## How to Run

1. Place your process sketch in `images/` and update `IMAGE_PATH` in `run_workflow.ipynb` if needed (default: `cs1.png`).
2. Choose the simulator: set `SKETCH2SIM_SIMULATOR`, or edit `SIMULATOR` in the notebook.
3. Open `run_workflow.ipynb` in Jupyter and run all cells.

The pipeline writes:
- Intermediate logs to `logs/`
- Generated Python scripts to `outputs/`
- Final merged script to `outputs/<simulator>_combined_latest.py`, for example `outputs/aspenplus_combined_latest.py`

### Tests

```bash
python -m pytest -q
```

The tests run the real normalizers and the agents' extract/merge code, then execute the merged HYSYS and Aspen Plus scripts against `tests/fakes/win32com`. No LLM or simulator licence is needed.

---

## Notes

- The code-generation stages rely on the Step marker comments shared by both templates. If you change them, update the regex-based extraction and replacement logic too.
- Do not edit the helper section above `# === Step 1:` from an agent. It survives every regex replacement by design.
- `logs/` and `outputs/` are generated at runtime and should not be committed.

---

## Status

Research/prototype workflow. Needs a valid HYSYS or Aspen Plus licence and model access to run end to end.

---

## Citation

If you use Sketch2Simulation in your research, please cite:

```bibtex
@misc{bahamdan2026sketch2simulationautomatingflowsheetgeneration,
      title={Sketch2Simulation: Automating Flowsheet Generation via Multi Agent Large Language Models},
      author={Abdullah Bahamdan and Emma Pajak and John D. Hedengren and Antonio del Rio Chanona},
      year={2026},
      eprint={2603.24629},
      archivePrefix={arXiv},
      primaryClass={cs.SE},
      url={https://arxiv.org/abs/2603.24629},
}
```
