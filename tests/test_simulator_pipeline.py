"""
End-to-end wiring test for both simulators, without an LLM or a simulator.

A "perfect LLM" writes Steps 1–6 exactly as the instruction files describe. The
output goes through the agents' real extract/merge functions and the merged
script is executed against the fake win32com in tests/fakes.
"""

import copy
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from sketch2sim.agents import basis_agent, configuration_agent, instantiation_agent
from sketch2sim.coupling_normalizer import couple_extraction
from sketch2sim.mixing_points_normalizer import normalize_mixing_points
from sketch2sim.simulators import get_simulator
from sketch2sim.utils.normalization_utils import build_unit_view_from_structure

FAKES = Path(__file__).parent / "fakes"

CS1 = {  # images/cs1.png: desalter
    "units": [
        {"id": "P-101", "name": "Crude oil pump", "tags": ["pump"]},
        {"id": "P-102", "name": "Fresh water pump", "tags": ["pump"]},
        {"id": "V-101", "name": "Mixing valve", "tags": ["valve"]},
        {"id": "D-101", "name": "Electrostatic separator", "tags": ["separator"]},
    ],
    "feed_streams": [
        {"id": "S1", "to_unit": "P-101", "name": "Crude oil", "connect_point": "unknown"},
        {"id": "S2", "to_unit": "V-101", "name": "Demulsifier agent", "connect_point": "unknown"},
        {"id": "S3", "to_unit": "P-102", "name": "Fresh water", "connect_point": "unknown"},
    ],
    "intermediate_streams": [
        {"id": "S4", "from_unit": "P-101", "to_unit": "V-101", "name": "", "connect_point": "unknown"},
        {"id": "S5", "from_unit": "P-102", "to_unit": "V-101", "name": "", "connect_point": "unknown"},
        {"id": "S6", "from_unit": "V-101", "to_unit": "D-101", "name": "", "connect_point": "unknown"},
    ],
    "product_streams": [
        {"id": "S7", "from_unit": "D-101", "name": "Desalted crude oil", "connect_point": "overhead"},
        {"id": "S8", "from_unit": "D-101", "name": "Effluent water", "connect_point": "bottom"},
    ],
}

CS3 = {  # column with condenser + reboiler that the coupling normalizer folds in
    "units": [
        {"id": "E-101", "name": "Pre-heat train", "tags": ["heater"]},
        {"id": "T-101", "name": "Atmospheric distillation column", "tags": []},
        {"id": "C-101", "name": "Condenser", "tags": []},
        {"id": "R-101", "name": "Reboiler", "tags": []},
    ],
    "feed_streams": [{"id": "F1", "to_unit": "E-101", "name": "Crude", "connect_point": "unknown"}],
    "intermediate_streams": [
        {"id": "F2", "from_unit": "E-101", "to_unit": "T-101", "name": "", "connect_point": "side"},
        {"id": "F3", "from_unit": "T-101", "to_unit": "C-101", "name": "", "connect_point": "overhead"},
        {"id": "F4", "from_unit": "T-101", "to_unit": "R-101", "name": "", "connect_point": "bottom"},
    ],
    "product_streams": [
        {"id": "NAPH", "from_unit": "C-101", "name": "Naphtha", "connect_point": "unknown"},
        {"id": "KERO", "from_unit": "T-101", "name": "Kero", "connect_point": "side"},
        {"id": "RES", "from_unit": "R-101", "name": "Residue", "connect_point": "bottom"},
    ],
}

_KINDS = [("column", "distillation_column"), ("pump", "pump"), ("valve", "valve"),
          ("heater", "heater"), ("separator", "separator")]


def _kind(unit):
    if unit["id"].startswith("MIXER_TO_"):
        return "mixer"
    text = " ".join([unit["id"], unit.get("name", ""), *unit.get("tags", [])]).lower()
    return next((kind for word, kind in _KINDS if word in text), "separator")


def _fake_basis_output(template, simulator):
    """What a well-behaved Basis LLM returns: the template with Steps 1–3 filled."""
    if simulator == "aspenplus":
        filled = (template.replace('"FILL_PROPERTY_PACKAGE"', '"Peng-Robinson"')
                  .replace('["ASPEN_COMP_1", "ASPEN_COMP_2"]', '["WATER", "N-HEPTANE"]'))
    else:
        filled = (template.replace('"FILL_PROPERTY_PACKAGE"', '"Peng-Robinson"')
                  .replace('["HYSYS_COMP_1", "HYSYS_COMP_2"]', '["H2O", "n-Heptane"]'))
    return "```python\n" + filled.replace('"FILL_CASE_NAME"', '"TestCase"') + "\n```"


def _fake_instantiation_output(structure):
    lines = ["# === Step 4: Add Unit Operations (Agent 2) ===", "units = {}"]
    lines += [f'units["{u["id"]}"] = add_unit("{u["id"]}", "{_kind(u)}")' for u in structure["units"]]
    lines += ["", "# === Step 5: Add Material Streams (Agent 2) ===", "streams = {}"]
    for group in ("feed_streams", "intermediate_streams", "product_streams"):
        lines += [f'streams["{s["id"]}"] = add_material_stream("{s["id"]}")' for s in structure[group]]
    return "```python\n" + "\n".join(lines) + "\n```"


def _fake_configuration_output(unit_view):
    lines = ["# === Step 6: Connect Streams (Agent 3) ==="]
    for unit in unit_view["units_unit_view"]:
        uid = unit["unit_id"]
        lines.append(f"# --- {uid} ---")
        lines += [f'connect_inlet(units["{uid}"], streams["{s["id"]}"])' for s in unit["streams_in"]]
        lines += [
            f'connect_outlet(units["{uid}"], streams["{s["id"]}"], "{s["connect_point"] or "unknown"}")'
            for s in unit["streams_out"]
        ]
    return "```python\n" + "\n".join(lines) + "\n\n# === Step 7: Save Simulation ===\n```"


def run_pipeline(tmp_path, simulator, extraction):
    profile = get_simulator(simulator)
    structure = couple_extraction(normalize_mixing_points(copy.deepcopy(extraction)))
    unit_view = build_unit_view_from_structure(structure)

    template = profile.template_path.read_text(encoding="utf-8")
    code = basis_agent._apply_safety_belt(
        template, basis_agent._extract_code(_fake_basis_output(template, simulator))
    )
    step45 = instantiation_agent._extract_steps_4_to_5_only(
        instantiation_agent._extract_code(_fake_instantiation_output(structure))
    )
    code = instantiation_agent._replace_steps_4_to_5(code, step45)
    code = configuration_agent._replace_step_6(
        code, configuration_agent._extract_code(_fake_configuration_output(unit_view))
    )

    script = tmp_path / f"{simulator}_combined.py"
    script.write_text(code, encoding="utf-8")
    log = tmp_path / "com_log.json"
    env = {**os.environ, "PYTHONPATH": str(FAKES), "FAKE_COM_LOG": str(log)}
    env.pop("ASPEN_PLUS_TEMPLATE", None)
    result = subprocess.run([sys.executable, str(script)], capture_output=True, text=True, env=env)
    assert result.returncode == 0, result.stderr
    return json.loads(log.read_text(encoding="utf-8")), result.stdout


def _pairs(log):
    return {(unit, port, stream) for unit, port, stream in log["connections"]}


def test_hysys_cs1(tmp_path):
    log, _ = run_pipeline(tmp_path, "hysys", CS1)
    pairs = _pairs(log)
    assert log["property"] == "Peng-Robinson"
    assert "H2O" in log["components"]
    # mixing normalizer inserted a mixer in front of the valve
    assert {("MIXER_TO_V-101", "Feeds", s) for s in ("S2", "S4", "S5")} <= pairs
    assert ("MIXER_TO_V-101", "Product", "S_MIXER_TO_V-101") in pairs
    assert ("V-101", "FeedStream", "S_MIXER_TO_V-101") in pairs
    assert ("D-101", "VapourProduct", "S7") in pairs
    assert ("D-101", "LiquidProduct", "S8") in pairs
    assert log["saved"].endswith("TestCase.hsc")


def test_aspenplus_cs1(tmp_path):
    log, stdout = run_pipeline(tmp_path, "aspenplus", CS1)
    pairs = _pairs(log)
    assert log["property"] == "PENG-ROB"  # HYSYS-style name mapped
    assert ["WATER", "WATER"] in log["components"]
    # long IDs are shortened to 8 characters but stay unique
    assert {("MIXER_TO", "F(IN)", s) for s in ("S2", "S4", "S5")} <= pairs
    assert ("MIXER_TO", "P(OUT)", "S_MIXER_") in pairs
    assert ("V-101", "F(IN)", "S_MIXER_") in pairs
    assert ("D-101", "V(OUT)", "S7") in pairs
    assert ("D-101", "L(OUT)", "S8") in pairs
    assert "MIXER_TO_V-101->MIXER_TO" in stdout
    assert log["saved"].endswith("TestCase.bkp")


def test_aspenplus_column_ports(tmp_path):
    log, stdout = run_pipeline(tmp_path, "aspenplus", CS3)
    pairs = _pairs(log)
    assert ("T-101", "F(IN)", "F2") in pairs
    assert ("T-101", "SP(OUT)", "KERO") in pairs
    assert ("T-101", "B(OUT)", "RES") in pairs
    assert "MANUAL ACTION NEEDED" not in stdout


def test_hysys_column_is_best_effort(tmp_path):
    log, stdout = run_pipeline(tmp_path, "hysys", CS3)
    assert "[MANUAL ACTION NEEDED]" in stdout  # column wiring reported, not fatal
    assert ("E-101", "FeedStream", "F1") in _pairs(log)
    assert log["saved"]


@pytest.mark.parametrize("simulator", ["hysys", "aspenplus"])
def test_templates_share_markers(simulator):
    text = get_simulator(simulator).template_path.read_text(encoding="utf-8")
    for marker in (
        "# === Step 1:",
        "# === Step 4: Add Unit Operations (Agent 2) ===",
        "# === Step 5: Add Material Streams (Agent 2) ===",
        "# === Step 6: Connect Streams (Agent 3) ===",
        "# === Step 7: Save Simulation ===",
    ):
        assert marker in text


@pytest.mark.parametrize("simulator, field", [("hysys", "HYSYS_name"), ("aspenplus", "Aspen_name")])
def test_basis_prompt_uses_simulator_field(simulator, field):
    prompt = basis_agent._system_prompt(get_simulator(simulator))
    assert field in prompt


def test_unknown_simulator():
    with pytest.raises(ValueError):
        get_simulator("dwsim")
    assert get_simulator("Aspen Plus").key == "aspenplus"
