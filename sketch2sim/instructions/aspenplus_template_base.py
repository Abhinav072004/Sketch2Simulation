import os
import re

from win32com.client import Dispatch

# ---------------------------------------------------------------------------
# Aspen Plus helper layer (fixed — agents must NOT edit this section).
# Same API as the HYSYS template: add_unit / add_material_stream /
# connect_inlet / connect_outlet, plus set_property_method / add_component.
# Verify block types and port names against your Aspen Plus version with
# sketch2sim/instructions/verify_aspenplus_setup.py
# ---------------------------------------------------------------------------
ASPEN_BLOCK_TYPES = {
    "mixer": "MIXER",
    "tee": "FSPLIT",
    "pump": "PUMP",
    "compressor": "COMPR",
    "expander": "COMPR",
    "valve": "VALVE",
    "heater": "HEATER",
    "cooler": "HEATER",
    "heat_exchanger": "HEATX",
    "separator": "FLASH2",
    "three_phase_separator": "FLASH3",
    "tank": "FLASH2",
    "conversion_reactor": "RSTOIC",
    "cstr": "RCSTR",
    "pfr": "RPLUG",
    "gibbs_reactor": "RGIBBS",
    "equilibrium_reactor": "REQUIL",
    "component_splitter": "SEP",
    "distillation_column": "RADFRAC",
    "absorber": "RADFRAC",
}

# Property method names accepted in HYSYS style and mapped to Aspen Plus.
PROPERTY_METHOD_ALIASES = {
    "PENGROBINSON": "PENG-ROB",
    "PENG-ROBINSON": "PENG-ROB",
    "PR": "PENG-ROB",
    "SRK": "RK-SOAVE",
    "SOAVEREDLICHKWONG": "RK-SOAVE",
    "ASMESTEAM": "STEAM-TA",
    "NBSSTEAM": "STEAM-TA",
    "GRAYSONSTREED": "GRAYSON",
}

# Kinds whose wiring needs stage numbers etc.: failures are reported, not fatal.
_BEST_EFFORT_KINDS = {"distillation_column", "absorber"}
ASPEN_MAX_ID_LEN = int(os.environ.get("ASPEN_MAX_ID_LEN", "8"))

# "PORT+" = port accepts several streams, "PORT" = one stream only.
_FEED = ["F(IN)+"]
_SINGLE_OUT = {"*": ["P(OUT)"]}
_FLASH_OUT = {"overhead": ["V(OUT)"], "bottom": ["L(OUT)"], "*": ["V(OUT)", "L(OUT)"]}
_REACTOR_OUT = {"overhead": ["V(OUT)"], "bottom": ["L(OUT)"], "*": ["P(OUT)", "V(OUT)", "L(OUT)"]}
_COLUMN_OUT = {
    "overhead": ["LD(OUT)", "VD(OUT)"],
    "bottom": ["B(OUT)"],
    "side": ["SP(OUT)+"],
    "*": ["LD(OUT)", "B(OUT)", "SP(OUT)+"],
}

_IN_PORTS = {"heat_exchanger": ["H(IN)", "C(IN)"]}
_OUT_PORTS = {
    "mixer": _SINGLE_OUT,
    "tee": {"*": ["P(OUT)+"]},
    "pump": _SINGLE_OUT,
    "compressor": _SINGLE_OUT,
    "expander": _SINGLE_OUT,
    "valve": _SINGLE_OUT,
    "heater": _SINGLE_OUT,
    "cooler": _SINGLE_OUT,
    "heat_exchanger": {"*": ["H(OUT)", "C(OUT)"]},
    "separator": _FLASH_OUT,
    "tank": _FLASH_OUT,
    "three_phase_separator": {
        "overhead": ["V(OUT)"],
        "side": ["L1(OUT)"],
        "bottom": ["L2(OUT)"],
        "*": ["V(OUT)", "L1(OUT)", "L2(OUT)"],
    },
    "conversion_reactor": _REACTOR_OUT,
    "cstr": _REACTOR_OUT,
    "pfr": _SINGLE_OUT,
    "gibbs_reactor": _REACTOR_OUT,
    "equilibrium_reactor": _REACTOR_OUT,
    "component_splitter": {"*": ["P(OUT)+"]},
    "distillation_column": _COLUMN_OUT,
    "absorber": _COLUMN_OUT,
}

_ASPEN_IDS = {}  # logical name -> Aspen ID
_USED_PORTS = set()  # (block id, port) already filled


class AspenObject:
    """Handle for a block or stream; .name is the logical (diagram) name."""

    def __init__(self, name, aspen_id, kind):
        self.name, self.aspen_id, self.kind = name, aspen_id, kind

    def __repr__(self):
        return f"<{self.kind} {self.name} as {self.aspen_id}>"


def aspen_id(name):
    """Map a diagram name to a unique, valid Aspen Plus ID."""
    if name in _ASPEN_IDS:
        return _ASPEN_IDS[name]
    base = re.sub(r"[^A-Z0-9\-_]", "", str(name).upper()) or "X"
    base = base[:ASPEN_MAX_ID_LEN]
    candidate, n = base, 2
    while candidate in _ASPEN_IDS.values():
        suffix = str(n)
        candidate = base[: ASPEN_MAX_ID_LEN - len(suffix)] + suffix
        n += 1
    _ASPEN_IDS[name] = candidate
    return candidate


def _data(node):
    return aspen.Tree.Elements("Data").Elements(node)


def set_property_method(name):
    key = re.sub(r"[\s_]", "", str(name).upper())
    method = PROPERTY_METHOD_ALIASES.get(key, str(name).strip().upper())
    for path in (
        r"\Data\Properties\Specifications\Input\GBASEOPSET",
        r"\Data\Properties\Specifications\Input\GOPSETNAME",
    ):
        aspen.Tree.FindNode(path).Value = method
    return method


_COMPONENT_IDS = {}


def add_component(databank_name):
    """Add a databank component (e.g. "WATER", "N-HEPTANE") with a short ID."""
    databank_name = str(databank_name).strip().upper()
    if databank_name in _COMPONENT_IDS:
        return _COMPONENT_IDS[databank_name]
    base =re.sub(r"[^A-Z0-9\-]", "", databank_name)[:8] or "COMP"
    cid, n = base, 2
    while cid in _COMPONENT_IDS.values():
        cid = base[: 8 - len(str(n))] + str(n)
        n += 1
    table = aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\TYPE")
    row = table.Elements.Count
    table.Elements.InsertRow(0, row)
    table.Elements.SetLabel(0, row, False, cid)
    aspen.Tree.FindNode(rf"\Data\Components\Specifications\Input\ANAME\{cid}").Value = databank_name
    _COMPONENT_IDS[databank_name] = cid
    return cid


def add_unit(name, kind):
    """Create a block. kind must be a key of ASPEN_BLOCK_TYPES."""
    kind = str(kind).strip().lower()
    if kind not in ASPEN_BLOCK_TYPES:
        raise ValueError(
            f"Unknown unit kind '{kind}' for {name}. Use one of {sorted(ASPEN_BLOCK_TYPES)}"
        )
    block_id = aspen_id(name)
    try:
        _data("Blocks").Elements.Add(f"{block_id}!{ASPEN_BLOCK_TYPES[kind]}")
    except Exception as exc:
        raise RuntimeError(f"Could not create {name} as '{kind}' ({ASPEN_BLOCK_TYPES[kind]}): {exc}")
    return AspenObject(name, block_id, kind)


def add_material_stream(name):
    stream_id = aspen_id(name)
    _data("Streams").Elements.Add(f"{stream_id}!MATERIAL")
    return AspenObject(name, stream_id, "stream")


def _try_port(unit, port, stream):
    multi = port.endswith("+")
    port = port.rstrip("+")
    if not multi and (unit.aspen_id, port) in _USED_PORTS:
        return False
    port_node = _data("Blocks").Elements(unit.aspen_id).Elements("Ports").Elements(port)
    port_node.Elements.Add(stream.aspen_id)
    _USED_PORTS.add((unit.aspen_id, port))
    return True


def _connect(unit, stream, ports, direction):
    errors = []
    for port in ports:
        try:
            if _try_port(unit, port, stream):
                return port.rstrip("+")
            errors.append(f"{port}: already used")
        except Exception as exc:
            errors.append(f"{port}: {exc}")
    message = (
        f"Could not connect stream {stream.name} as {direction} of {unit.name} "
        f"(kind={unit.kind}). Tried: {errors}"
    )
    if unit.kind in _BEST_EFFORT_KINDS:
        print(f"[MANUAL ACTION NEEDED] {message}")
        return None
    raise RuntimeError(message)


def connect_inlet(unit, stream):
    """Connect stream as an inlet of unit."""
    return _connect(unit, stream, _IN_PORTS.get(unit.kind, _FEED), "inlet")


def connect_outlet(unit, stream, connect_point="unknown"):
    """Connect stream as an outlet of unit. connect_point: overhead|bottom|side|unknown."""
    table = _OUT_PORTS.get(unit.kind, _SINGLE_OUT)
    point = str(connect_point or "unknown").strip().lower()
    ports = table.get(point) or table["*"]
    if point in table:
        ports = ports + [p for p in table["*"] if p not in ports]
    return _connect(unit, stream, ports, "outlet")


# === Step 1: Launch Aspen Plus and create a new case (Agent 1) ===
case_name = "FILL_CASE_NAME"  # Basis Agent must fill
aspen = Dispatch("Apwn.Document")
_start_file = os.environ.get("ASPEN_PLUS_TEMPLATE", "")
if _start_file.lower().endswith(".apt"):
    aspen.InitFromTemplate2(os.path.abspath(_start_file))
elif _start_file:
    aspen.InitFromArchive2(os.path.abspath(_start_file))
else:
    aspen.InitNew2()
aspen.Visible = True
aspen.SuppressDialogs = 1

# === Step 2: Setup Property Method (Agent 1) ===
property_method = "FILL_PROPERTY_PACKAGE"  # Basis Agent must fill
set_property_method(property_method)

# === Step 3: Add Components ===
components = [
    "COMPONENT_1",
    "COMPONENT_2",
]  # Basis Agent must fill the component list (Aspen databank names)

# --- Component replacement blocks ---
# Add as many replacement blocks as needed.
if "COMPONENT_1" in components:
    components.remove("COMPONENT_1")
    components.extend(["ASPEN_COMP_1", "ASPEN_COMP_2"])  # Basis Agent must fill

if "COMPONENT_2" in components:
    components.remove("COMPONENT_2")
    components.extend(["ASPEN_COMP_1", "ASPEN_COMP_2"])  # Basis Agent must fill

for comp in components:
    add_component(comp)

# === Step 4: Add Unit Operations (Agent 2) ===
units = {}  # Instantiation Agent must fill, for example: "P-101" or "Pump-1"
# Instantiation Agent must add each unit operation in units according to the instructions file.

# === Step 5: Add Material Streams (Agent 2) ===
streams = {}  # Instantiation Agent must fill; key = stream id
# Instantiation Agent must add each material stream in streams according to the instructions file.

# === Step 6: Connect Streams (Agent 3) ===
# Configuration Agent must connect streams to units here according to the instructions file.
# Connections must be grouped per unit.

# === Step 7: Save Simulation ===
base_dir = os.path.dirname(os.path.abspath(__file__))

safe_case_name = re.sub(r'[\\/*?:"<>|]', "_", case_name)
save_path = os.path.join(base_dir, safe_case_name + ".bkp")
if _ASPEN_IDS:
    print("Aspen Plus IDs:", ", ".join(f"{k}->{v}" for k, v in _ASPEN_IDS.items()))
if os.environ.get("ASPEN_RUN") == "1":
    aspen.Engine.Run2()
aspen.SaveAs(save_path)
print(f"Simulation saved at: {save_path}")

# === End of Template ===
