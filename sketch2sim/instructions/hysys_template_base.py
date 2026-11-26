import os
import re

from win32com.client import Dispatch

# ---------------------------------------------------------------------------
# HYSYS helper layer (fixed — agents must NOT edit this section).
# Agents call add_unit / add_material_stream / connect_inlet / connect_outlet
# instead of raw COM calls. Unit kinds and port names live in the tables below.
# Verify the type names against your HYSYS version with
# sketch2sim/instructions/verify_hysys_op_types.py
# ---------------------------------------------------------------------------
HYSYS_OP_TYPES = {
    "mixer": ["mixerop"],
    "tee": ["teeop"],
    "pump": ["pumpop"],
    "compressor": ["compressor"],
    "expander": ["expandop"],
    "valve": ["valveop"],
    "heater": ["heaterop"],
    "cooler": ["coolerop"],
    "heat_exchanger": ["heatexop"],
    "separator": ["flashtank"],
    "three_phase_separator": ["sep3op"],
    "tank": ["tankop"],
    "conversion_reactor": ["convreactorop", "conversionreactorop"],
    "cstr": ["cstrop"],
    "pfr": ["pfreactorop"],
    "gibbs_reactor": ["gibbsreactorop"],
    "equilibrium_reactor": ["eqmreactorop", "equilibriumreactorop"],
    "component_splitter": ["compsplitop", "compsplitter"],
    "distillation_column": ["distillation"],
    "absorber": ["absorber"],
}

# Kinds that get an energy stream "Q-<unit name>" attached automatically.
_ENERGY_KINDS = {"pump", "compressor", "expander", "heater", "cooler"}

# Kinds whose COM wiring is version-dependent: failures are reported, not fatal.
_BEST_EFFORT_KINDS = {"distillation_column", "absorber"}

# "Name+" = collection port (call .Add), "Name" = single-valued port.
_VESSEL_OUT = {
    "overhead": ["VapourProduct"],
    "bottom": ["LiquidProduct"],
    "*": ["VapourProduct", "LiquidProduct"],
}
_SINGLE_OUT = {"*": ["ProductStream"]}

_IN_PORTS = {
    "mixer": ["Feeds+"],
    "tee": ["Feed", "FeedStream"],
    "heat_exchanger": ["TubeSideFeed", "ShellSideFeed"],
    "pump": ["FeedStream"],
    "compressor": ["FeedStream"],
    "expander": ["FeedStream"],
    "valve": ["FeedStream"],
    "heater": ["FeedStream"],
    "cooler": ["FeedStream"],
    "pfr": ["Feeds+"],
}
_DEFAULT_IN = ["Feeds+", "FeedStream", "Feed"]

_OUT_PORTS = {
    "mixer": {"*": ["Product", "ProductStream"]},
    "tee": {"*": ["Products+"]},
    "heat_exchanger": {"*": ["TubeSideProduct", "ShellSideProduct"]},
    "pump": _SINGLE_OUT,
    "compressor": _SINGLE_OUT,
    "expander": _SINGLE_OUT,
    "valve": _SINGLE_OUT,
    "heater": _SINGLE_OUT,
    "cooler": _SINGLE_OUT,
    "pfr": {"*": ["ProductStream", "Product"]},
    "separator": _VESSEL_OUT,
    "tank": _VESSEL_OUT,
    "conversion_reactor": _VESSEL_OUT,
    "cstr": _VESSEL_OUT,
    "gibbs_reactor": _VESSEL_OUT,
    "equilibrium_reactor": _VESSEL_OUT,
    "three_phase_separator": {
        "overhead": ["VapourProduct"],
        "side": ["LightLiquidProduct"],
        "bottom": ["HeavyLiquidProduct"],
        "*": ["VapourProduct", "LightLiquidProduct", "HeavyLiquidProduct"],
    },
}
_DEFAULT_OUT = {
    "overhead": ["VapourProduct", "OverheadProduct", "DistillateStream"],
    "bottom": ["LiquidProduct", "BottomsProduct", "BottomsStream"],
    "*": ["Products+", "ProductStream", "Product", "VapourProduct", "LiquidProduct"],
}

_UNIT_KIND = {}  # HYSYS unit name -> kind
_USED_PORTS = set()  # (unit name, port) already filled


def _name_of(obj):
    try:
        return str(obj.name)
    except Exception:
        return str(obj)


def add_unit(name, kind):
    """Create a unit operation. kind must be a key of HYSYS_OP_TYPES."""
    kind = str(kind).strip().lower()
    if kind not in HYSYS_OP_TYPES:
        raise ValueError(
            f"Unknown unit kind '{kind}' for {name}. Use one of {sorted(HYSYS_OP_TYPES)}"
        )
    errors = []
    for type_name in HYSYS_OP_TYPES[kind]:
        try:
            op = flowsheet.Operations.Add(name, type_name)
            break
        except Exception as exc:
            errors.append(f"{type_name}: {exc}")
    else:
        raise RuntimeError(f"Could not create {name} as '{kind}': {errors}")
    _UNIT_KIND[name] = kind
    if kind in _ENERGY_KINDS:
        try:
            op.EnergyStream = flowsheet.EnergyStreams.Add(f"Q-{name}")
        except Exception as exc:
            print(f"[warn] Could not attach energy stream to {name}: {exc}")
    return op


def add_material_stream(name):
    return flowsheet.MaterialStreams.Add(name)


def _try_port(op, unit_name, port, stream):
    if port.endswith("+"):
        collection = getattr(op, port[:-1])
        try:
            collection.Add(stream)
        except Exception:
            collection.Add(_name_of(stream))
        return True
    if (unit_name, port) in _USED_PORTS:
        return False
    setattr(op, port, stream)
    _USED_PORTS.add((unit_name, port))
    return True


def _connect(op, stream, ports, direction):
    unit_name = _name_of(op)
    errors = []
    for port in ports:
        try:
            if _try_port(op, unit_name, port, stream):
                return port
            errors.append(f"{port}: already used")
        except Exception as exc:
            errors.append(f"{port}: {exc}")
    message = (
        f"Could not connect stream {_name_of(stream)} as {direction} of {unit_name} "
        f"(kind={_UNIT_KIND.get(unit_name)}). Tried: {errors}"
    )
    if _UNIT_KIND.get(unit_name) in _BEST_EFFORT_KINDS:
        print(f"[MANUAL ACTION NEEDED] {message}")
        return None
    raise RuntimeError(message)


def connect_inlet(op, stream):
    """Connect stream as an inlet of op."""
    ports = _IN_PORTS.get(_UNIT_KIND.get(_name_of(op)), _DEFAULT_IN)
    return _connect(op, stream, ports, "inlet")


def connect_outlet(op, stream, connect_point="unknown"):
    """Connect stream as an outlet of op. connect_point: overhead|bottom|side|unknown."""
    table = _OUT_PORTS.get(_UNIT_KIND.get(_name_of(op)), _DEFAULT_OUT)
    point = str(connect_point or "unknown").strip().lower()
    ports = table.get(point) or table["*"]
    if point in table:
        ports = ports + [p for p in table["*"] if p not in ports]
    return _connect(op, stream, ports, "outlet")


# === Step 1: Launch HYSYS and create a new case (Agent 1) ===
case_name = "FILL_CASE_NAME"  # Basis Agent must fill
hysys = Dispatch("HYSYS.Application")
hysys.Visible = True
case = hysys.SimulationCases.Add(case_name)
case.Visible = True

# === Step 2: Setup Property Package (Agent 1) ===
hybasis = case.BasisManager
fluidpkg = hybasis.FluidPackages.Add()
fluidpkg.PropertyPackageName = "FILL_PROPERTY_PACKAGE"  # Basis Agent must fill

# === Step 3: Add Components ===
components = [
    "COMPONENT_1",
    "COMPONENT_2",
]  # Basis Agent must fill the component list

# --- Component replacement blocks ---
# Add as many replacement blocks as needed.
if "COMPONENT_1" in components:
    components.remove("COMPONENT_1")
    components.extend(["HYSYS_COMP_1", "HYSYS_COMP_2"])  # Basis Agent must fill

if "COMPONENT_2" in components:
    components.remove("COMPONENT_2")
    components.extend(["HYSYS_COMP_1", "HYSYS_COMP_2"])  # Basis Agent must fill

for comp in components:
    fluidpkg.Components.Add(comp)
hybasis.EndBasisChange()
flowsheet = case.Flowsheet

# === Step 4: Add Unit Operations (Agent 2) ===
units = {}  # Instantiation Agent must fill, for example: "P-101" or "Pump-1"
# Instantiation Agent must add each unit operation in units according to the instructions file.

# === Step 5: Add Material Streams (Agent 2) ===
streams = {}  # Instantiation Agent must fill; key = stream id, HYSYS name = stream id
# Instantiation Agent must add each material stream in streams according to the instructions file.

# === Step 6: Connect Streams (Agent 3) ===
# Configuration Agent must connect streams to units here according to the instructions file.
# Connections must be grouped per unit.

# === Step 7: Save Simulation ===
base_dir = os.path.dirname(os.path.abspath(__file__))

safe_case_name = re.sub(r'[\\/*?:"<>|]', "_", case_name)
save_filename = safe_case_name + ".hsc"
save_path = os.path.join(base_dir, save_filename)
case.SaveAs(save_path)
print(f"Simulation saved at: {save_path}")

# === End of Template ===
