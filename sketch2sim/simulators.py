"""
Simulator profiles.

Every simulator template shares the same Step 1–7 markers and the same helper
API (add_unit, add_material_stream, connect_inlet, connect_outlet), so the
Instantiation, Configuration and Fixer agents and their instruction files are
simulator-agnostic. Only the Basis step and the template differ.
"""

from dataclasses import dataclass
from pathlib import Path

INSTRUCTIONS_DIR = Path(__file__).resolve().parent / "instructions"


@dataclass(frozen=True)
class SimulatorProfile:
    key: str
    display_name: str
    template_path: Path
    component_field: str  # KB field holding the name the template expects
    recipe_field: str  # KB field holding mixture components
    property_package_field: str  # template line the Basis agent fills
    property_package_examples: str
    script_suffix: str


SIMULATORS: dict[str, SimulatorProfile] = {
    "hysys": SimulatorProfile(
        key="hysys",
        display_name="Aspen HYSYS",
        template_path=INSTRUCTIONS_DIR / "hysys_template_base.py",
        component_field="HYSYS_name",
        recipe_field="HYSYS_components",
        property_package_field="fluidpkg.PropertyPackageName",
        property_package_examples='"Peng-Robinson", "SRK", "NRTL", "UNIQUAC", "ASME Steam"',
        script_suffix="hysys",
    ),
    "aspenplus": SimulatorProfile(
        key="aspenplus",
        display_name="Aspen Plus",
        template_path=INSTRUCTIONS_DIR / "aspenplus_template_base.py",
        component_field="Aspen_name",
        recipe_field="Aspen_components",
        property_package_field="property_method",
        property_package_examples='"PENG-ROB", "RK-SOAVE", "NRTL", "UNIQUAC", "STEAM-TA"',
        script_suffix="aspenplus",
    ),
}

ALIASES = {"aspen plus": "aspenplus", "aspen_plus": "aspenplus", "aspen hysys": "hysys"}


def get_simulator(name: str) -> SimulatorProfile:
    key = ALIASES.get(name.strip().lower(), name.strip().lower())
    if key not in SIMULATORS:
        raise ValueError(f"Unknown simulator '{name}'. Choose one of {sorted(SIMULATORS)}")
    return SIMULATORS[key]
