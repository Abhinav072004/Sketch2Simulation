"""
Check which HYSYS unit-operation type names in HYSYS_OP_TYPES work on this machine.

Run on a machine with a licensed Aspen HYSYS:
    python sketch2sim/instructions/verify_hysys_op_types.py

For each kind it tries every candidate type name, then lists the port-like
attributes of the created operation. If a kind reports FAIL, find the correct
type name for your HYSYS version (record a macro in HYSYS or check its
type library) and add it to HYSYS_OP_TYPES in hysys_template_base.py.
"""

import ast
from pathlib import Path

from win32com.client import Dispatch

TEMPLATE = Path(__file__).with_name("hysys_template_base.py")
PORT_NAMES = [
    "Feeds", "FeedStream", "Feed", "Product", "Products", "ProductStream",
    "VapourProduct", "LiquidProduct", "LightLiquidProduct", "HeavyLiquidProduct",
    "TubeSideFeed", "TubeSideProduct", "ShellSideFeed", "ShellSideProduct",
    "EnergyStream",
]


def load_op_types() -> dict[str, list[str]]:
    tree = ast.parse(TEMPLATE.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") == "HYSYS_OP_TYPES":
            return ast.literal_eval(node.value)
    raise RuntimeError("HYSYS_OP_TYPES not found in template")


def main() -> None:
    op_types = load_op_types()
    hysys = Dispatch("HYSYS.Application")
    case = hysys.SimulationCases.Add("op_type_check")
    basis = case.BasisManager
    pkg = basis.FluidPackages.Add()
    pkg.PropertyPackageName = "Peng-Robinson"
    pkg.Components.Add("Water")
    basis.EndBasisChange()
    flowsheet = case.Flowsheet

    for kind, candidates in op_types.items():
        for type_name in candidates:
            try:
                op = flowsheet.Operations.Add(f"{kind}_{type_name}"[:20], type_name)
            except Exception as exc:
                print(f"FAIL  {kind:24s} {type_name:24s} {exc}")
                continue
            ports = []
            for port in PORT_NAMES:
                try:
                    getattr(op, port)
                    ports.append(port)
                except Exception:
                    pass
            print(f"OK    {kind:24s} {type_name:24s} ports: {', '.join(ports)}")
            break

    case.Close(False)


if __name__ == "__main__":
    main()
