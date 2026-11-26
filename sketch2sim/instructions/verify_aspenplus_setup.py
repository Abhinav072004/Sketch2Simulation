"""
Check the Aspen Plus COM calls used by aspenplus_template_base.py on this machine.

Run on a machine with a licensed Aspen Plus:
    python sketch2sim/instructions/verify_aspenplus_setup.py

It creates a blank case, sets a property method, adds a component, creates one
block of every type in ASPEN_BLOCK_TYPES and prints the port names each block
really exposes. Compare them with _OUT_PORTS / _IN_PORTS in the template and
adjust the template if your version differs.

Set ASPEN_PLUS_TEMPLATE to a blank .apt/.bkp file if InitNew2 is unavailable.
"""

import ast
import os
from pathlib import Path

from win32com.client import Dispatch

TEMPLATE = Path(__file__).with_name("aspenplus_template_base.py")


def load_block_types() -> dict[str, str]:
    tree = ast.parse(TEMPLATE.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") == "ASPEN_BLOCK_TYPES":
            return ast.literal_eval(node.value)
    raise RuntimeError("ASPEN_BLOCK_TYPES not found in template")


def child_names(node) -> list[str]:
    names = []
    elements = node.Elements
    for i in range(elements.Count):
        try:
            names.append(str(elements.Item(i).Name))
        except Exception:
            pass
    return names


def check(label: str, fn) -> None:
    try:
        result = fn()
        print(f"OK    {label}" + (f": {result}" if result is not None else ""))
    except Exception as exc:
        print(f"FAIL  {label}: {exc}")


def main() -> None:
    aspen = Dispatch("Apwn.Document")
    start = os.environ.get("ASPEN_PLUS_TEMPLATE", "")
    if start.lower().endswith(".apt"):
        check("InitFromTemplate2", lambda: aspen.InitFromTemplate2(os.path.abspath(start)))
    elif start:
        check("InitFromArchive2", lambda: aspen.InitFromArchive2(os.path.abspath(start)))
    else:
        check("InitNew2", lambda: aspen.InitNew2())
    aspen.SuppressDialogs = 1

    def set_method():
        aspen.Tree.FindNode(r"\Data\Properties\Specifications\Input\GBASEOPSET").Value = "PENG-ROB"

    def add_water():
        table = aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\TYPE")
        table.Elements.InsertRow(0, 0)
        table.Elements.SetLabel(0, 0, False, "WATER")
        aspen.Tree.FindNode(r"\Data\Components\Specifications\Input\ANAME\WATER").Value = "WATER"

    check("property method GBASEOPSET", set_method)
    check("add component WATER", add_water)

    data = aspen.Tree.Elements("Data")
    check("add stream S1", lambda: data.Elements("Streams").Elements.Add("S1!MATERIAL"))

    seen = set()
    for kind, block_type in load_block_types().items():
        if block_type in seen:
            continue
        seen.add(block_type)
        block_id = f"B{len(seen)}"
        try:
            data.Elements("Blocks").Elements.Add(f"{block_id}!{block_type}")
            ports = data.Elements("Blocks").Elements(block_id).Elements("Ports")
            print(f"OK    {kind:22s} {block_type:8s} ports: {', '.join(child_names(ports))}")
        except Exception as exc:
            print(f"FAIL  {kind:22s} {block_type:8s} {exc}")

    check(
        "connect S1 to B1 F(IN)",
        lambda: data.Elements("Blocks").Elements("B1").Elements("Ports")
        .Elements("F(IN)").Elements.Add("S1"),
    )


if __name__ == "__main__":
    main()
