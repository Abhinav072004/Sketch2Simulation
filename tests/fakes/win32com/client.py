"""
Fake win32com.client used by the tests to execute generated simulator scripts
without Aspen HYSYS or Aspen Plus. Every port assignment is recorded and dumped
as JSON to the path in FAKE_COM_LOG when the case is saved.

The port tables mirror the names the templates expect, so the tests check the
pipeline wiring (regex merge + helper layer), not the real COM type libraries.
"""

import json
import os

LOG = {"connections": [], "components": [], "property": None, "saved": None, "warnings": []}


def _dump():
    path = os.environ.get("FAKE_COM_LOG")
    if path:
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(LOG, fh, indent=1)


class _Anything:
    def __getattr__(self, name):
        return _Anything()

    def __call__(self, *args, **kwargs):
        return _Anything()


# ----------------------------------------------------------------------------- HYSYS
HYSYS_PORTS = {
    "mixerop": {"Feeds+", "Product"},
    "teeop": {"Feed", "Products+"},
    "pumpop": {"FeedStream", "ProductStream", "EnergyStream"},
    "compressor": {"FeedStream", "ProductStream", "EnergyStream"},
    "expandop": {"FeedStream", "ProductStream", "EnergyStream"},
    "valveop": {"FeedStream", "ProductStream"},
    "heaterop": {"FeedStream", "ProductStream", "EnergyStream"},
    "coolerop": {"FeedStream", "ProductStream", "EnergyStream"},
    "heatexop": {"TubeSideFeed", "TubeSideProduct", "ShellSideFeed", "ShellSideProduct"},
    "flashtank": {"Feeds+", "VapourProduct", "LiquidProduct"},
    "sep3op": {"Feeds+", "VapourProduct", "LightLiquidProduct", "HeavyLiquidProduct"},
    "tankop": {"Feeds+", "VapourProduct", "LiquidProduct"},
    "distillation": set(),  # column wiring is version dependent -> best effort
}


class _HysysCollection:
    def __init__(self, owner, port):
        self.owner, self.port = owner, port

    def Add(self, item):
        LOG["connections"].append([self.owner, self.port, getattr(item, "name", item)])
        return item


class _HysysObject:
    def __init__(self, name):
        object.__setattr__(self, "name", name)


class _HysysOp(_HysysObject):
    def __init__(self, name, type_name):
        super().__init__(name)
        object.__setattr__(self, "_type", type_name)
        for port in HYSYS_PORTS[type_name]:
            if port.endswith("+"):
                object.__setattr__(self, port[:-1], _HysysCollection(name, port[:-1]))

    def __setattr__(self, key, value):
        if key not in HYSYS_PORTS[self._type]:
            raise AttributeError(f"{self._type} has no {key}")
        object.__setattr__(self, key, value)
        LOG["connections"].append([self.name, key, value.name])


class _HysysOperations:
    def Add(self, name, type_name):
        if type_name not in HYSYS_PORTS:
            raise Exception(f"unknown operation type {type_name}")
        return _HysysOp(name, type_name)


class _HysysStreams:
    def Add(self, name):
        return _HysysObject(name)


class _HysysComponents:
    def Add(self, name):
        LOG["components"].append(name)


class _FluidPackage:
    def __init__(self):
        self.Components = _HysysComponents()

    def __setattr__(self, key, value):
        object.__setattr__(self, key, value)
        if key == "PropertyPackageName":
            LOG["property"] = value


class _BasisManager:
    def __init__(self):
        self.FluidPackages = self

    def Add(self):
        return _FluidPackage()

    def EndBasisChange(self):
        pass


class _HysysFlowsheet:
    Operations = _HysysOperations()
    MaterialStreams = _HysysStreams()
    EnergyStreams = _HysysStreams()


class _HysysCase(_Anything):
    def __init__(self):
        self.Flowsheet = _HysysFlowsheet()
        self.BasisManager = _BasisManager()

    def SaveAs(self, path):
        LOG["saved"] = path
        _dump()


class _HysysApp(_Anything):
    class SimulationCases:
        @staticmethod
        def Add(name):
            return _HysysCase()


# ------------------------------------------------------------------------ Aspen Plus
ASPEN_PORTS = {
    "MIXER": {"F(IN)", "P(OUT)"},
    "FSPLIT": {"F(IN)", "P(OUT)"},
    "PUMP": {"F(IN)", "P(OUT)"},
    "COMPR": {"F(IN)", "P(OUT)"},
    "VALVE": {"F(IN)", "P(OUT)"},
    "HEATER": {"F(IN)", "P(OUT)"},
    "HEATX": {"H(IN)", "H(OUT)", "C(IN)", "C(OUT)"},
    "FLASH2": {"F(IN)", "V(OUT)", "L(OUT)"},
    "FLASH3": {"F(IN)", "V(OUT)", "L1(OUT)", "L2(OUT)"},
    "RSTOIC": {"F(IN)", "P(OUT)", "V(OUT)", "L(OUT)"},
    "RCSTR": {"F(IN)", "P(OUT)", "V(OUT)", "L(OUT)"},
    "RPLUG": {"F(IN)", "P(OUT)"},
    "RGIBBS": {"F(IN)", "P(OUT)", "V(OUT)", "L(OUT)"},
    "REQUIL": {"F(IN)", "P(OUT)", "V(OUT)", "L(OUT)"},
    "SEP": {"F(IN)", "P(OUT)"},
    "RADFRAC": {"F(IN)", "LD(OUT)", "VD(OUT)", "B(OUT)", "SP(OUT)"},
}


class _AspenElements:
    """Callable collection: node.Elements("X") and node.Elements.Add(...)."""

    def __init__(self, node):
        self.node = node
        self.children = {}
        self.rows = []

    def __call__(self, name):
        if name not in self.children:
            raise Exception(f"{self.node.path}: no element {name}")
        return self.children[name]

    @property
    def Count(self):
        return len(self.rows)

    def Add(self, spec):
        self.node.add(spec)

    def InsertRow(self, dim, row):
        self.rows.insert(row, None)

    def SetLabel(self, dim, row, force, label):
        self.rows[row] = label


class _AspenNode:
    def __init__(self, path, kind="folder", block_type=None):
        self.path, self.kind, self.block_type = path, kind, block_type
        self.Elements = _AspenElements(self)
        self.Value = None

    def child(self, name, **kw):
        node = self.Elements.children.get(name)
        if node is None:
            node = _AspenNode(f"{self.path}\\{name}", **kw)
            self.Elements.children[name] = node
        return node

    def add(self, spec):
        if self.kind == "blocks":
            block_id, block_type = spec.split("!")
            if block_type not in ASPEN_PORTS:
                raise Exception(f"unknown block type {block_type}")
            if len(block_id) > 8:
                raise Exception(f"ID too long: {block_id}")
            block = self.child(block_id, kind="block", block_type=block_type)
            ports = block.child("Ports")
            for port in ASPEN_PORTS[block_type]:
                ports.child(port, kind="port")
        elif self.kind == "streams":
            stream_id = spec.split("!")[0]
            if len(stream_id) > 8:
                raise Exception(f"ID too long: {stream_id}")
            self.child(stream_id, kind="stream")
        elif self.kind == "port":
            parts = self.path.split("\\")
            LOG["connections"].append([parts[-3], parts[-1], spec])
        else:
            raise Exception(f"cannot add to {self.path}")


class _AspenValueNode:
    def __init__(self, path):
        self.path = path

    def __setattr__(self, key, value):
        object.__setattr__(self, key, value)
        if key == "Value":
            if "ANAME" in self.path:
                LOG["components"].append([self.path.rsplit("\\", 1)[-1], value])
            elif "GBASEOPSET" in self.path:
                LOG["property"] = value


class _AspenTree:
    def __init__(self):
        root = _AspenNode("Root")
        data = root.child("Data")
        data.child("Blocks", kind="blocks")
        data.child("Streams", kind="streams")
        self.Elements = root.Elements
        self.component_table = _AspenNode("\\Data\\Components\\Specifications\\Input\\TYPE")

    def FindNode(self, path):
        if path.endswith("\\TYPE"):
            return self.component_table
        return _AspenValueNode(path)


class _AspenDocument(_Anything):
    def __init__(self):
        self.Tree = _AspenTree()

    def InitNew2(self):
        pass

    def SaveAs(self, path):
        LOG["saved"] = path
        _dump()


def Dispatch(prog_id):
    if prog_id == "HYSYS.Application":
        return _HysysApp()
    if prog_id == "Apwn.Document":
        return _AspenDocument()
    raise Exception(f"unknown ProgID {prog_id}")
