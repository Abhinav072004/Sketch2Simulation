import json

import pytest

from sketch2sim.RAG.vectorless_rag import (
    aspen_databank_name,
    build_component_retriever,
    clean_query,
)


@pytest.fixture(scope="module")
def retriever():
    return build_component_retriever()


def first(docs):
    meta = docs[0].metadata
    return meta.get("component_name") or meta.get("mixture_name")


@pytest.mark.parametrize(
    "query, expected",
    [
        ("Crude oil", "Crude Oil"),  # mixture recipe, exact
        ("Fresh water", "H2O"),  # noise word stripped, alias "Water"
        ("steam", "H2O"),  # synonym table
        ("Carbon dioxide", "CO2"),  # underscore alias Carbon_Dioxide
        ("Tolune", "Toluene"),  # typo -> fuzzy
        ("Kerosene", "Jet Fuel"),  # recipe alias
        ("Methanol feed", "Methanol"),
    ],
)
def test_lexical_retrieval(retriever, query, expected):
    docs = retriever.invoke(query)
    assert docs, query
    assert first(docs) == expected


def test_aliases_are_grouped_into_one_node(retriever):
    names = [n.title for n in retriever.index.nodes if n.kind == "pure_component"]
    assert len(names) == len(set(names))


def test_documents_carry_both_simulator_names(retriever):
    text = retriever.invoke("Water")[0].page_content
    assert "HYSYS_name: H2O" in text
    assert "Aspen_name: WATER" in text


def test_recipe_has_aspen_components(retriever):
    text = retriever.invoke("Crude oil")[0].page_content
    assert "HYSYS_components:" in text
    assert "Aspen_components:" in text
    assert "N-HEPTANE (0.25)" in text


def test_empty_and_unknown_queries(retriever):
    assert retriever.invoke("") == []
    assert retriever.invoke("zzqqxx") == []


def test_reasoning_stage_selects_candidate():
    calls = []

    def reasoner(prompt):
        calls.append(prompt)
        ids = [line.split(":")[0] for line in prompt.splitlines() if line[:1] in "CM" and ":" in line]
        return json.dumps({"selected": [ids[-1]], "reason": "picked last"})

    r = build_component_retriever(reasoner=reasoner, top_k=1)
    docs = r.invoke("Natural gas")  # two fuzzy recipes -> reasoning runs
    assert calls and len(docs) == 1
    assert docs[0].metadata["match"] == "reasoning"
    assert r.last_trace["reason"] == "picked last"


def test_reasoning_skipped_for_single_exact_hit():
    def reasoner(prompt):
        raise AssertionError("should not be called")

    docs = build_component_retriever(reasoner=reasoner).invoke("Toluene")
    assert first(docs) == "Toluene"


@pytest.mark.parametrize("bad_output", ["not json", '{"selected": ["NOPE"]}', None])
def test_reasoning_failure_falls_back_to_lexical(bad_output):
    def reasoner(prompt):
        if bad_output is None:
            raise ConnectionError("ollama down")
        return bad_output

    r = build_component_retriever(reasoner=reasoner)
    docs = r.invoke("Natural gas")
    assert docs and docs[0].metadata["match"] != "reasoning"


@pytest.mark.parametrize(
    "name, aliases, expected",
    [
        ("Toluene", [], "TOLUENE"),
        ("p-Xylene", [], "P-XYLENE"),
        ("i-Butane", ["i-C4"], "ISOBUTANE"),
        ("H2O", ["Water"], "WATER"),
        ("NaOH", ["Sodium_Hydroxide"], "SODIUM-HYDROXIDE"),
        ("C2H6O2", ["Ethylene_Glycol"], "ETHYLENE-GLYCOL"),
    ],
)
def test_aspen_databank_name(name, aliases, expected):
    assert aspen_databank_name(name, aliases) == expected


def test_clean_query():
    assert clean_query("  Fresh   Water feed ") == "water"
