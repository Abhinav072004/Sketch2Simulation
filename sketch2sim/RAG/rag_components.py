"""
Backward-compatible entry point for component retrieval.

The Chroma + sentence-transformers vector store was replaced by the vectorless
retriever in vectorless_rag.py. `build_component_retriever` keeps its old name
and still returns an object with `.invoke(query) -> list[Document]`.
"""

from .vectorless_rag import (
    VectorlessIndex as VectorlessIndex,
)
from .vectorless_rag import (
    VectorlessRetriever as VectorlessRetriever,
)
from .vectorless_rag import (
    build_component_retriever as build_component_retriever,
)
from .vectorless_rag import (
    build_index as build_index,
)
