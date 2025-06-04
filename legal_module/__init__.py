"""Public API for the legal document generator."""

from .filing import create_legal_filing, FilingType

__all__ = ["create_legal_filing", "FilingType"]
