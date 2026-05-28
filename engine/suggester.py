"""Backward-compat shim — use engine.drafter instead."""
# PatternSuggester is exported as an alias so existing code and tests
# that import it directly continue to work without modification.
from .drafter import (  # noqa: F401
    ExcelAnalyzer,
    LlamaCppClient,
    PatternDrafter as PatternSuggester,
    PatternWriter,
    _SYSTEM_PROMPT,
    _USER_PROMPT_TEMPLATE,
)
