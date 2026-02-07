"""Step 5: Unmask — restore original PII from the mapping (fully local)."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)

TOKEN_PATTERN = re.compile(r"\[[A-Z_]+_\d+\]")


@dataclass
class UnmaskResult:
    final_text: str = ""
    unresolved_tokens: list[str] = field(default_factory=list)


def unmask_text(masked_text: str, entity_mapping: dict[str, str]) -> UnmaskResult:
    """Replace all tokens back to their original values using the mapping."""
    result = UnmaskResult()
    text = masked_text

    # Replace each token with its original value
    for token, original in entity_mapping.items():
        text = text.replace(token, original)

    # Check for any remaining unresolved tokens
    remaining = TOKEN_PATTERN.findall(text)
    result.unresolved_tokens = remaining
    if remaining:
        logger.warning("Unresolved tokens after unmasking: %s", remaining)

    result.final_text = text
    return result
