"""Step 3: Mask PII with reversible token replacement (fully local)."""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

from presidio_analyzer import RecognizerResult

logger = logging.getLogger(__name__)


@dataclass
class MaskResult:
    masked_text: str = ""
    # Maps token → original value, e.g. {"[PERSON_1]": "John Smith"}
    entity_mapping: dict[str, str] = field(default_factory=dict)
    # Reverse: original value → token
    reverse_mapping: dict[str, str] = field(default_factory=dict)
    entities_found: list[dict] = field(default_factory=list)


def mask_text(text: str, analyzer_results: list[RecognizerResult]) -> MaskResult:
    """Replace detected PII with consistent tokens like [PERSON_1], [LOCATION_2], etc."""
    result = MaskResult()

    # Counter per entity type for sequential numbering
    type_counters: dict[str, int] = {}
    # Cache: original value → token (for consistency across the document)
    value_to_token: dict[str, str] = {}

    # Process entities from end to start so string indices stay valid
    sorted_results = sorted(analyzer_results, key=lambda r: r.start, reverse=True)

    masked = text
    for entity in sorted_results:
        original_value = text[entity.start:entity.end]

        # Reuse token if we've seen this exact value before
        if original_value in value_to_token:
            token = value_to_token[original_value]
        else:
            entity_type = entity.entity_type
            type_counters.setdefault(entity_type, 0)
            type_counters[entity_type] += 1
            token = f"[{entity_type}_{type_counters[entity_type]}]"

            value_to_token[original_value] = token
            result.entity_mapping[token] = original_value
            result.reverse_mapping[original_value] = token

        masked = masked[:entity.start] + token + masked[entity.end:]

        result.entities_found.append({
            "entity_type": entity.entity_type,
            "original": original_value,
            "token": token,
            "score": round(entity.score, 2),
        })

    result.masked_text = masked
    logger.info(
        "Masked %d entities (%d unique values)",
        len(sorted_results),
        len(result.entity_mapping),
    )
    return result
