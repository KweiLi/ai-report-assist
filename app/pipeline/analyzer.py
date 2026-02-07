"""Step 2: PII and sensitive information detection using Presidio (fully local)."""

from __future__ import annotations

import logging
import re

from presidio_analyzer import AnalyzerEngine, RecognizerResult
from presidio_analyzer.nlp_engine import NlpEngineProvider

from dataclasses import dataclass, field

from app.config import PII_CONFIDENCE_THRESHOLD, SPACY_MODEL
from app.recognizers.law_enforcement import LAW_ENFORCEMENT_ACRONYMS, get_all_law_enforcement_recognizers

logger = logging.getLogger(__name__)

_engine: AnalyzerEngine | None = None


def _get_engine() -> AnalyzerEngine:
    """Lazily initialise the Presidio analyzer with spaCy + custom recognizers."""
    global _engine
    if _engine is not None:
        return _engine

    logger.info("Initializing Presidio analyzer with spaCy model: %s", SPACY_MODEL)

    nlp_config = {
        "nlp_engine_name": "spacy",
        "models": [{"lang_code": "en", "model_name": SPACY_MODEL}],
    }
    nlp_engine = NlpEngineProvider(nlp_configuration=nlp_config).create_engine()

    _engine = AnalyzerEngine(nlp_engine=nlp_engine, supported_languages=["en"])

    for recognizer in get_all_law_enforcement_recognizers():
        _engine.registry.add_recognizer(recognizer)
        logger.info("Registered custom recognizer: %s", recognizer.name)

    return _engine


def _remove_overlaps(results: list[RecognizerResult]) -> list[RecognizerResult]:
    """Keep only the highest-scoring entity when spans overlap."""
    if not results:
        return []
    # Sort by score descending so we keep the best match first
    by_score = sorted(results, key=lambda r: r.score, reverse=True)
    kept: list[RecognizerResult] = []
    for candidate in by_score:
        overlaps = any(
            candidate.start < k.end and candidate.end > k.start
            for k in kept
        )
        if not overlaps:
            kept.append(candidate)
    kept.sort(key=lambda r: r.start)
    return kept


@dataclass
class AnalysisResult:
    entities: list[RecognizerResult] = field(default_factory=list)
    acronyms_preserved: list[dict] = field(default_factory=list)


def analyze_text(text: str) -> AnalysisResult:
    """Detect PII entities in text. Filters out known acronyms/abbreviations."""
    engine = _get_engine()
    results = engine.analyze(
        text=text,
        language="en",
        score_threshold=PII_CONFIDENCE_THRESHOLD,
    )
    results = _remove_overlaps(results)

    # Classify each entity: real PII, known acronym (preserve), or unknown acronym (reclassify)
    pii_entities: list[RecognizerResult] = []
    acronyms_preserved: list[dict] = []
    seen_acronyms: set[str] = set()

    for r in results:
        matched_text = text[r.start:r.end].strip()

        # 1) Military time (e.g. @1715HRS, 0830HRS, @2300) → not PII, preserve
        if _is_military_time(matched_text):
            acronyms_preserved.append({
                "text": matched_text,
                "detected_as": r.entity_type,
                "reason": "Military/24-hour time notation - preserved",
            })
            logger.debug("Preserved military time: %s (detected as %s)", matched_text, r.entity_type)
            continue

        # 2) Known law enforcement acronym → preserve entirely (don't mask)
        if matched_text.upper() in LAW_ENFORCEMENT_ACRONYMS:
            if matched_text.upper() not in seen_acronyms:
                acronyms_preserved.append({
                    "text": matched_text,
                    "detected_as": r.entity_type,
                    "reason": "Known law enforcement acronym/abbreviation - preserved",
                })
                seen_acronyms.add(matched_text.upper())
            logger.debug("Preserved acronym: %s (detected as %s)", matched_text, r.entity_type)

        # 3) Short uppercase text that looks like an acronym → reclassify as ACRONYM
        elif _looks_like_acronym(matched_text):
            r.entity_type = "ACRONYM"
            pii_entities.append(r)
            logger.debug("Reclassified as ACRONYM: %s (was %s)", matched_text, r.entity_type)

        # 4) Everything else → real PII, keep as-is
        else:
            pii_entities.append(r)

    logger.info(
        "Found %d PII entities, preserved %d known acronyms",
        len(pii_entities), len(acronyms_preserved),
    )
    return AnalysisResult(entities=pii_entities, acronyms_preserved=acronyms_preserved)


# Matches: @1715HRS, 1715HRS, @1715hrs, @0830, 2300HRS, @1715Hrs, etc.
_MILITARY_TIME_RE = re.compile(
    r"^@?\d{3,4}\s?(?:HRS|hrs|Hrs|H)?$"
)


def _is_military_time(text: str) -> bool:
    """Detect military/24-hour time notation common in police reports."""
    return bool(_MILITARY_TIME_RE.match(text.strip()))


MAX_ACRONYM_LENGTH = 6


def _looks_like_acronym(text: str) -> bool:
    """Heuristic: short, mostly uppercase with optional digits → likely an acronym."""
    text = text.strip()
    if len(text) > MAX_ACRONYM_LENGTH or len(text) < 2:
        return False
    # Must be all uppercase letters and/or digits (e.g. S1, SP, K9, 10-4)
    alpha_chars = [c for c in text if c.isalpha()]
    if not alpha_chars:
        return False
    return all(c.isupper() or c.isdigit() or c in "-/" for c in text)
