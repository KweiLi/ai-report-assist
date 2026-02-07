"""Step 4: Send masked text to OpenAI for debiasing."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field

from openai import OpenAI

from app.config import OPENAI_API_KEY, OPENAI_MODEL

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """\
You are an expert reviewer of law enforcement reports, specializing in identifying \
and removing implicit bias while preserving factual accuracy.

Your task is to perform an exhaustive, line-by-line review of the report. You must:
1. Read every sentence carefully and identify ALL instances of biased language.
2. Classify each instance by its bias type (you may assign multiple types if applicable).
3. Provide a neutral, factual replacement for each biased phrase.
4. Explain why each phrase is biased.
5. Produce the complete debiased report with all changes applied.

Be thorough. Do NOT skip subtle bias. Common examples in law enforcement reports include:
- Describing someone as "suspicious" or "nervous" without behavioural evidence
- Using words like "admitted", "claimed", "alleged" asymmetrically
- Characterising neighbourhoods as "high-crime", "known for drugs", etc.
- Describing physical appearance, clothing, or manner in a way that implies threat
- Using passive voice selectively to obscure agency (e.g. "the suspect was shot")
- Assuming intent or motive without stated evidence
- Labelling people ("transient", "vagrant", "gang member") without factual basis

Bias types — use ONLY these exact values:
- RACIAL_ETHNIC: Language suggesting racial, ethnic, or demographic profiling, \
including unnecessary mention of race/ethnicity when not relevant to identification
- GENDER: Gender-based stereotypes, assumptions, or unnecessarily gendered language
- SOCIOECONOMIC: Assumptions based on neighbourhood, employment, housing, \
clothing, appearance, or economic status
- STEREOTYPING: Generalizations about groups, communities, or locations
- INFLAMMATORY: Prejudicial, emotionally charged, or loaded word choices \
that imply guilt or threat without evidence
- CONFIRMATION: Selective emphasis of facts that supports a pre-formed conclusion \
while omitting mitigating context
- SUBJECTIVE: Opinions, characterizations, or judgments not directly supported \
by the stated observable facts

Rules:
- Preserve ALL factual content (dates, times, actions, sequences of events).
- Preserve ALL entity tokens exactly as they appear (e.g. [PERSON_1], [LOCATION_2]).
  Do NOT rename, remove, or alter any token.
- Maintain the original report structure and format.
- Do NOT add information that is not in the original report.
- Every change in debiased_text MUST have a corresponding entry in the changes array.
- The "original_phrase" must be the EXACT text from the original report.

Return your response as valid JSON with this exact structure:
{
  "debiased_text": "The full rewritten report with all bias removed...",
  "changes": [
    {
      "original_phrase": "the exact biased phrase copied from the original",
      "replacement_phrase": "the neutral replacement as it appears in debiased_text",
      "bias_type": "ONE_OF_THE_SEVEN_TYPES_ABOVE",
      "explanation": "Why this is biased and how the replacement fixes it"
    }
  ]
}

If no bias is found, return the original text unchanged with an empty changes array.\
"""

# Maps bias types to display colours
BIAS_COLORS: dict[str, str] = {
    "RACIAL_ETHNIC": "#e74c3c",     # red
    "GENDER": "#e67e22",            # orange
    "SOCIOECONOMIC": "#9b59b6",     # purple
    "STEREOTYPING": "#f39c12",      # amber
    "INFLAMMATORY": "#c0392b",      # dark red
    "CONFIRMATION": "#3498db",      # blue
    "SUBJECTIVE": "#e84393",        # pink
}


@dataclass
class BiasChange:
    original_phrase: str = ""
    replacement_phrase: str = ""
    bias_type: str = ""
    explanation: str = ""


@dataclass
class DebiasResult:
    debiased_text: str = ""
    changes: list[BiasChange] = field(default_factory=list)
    changes_summary: str = ""


def debias_text(masked_text: str) -> DebiasResult:
    """Send masked text to OpenAI for debiasing. Returns structured bias analysis."""
    if not OPENAI_API_KEY:
        raise RuntimeError("OPENAI_API_KEY is not set. Add it to your .env file.")

    client = OpenAI(api_key=OPENAI_API_KEY)

    logger.info("Sending masked text to OpenAI (%s) for debiasing", OPENAI_MODEL)

    response = client.chat.completions.create(
        model=OPENAI_MODEL,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": masked_text},
        ],
        temperature=0.3,
        response_format={"type": "json_object"},
    )

    raw = response.choices[0].message.content or "{}"
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        logger.error("Failed to parse OpenAI JSON response")
        return DebiasResult(debiased_text=masked_text)

    debiased_text = data.get("debiased_text", masked_text)
    raw_changes = data.get("changes", [])

    changes: list[BiasChange] = []
    summary_lines: list[str] = []
    for c in raw_changes:
        change = BiasChange(
            original_phrase=c.get("original_phrase", ""),
            replacement_phrase=c.get("replacement_phrase", ""),
            bias_type=c.get("bias_type", "SUBJECTIVE"),
            explanation=c.get("explanation", ""),
        )
        changes.append(change)
        summary_lines.append(
            f"- [{change.bias_type}] \"{change.original_phrase}\" -> "
            f"\"{change.replacement_phrase}\": {change.explanation}"
        )

    logger.info("Debiasing found %d biased phrases", len(changes))

    return DebiasResult(
        debiased_text=debiased_text,
        changes=changes,
        changes_summary="\n".join(summary_lines),
    )


def highlight_original(text: str, changes: list[BiasChange]) -> str:
    """Return HTML with biased phrases color-coded by bias type in the original text."""
    if not changes:
        return _escape(text)

    # Sort changes by position in text (longest match first to avoid partial overlaps)
    located: list[tuple[int, int, BiasChange]] = []
    search_from = 0
    for change in sorted(changes, key=lambda c: text.find(c.original_phrase)):
        idx = text.find(change.original_phrase, search_from)
        if idx >= 0:
            located.append((idx, idx + len(change.original_phrase), change))
            search_from = idx + 1

    # Sort by start position descending to build HTML from end
    located.sort(key=lambda x: x[0], reverse=True)

    html = text
    for start, end, change in located:
        color = BIAS_COLORS.get(change.bias_type, "#999")
        phrase = _escape(html[start:end])
        tooltip = _escape(f"[{change.bias_type}] {change.explanation}")
        span = (
            f'<span class="bias-highlight" '
            f'style="background-color: {color}22; border-bottom: 2px solid {color}; '
            f'cursor: help;" '
            f'title="{tooltip}" data-bias-type="{change.bias_type}">'
            f'{phrase}</span>'
        )
        html = html[:start] + span + html[end:]

    return html


def highlight_debiased(text: str, changes: list[BiasChange]) -> str:
    """Return HTML with replacement phrases highlighted in green in the debiased text."""
    if not changes:
        return _escape(text)

    located: list[tuple[int, int, BiasChange]] = []
    search_from = 0
    for change in sorted(changes, key=lambda c: text.find(c.replacement_phrase)):
        if not change.replacement_phrase:
            continue
        idx = text.find(change.replacement_phrase, search_from)
        if idx >= 0:
            located.append((idx, idx + len(change.replacement_phrase), change))
            search_from = idx + 1

    located.sort(key=lambda x: x[0], reverse=True)

    html = text
    for start, end, change in located:
        phrase = _escape(html[start:end])
        tooltip = _escape(f"Was: \"{change.original_phrase}\" [{change.bias_type}] {change.explanation}")
        span = (
            f'<span class="debias-highlight" '
            f'style="background-color: #2ecc7122; border-bottom: 2px solid #27ae60; '
            f'cursor: help;" '
            f'title="{tooltip}">'
            f'{phrase}</span>'
        )
        html = html[:start] + span + html[end:]

    return html


def _escape(text: str) -> str:
    """Escape HTML special characters."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )
