"""Step 4: Send masked text to OpenAI for debiasing."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field

from openai import OpenAI

from app.config import OPENAI_API_KEY, OPENAI_MODEL

logger = logging.getLogger(__name__)

SYSTEM_PROMPT = """\
You are an expert reviewer of law enforcement reports trained in identifying \
implicit and explicit bias. Your goal is MAXIMUM RECALL — find every instance \
of biased language. It is far worse to MISS a bias than to flag a borderline case.

## PROCESS

Perform a TWO-PASS analysis:

PASS 1 — SCAN: Read every sentence and check it against ALL bias categories below. \
For each sentence, ask: "Does this contain any loaded verbs, subjective characterizations, \
labels, area bias, passive voice hiding agency, presumption of guilt, appearance bias, \
selective emphasis, or unnecessary demographic details?"

PASS 2 — PRODUCE: Write the complete debiased report and the changes array. \
Every biased phrase you identified in Pass 1 MUST appear in the changes array.

## BIAS CATEGORIES AND SPECIFIC PATTERNS TO FLAG

Use ONLY these exact bias_type values. Check EVERY category for EVERY sentence.

### INFLAMMATORY
Prejudicial, emotionally charged, or loaded word choices that imply guilt or threat.

Flag these verbs and replace with neutral alternatives:
- "admitted" / "confessed" → "stated" / "said"
- "claimed" / "insisted" / "alleged" → "stated" / "reported"
- "fled" / "took off" → "left" / "departed" / "ran"
- "lurking" / "prowling" / "skulking" / "sneaking" → "standing" / "walking" / "present"
- "loitering" → "standing" / "present in the area"
- "brandished" / "wielded" → "held" / "displayed"
- "lunged" / "charged" → "moved toward" / "approached"
- "confronted" / "accosted" → "approached" / "contacted"
- "combative" / "belligerent" / "hostile" / "aggressive" → describe the specific actions
- "darting" → "moving quickly"
- "uncooperative" / "defiant" / "non-compliant" → describe the specific refusal
- "resisted" / "actively resisted" → describe specific physical actions observed
- "perpetrator" / "offender" (before conviction) → "individual" / "person"
- "the stolen [item]" (before confirmed stolen) → "the reported stolen [item]"

### SUBJECTIVE
Opinions, characterizations, or judgments not supported by stated observable facts.

Flag these and replace with specific observable behaviors:
- "suspicious" → describe what specific behavior was observed
- "nervous" → "was fidgeting" / "shifted weight between feet" (specific behavior)
- "agitated" / "erratic" / "irrational" → describe the specific behavior
- "evasive" → describe what the person specifically did
- "appeared intoxicated" (without objective signs) → list specific signs observed
- "acted strange" / "seemed angry" / "looked guilty" → describe observable actions
- "furtive movements" / "furtive gestures" → describe the specific movement
- "avoided eye contact" → state factually, note it is not evidence of deception
- "blading" / "tensed up" / "bracing" → describe the specific body position
- "reached for waistband" → describe exactly what movement was observed
- "changed story" / "inconsistent statements" → quote the specific inconsistency

### STEREOTYPING
Generalizations about groups, communities, locations, or individuals based on associations.

Flag these:
- "high-crime area" / "known drug corridor" / "high narcotics area" → use the specific address
- "known gang territory" / "dangerous neighborhood" → use the specific location name
- "blighted area" / "rough neighborhood" / "sketchy area" → use the specific location
- "gang member" / "gang affiliate" / "gang associate" (without evidence) → "individual"
- "known offender" / "repeat offender" / "career criminal" / "frequent flyer" → do not include
- "transient" / "vagrant" / "drifter" → "person" / "individual"
- "fits the description" (if vague) → provide the specific description elements
- "known to police" / "prior contact" → omit unless directly relevant to current incident
- "wearing gang colors" / "gang attire" → describe the actual clothing

### SOCIOECONOMIC
Assumptions based on neighborhood, housing, employment, clothing, appearance, or economic status.

Flag these:
- "unkempt" / "disheveled" / "slovenly" / "poorly groomed" / "dirty" → omit or describe \
factually only if relevant to identification
- "homeless person" → "person" (mention housing status only if directly relevant)
- "public housing" / "Section 8" / "government housing" / "low-income housing" → use the address
- References to clothing quality implying economic status
- "unemployed" mentioned when irrelevant to the incident
- Describing someone by their apparent economic status

### RACIAL_ETHNIC
Language suggesting racial, ethnic, or demographic profiling, or unnecessary mention \
of race/ethnicity.

Flag these:
- Race/ethnicity mentioned when NOT part of a specific suspect description from a victim/witness
- Race mentioned for suspects but not for officers or witnesses (asymmetric use)
- Immigration status mentioned when irrelevant ("illegal alien" → never appropriate)
- Racial descriptors given more detail for minorities than for white individuals
- Using ethnicity or national origin as part of justification for suspicion

### GENDER
Gender-based stereotypes, assumptions, or unnecessarily gendered language.

Flag these:
- Focus on what a victim was wearing (implying responsibility)
- Emphasis on victim's intoxication or sexual history
- "Delayed reporting" emphasized as credibility issue
- Questioning "lack of resistance" by a victim
- Gendered assumptions about aggressor/victim roles in domestic incidents

### CONFIRMATION
Selective emphasis of facts supporting a pre-formed conclusion while omitting context.

Flag these:
- Criminal history emphasized early or repeatedly when not relevant to current incident
- "Previously arrested for..." when not relevant
- Mentioning prior contacts with police to establish pattern when not relevant
- Emphasizing incriminating details while minimizing exculpatory information
- Using passive voice SELECTIVELY for officer actions but active voice for suspect actions \
(e.g., "the suspect struck the officer" but "force was applied" — flag the passive one)
- "Officer-involved shooting" → "[Officer name] discharged their firearm"
- "Suspect was struck" → "[Officer name] struck the individual"
- "Force was applied" → "[Officer name] used [specific force type]"
- "Shots were fired" → "[Officer name] fired [number] rounds"
- "In-custody death" → describe what actually occurred

## RULES

- Preserve ALL factual content (dates, times, locations, actions, sequences of events).
- Preserve ALL entity tokens exactly as they appear (e.g. [PERSON_1], [LOCATION_2]). \
Do NOT rename, remove, or alter any token in brackets.
- Maintain the original report structure, formatting, and paragraph breaks.
- Do NOT add information not in the original report.
- Every change in debiased_text MUST have a corresponding entry in changes array.
- The "original_phrase" must be the EXACT text copied from the original report.
- When in doubt, FLAG IT. It is better to over-identify than to miss bias.
- A single phrase can be flagged under multiple bias types — create separate entries for each.

## OUTPUT FORMAT

Return valid JSON:
{
  "debiased_text": "The full rewritten report with all bias removed...",
  "changes": [
    {
      "original_phrase": "exact biased text from the original",
      "replacement_phrase": "neutral replacement as it appears in debiased_text",
      "bias_type": "ONE_OF_THE_SEVEN_TYPES",
      "explanation": "Why this is biased and how the replacement corrects it"
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
        temperature=0,
        seed=42,
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
