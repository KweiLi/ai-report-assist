"""Custom Presidio recognizers for law-enforcement-specific entities."""

from presidio_analyzer import Pattern, PatternRecognizer

# Known law enforcement acronyms and abbreviations that should NOT be masked.
# These are standard terminology, not PII.
LAW_ENFORCEMENT_ACRONYMS: set[str] = {
    # Agencies / Units
    "FBI", "DEA", "ATF", "ICE", "CBP", "DHS", "CIA", "NSA", "USMS",
    "SWAT", "SRT", "HRT", "K9", "CSI", "CSU", "IAB", "IAD", "IA",
    "CID", "SIU", "OIS", "PD", "SO", "LEA", "DOJ", "DOC", "DA",
    "ADA", "AG", "PO", "CO", "SGT", "LT", "CPT", "DET", "OFC",
    # Offenses / Charges
    "DUI", "DWI", "OVI", "OWI", "DWAI", "BUI",
    "GTA", "B&E", "AWOL", "UUMV",
    "CSA", "CSAM", "IPV", "DV",
    # Reports / Procedures
    "BOLO", "APB", "BOL",
    "FIR", "UCR", "NIBRS",
    "SOP", "ROE", "MOU", "TRO", "EPO", "PPO",
    "NCIC", "AFIS", "CODIS", "IAFIS", "NLETS",
    "RMS", "CAD", "MDT", "MDC",
    "PC", "RS", "TC", "RAS",
    "FTA", "GOA", "UTL", "AKA",
    # Medical / Scene
    "DOA", "DOB", "EMS", "EMT", "ER", "ED", "ME",
    "GSW", "OD", "BAC", "FST", "SFT", "HGN",
    "MVA", "MVC", "TC", "PI",
    "PPE", "AED", "CPR", "NARCAN",
    # Evidence / Forensics
    "DNA", "GSR", "SEM", "FTIR",
    "BWC", "ALPR", "LPR", "CCTV", "DVR", "NVR",
    # Legal / Court
    "RICO", "CCW", "CPL", "FOID", "NFA",
    "ROR", "PR", "OR", "PTI",
    "PSI", "VIS", "VINE",
    # Vehicle / Traffic
    "VIN", "DMV", "BMV", "MVD", "CDL",
    "RO", "RP", "POI", "POC",
    # Radio / Communication
    "ETA", "ETD", "EOW", "ASAP",
    "QRT", "TOD", "TOC",
    # Misc
    "AKA", "NKA", "NFI", "NFA", "GOA", "UTL",
    "AWOL", "MIA", "WMA", "WFA", "BMA", "BFA", "HMA", "HFA",
    "AMA", "AFA", "NMI",
}


def build_badge_number_recognizer() -> PatternRecognizer:
    return PatternRecognizer(
        supported_entity="BADGE_NUMBER",
        name="badge_number_recognizer",
        patterns=[
            Pattern("badge_alpha_num", r"\b[A-Z]{1,3}\d{4,7}\b", 0.6),
        ],
        context=["badge", "officer id", "shield", "shield no", "badge no", "badge number"],
        supported_language="en",
    )


def build_case_number_recognizer() -> PatternRecognizer:
    return PatternRecognizer(
        supported_entity="CASE_NUMBER",
        name="case_number_recognizer",
        patterns=[
            Pattern("case_dashed", r"\b\d{2,4}[-/]\d{2,6}[-/]?\d{0,6}\b", 0.6),
            Pattern("case_cr", r"\bCR[-]?\d{4,10}\b", 0.8),
        ],
        context=["case", "case no", "case number", "docket", "file no", "file number", "incident"],
        supported_language="en",
    )


def build_evidence_id_recognizer() -> PatternRecognizer:
    return PatternRecognizer(
        supported_entity="EVIDENCE_ID",
        name="evidence_id_recognizer",
        patterns=[
            Pattern("evidence_ev", r"\bEV[-]?\d{4,8}\b", 0.8),
            Pattern("evidence_exhibit", r"\bEXH[-]?\d{3,8}\b", 0.7),
        ],
        context=["evidence", "exhibit", "item", "evidence id", "evidence no"],
        supported_language="en",
    )


def build_license_plate_recognizer() -> PatternRecognizer:
    return PatternRecognizer(
        supported_entity="LICENSE_PLATE",
        name="license_plate_recognizer",
        patterns=[
            Pattern("plate_us_standard", r"\b[A-Z]{2,3}[-\s]\d{3,4}\b", 0.5),
            Pattern("plate_num_alpha", r"\b\d{1,4}[A-Z]{3,4}\b", 0.5),
            Pattern("plate_compact", r"\b[A-Z]{3}\d{4}\b", 0.5),
        ],
        context=["plate", "license plate", "tag", "registration"],
        supported_language="en",
    )


def build_weapon_serial_recognizer() -> PatternRecognizer:
    return PatternRecognizer(
        supported_entity="WEAPON_SERIAL",
        name="weapon_serial_recognizer",
        patterns=[
            Pattern("weapon_serial_alpha", r"\b[A-Z]{1,3}\d{4,10}[A-Z]{0,2}\b", 0.7),
            Pattern("weapon_serial_num", r"\b\d{6,12}\b", 0.3),
        ],
        context=[
            "serial", "serial no", "serial number",
            "weapon", "firearm", "gun", "pistol", "rifle", "shotgun",
            "s/n",
        ],
        supported_language="en",
    )


def get_all_law_enforcement_recognizers() -> list[PatternRecognizer]:
    """Return all custom law-enforcement recognizers."""
    return [
        build_badge_number_recognizer(),
        build_case_number_recognizer(),
        build_evidence_id_recognizer(),
        build_license_plate_recognizer(),
        build_weapon_serial_recognizer(),
    ]
