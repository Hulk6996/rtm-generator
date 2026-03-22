# ─────────────────────────────────────────────
#  RTM Generator — Configuration
# ─────────────────────────────────────────────

# ── Paths ────────────────────────────────────
INPUT_FILE  = "data/requirements_sample.xlsx"
OUTPUT_DIR  = "output"
EXCEL_OUT   = "output/rtm_report.xlsx"
WORD_OUT    = "output/rtm_summary.docx"

# ── Sheet names in the input Excel ───────────
SHEET_BR = "Business Requirements"
SHEET_FR = "Functional Requirements"
SHEET_TC = "Test Cases"

# ── Column names ─────────────────────────────
BR_COLS = {
    "id":          "BR_ID",
    "title":       "Title",
    "description": "Description",
    "priority":    "Priority",   # High / Medium / Low
    "category":    "Category",   # e.g. Auth, Reporting, UI
    "source":      "Source",     # e.g. Stakeholder, Regulation
    "status":      "Status",     # Active / Deprecated
}

FR_COLS = {
    "id":          "FR_ID",
    "title":       "Title",
    "description": "Description",
    "br_refs":     "BR_REF",     # comma-separated BR IDs
    "type":        "Type",       # Functional / Integration / UI / Performance
    "component":   "Component",  # e.g. Backend, Frontend, DB
    "status":      "Status",     # Active / Draft / Rejected
}

TC_COLS = {
    "id":          "TC_ID",
    "title":       "Title",
    "fr_refs":     "FR_REF",     # comma-separated FR IDs
    "type":        "Type",       # Manual / Auto
    "result":      "Result",     # Passed / Failed / Not Run / Blocked
    "priority":    "Priority",
}

# ── Colour palette ────────────────────────────
COLORS = dict(
    hdr_bg  = "FF1F3864",   hdr_fg  = "FFFFFFFF",
    sub_bg  = "FF2E75B6",   sub_fg  = "FFFFFFFF",
    grn_bg  = "FFC6EFCE",   grn_fg  = "FF276221",
    yel_bg  = "FFFFEB9C",   yel_fg  = "FF9C6500",
    red_bg  = "FFFFC7CE",   red_fg  = "FF9C0006",
    gry_bg  = "FFF2F2F2",   gry_fg  = "FF595959",
    acc     = "FF4472C4",
    white   = "FFFFFFFF",
    brd     = "FFBDD7EE",
    ora_bg  = "FFFCE4D6",   ora_fg  = "FF833C00",
)

# ── Coverage thresholds (%) ──────────────────
COVERAGE_GREEN  = 80
COVERAGE_YELLOW = 50

# ── Health score weights ─────────────────────
WEIGHT_BR_COVERAGE = 0.40
WEIGHT_FR_COVERAGE = 0.30
WEIGHT_TEST_PASS   = 0.30

# ── Report metadata ───────────────────────────
REPORT_TITLE    = "Requirements Traceability Matrix"
REPORT_PROJECT  = "Demo Project"
REPORT_VERSION  = "1.0"
REPORT_AUTHOR   = "RTM Generator"
