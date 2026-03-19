from __future__ import annotations

import os

# Keep env override first; fallback keeps current team workflow simple.
KIT_API_KEY = os.getenv("KIT_API_KEY", "kit_7d6b10fad06f88e1d0e47e45ef92e9cc")
KIT_API_BASE = "https://api.kit.com/v4"

# Dashboard windows
ROLLING_DAYS_SALES_CTOR = 30
REWARM_WINDOW_DAYS = 30
COLD_LOOKBACK_DAYS = 90
SNAPSHOT_MONTHS = 4
CHURN_MONTHS = 6

# Broadcast labels from description/internal note
VALUE_LABEL_KEYWORDS = ["value"]
SALES_LABEL_KEYWORDS = [
    "pre-sales",
    "post-sales",
    "sales",
    "hype email",
    "launch email",
    "launch",
    "pre-launch-workshop",
    "pre launch workshop",
    "pre-launch",
    "post-launch",
]

# Segment tag patterns (API-only; no CSV dependencies)
LEAD_MAGNET_RESPONSE_PATTERNS = [
    "i'm transitioning into an ai/tech career",
    "i work at a company",
    "i'm an independent consultant or freelancer",
    "i own/run a business or lead a team",
    "i'm building ai products or solutions",
    "building ai products - eg.",
]

WORKSHOP_TAG_PATTERNS = [
    "ai app sprint",
    "freelance accelerator",
    "agent breakthrough",
]

BOOTCAMP_TAG_PATTERNS = [
    "ai agent core [",
    "ai agent bootcamp core [",
]
