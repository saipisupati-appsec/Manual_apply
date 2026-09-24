from __future__ import annotations

import gc
import json
import re
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from typing import Any

import httpx
import pandas as pd
import pyarrow.parquet as pq
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from ats_scrapers import Client

# ============================================================
# Configuration
# ============================================================

DAYS_TO_KEEP = 7
OUTPUT_FILE = Path("jobs_latest.xlsx")
JSON_OUTPUT = Path("docs/jobs_latest.json")
HTTP_TIMEOUT = 90.0  # seconds per ATS download

# Prefer ATS that are relevant for tech / security roles and Europe/Canada.
# Large general boards (eures, bundesagentur, workday, successfactors, oracle)
# are skipped to keep GitHub Actions runtime and memory practical.
PREFERRED_ATS = [
    "greenhouse",
    "lever",
    "ashby",
    "workable",
    "smartrecruiters",
    "personio",
    "ycombinator",
    "wellfound",
    "weworkremotely",
    "gem",
    "jobbankca",
    "welcometothejungle",
    "join_com",
    "softgarden",
    "teamtailor",
    "recruitee",
    "breezy",
    "bamboohr",
    "pinpoint",
    "eightfold",
    "jazzhr",
    "phenom",
    "icims",
    "rippling",
    "dayforce",
    "jobvite",
    "avature",
    "google",
    "apple",
    "amazon",
    "tesla",
    "uber",
    "builtin",
    "mercor",
]

JOB_KEYWORDS = [
    "application security engineer",
    "senior application security engineer",
    "application security",
    "product security engineer",
    "senior product security engineer",
    "product security",
    "security engineer",
    "senior security engineer",
    "software security engineer",
    "cloud security engineer",
    "cloud application security",
    "devsecops engineer",
    "devsecops security engineer",
    "security software engineer",
    "application security architect",
    "product security architect",
    "cloud security architect",
    "api security engineer",
    "security automation engineer",
    "platform security engineer",
    "software security",
]

# Title must contain at least one of these (case-insensitive) after exclude filter
TITLE_INCLUDE_TERMS = [
    "application security",
    "product security",
    "security engineer",
    "software security",
    "cloud security",
    "devsecops",
    "security architect",
    "api security",
    "security automation",
    "platform security",
    "appsec",
    "prodsec",
    "threat modeling",
    "threat modelling",
    "vulnerability management",
    "security research",
    "infrastructure security",
    "security platform",
]

EXCLUDE_TITLE_TERMS = [
    "soc",
    "security operations",
    "grc",
    "governance",
    "risk and compliance",
    "compliance",
    "it support",
    "iam administrator",
    "identity administrator",
    "network security",
    "physical security",
    "security guard",
    "security officer",
    "cybersecurity analyst",
    "security analyst",
]

TARGET_COUNTRIES = [
    "austria", "belgium", "bulgaria", "croatia", "cyprus",
    "czech republic", "czechia", "denmark", "estonia", "finland",
    "france", "germany", "greece", "hungary", "iceland",
    "ireland", "italy", "latvia", "lithuania", "luxembourg",
    "malta", "netherlands", "norway", "poland", "portugal",
    "romania", "slovakia", "slovenia", "spain", "sweden",
    "switzerland", "united kingdom", "uk", "england", "scotland",
    "wales", "northern ireland", "canada",
]

TARGET_COUNTRY_ISO = {
    "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR",
    "DE", "GR", "HU", "IS", "IE", "IT", "LV", "LT", "LU", "MT",
    "NL", "NO", "PL", "PT", "RO", "SK", "SI", "ES", "SE", "CH",
    "GB", "UK", "CA",
}

REMOTE_TERMS = [
    "remote", "fully remote", "work from home", "remote-first",
    "remote work", "work remotely",
]

VISA_POSITIVE_TERMS = [
    "visa sponsorship", "visa sponsor", "visa support",
    "sponsorship available", "sponsorship provided", "visa assistance",
    "immigration support", "immigration sponsorship", "work visa",
    "work permit", "relocation assistance", "relocation support",
    "relocation package", "skilled worker", "skilled worker visa",
    "eu blue card", "highly skilled migrant", "lmia",
    "employer sponsorship", "will sponsor", "sponsors visas",
]

VISA_NEGATIVE_TERMS = [
    "no visa sponsorship", "visa sponsorship is not available",
    "we do not sponsor", "we don't sponsor", "unable to sponsor",
    "cannot sponsor", "must have the right to work",
    "must already have the right to work", "without sponsorship",
    "no sponsorship", "does not sponsor", "not able to sponsor",
]

COLUMNS_NEEDED = [
    "title", "company", "location", "country_iso", "region",
    "is_remote", "salary_summary", "salary_min", "salary_max",
    "salary_currency", "employment_type", "description",
    "posted_at", "apply_url", "url", "ats_type",
]


@dataclass
class Job:
    posted_date: str
    company: str
    title: str
    location: str
    country: str
    remote: str
    visa: str
    salary: str
    experience: str
    ats: str
    apply_url: str
    job_url: str
    description: str


def log(msg: str) -> None:
    print(msg, flush=True)
