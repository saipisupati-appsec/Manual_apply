from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from ats_scrapers import search
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation


# ============================================================
# Configuration
# ============================================================

DAYS_TO_KEEP = 7
OUTPUT_FILE = Path("jobs_latest.xlsx")

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

TARGET_COUNTRIES = [
    "austria",
    "belgium",
    "bulgaria",
    "croatia",
    "cyprus",
    "czech republic",
    "czechia",
    "denmark",
    "estonia",
    "finland",
    "france",
    "germany",
    "greece",
    "hungary",
    "iceland",
    "ireland",
    "italy",
    "latvia",
    "lithuania",
    "luxembourg",
    "malta",
    "netherlands",
    "norway",
    "poland",
    "portugal",
    "romania",
    "slovakia",
    "slovenia",
    "spain",
    "sweden",
    "switzerland",
    "united kingdom",
    "uk",
    "england",
    "scotland",
    "wales",
    "northern ireland",
    "canada",
]

REMOTE_TERMS = [
    "remote",
    "fully remote",
    "work from home",
    "remote-first",
]

VISA_POSITIVE_TERMS = [
    "visa sponsorship",
    "visa sponsor",
    "visa support",
    "sponsorship available",
    "sponsorship provided",
    "visa assistance",
    "immigration support",
    "immigration sponsorship",
    "work visa",
    "work permit",
    "relocation assistance",
    "relocation support",
    "relocation package",
    "skilled worker",
    "skilled worker visa",
    "eu blue card",
    "highly skilled migrant",
    "lmia",
    "employer sponsored",
    "employer sponsorship",
]

VISA_NEGATIVE_TERMS = [
    "no visa sponsorship",
    "visa sponsorship is not available",
    "we do not sponsor",
    "we don't sponsor",
    "unable to sponsor",
    "cannot sponsor",
    "must have the right to work",
    "must already have the right to work",
    "without sponsorship",
]

EXCLUDE_TITLE_TERMS = [
    "soc analyst",
    "soc engineer",
    "security operations",
    "grc",
    "governance risk compliance",
    "compliance analyst",
    "compliance engineer",
    "it support",
    "iam administrator",
    "identity administrator",
    "network security engineer",
    "network security analyst",
    "physical security",
]


# ============================================================
# Data model
# ============================================================

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


# ============================================================
# Utility functions
# ============================================================

def text(value: Any) -> str:
    """Safely convert any value to searchable text."""
    if value is None:
        return ""

    if isinstance(value, (list, tuple, set)):
        return " ".join(text(item) for item in value)

    if isinstance(value, dict):
        return " ".join(text(v) for v in value.values())

    return str(value)


def first_value(item: dict[str, Any], *keys: str) -> str:
    """Return the first non-empty value from possible field names."""
    for key in keys:
        value = text(item.get(key)).strip()
        if value:
            return value

    return ""


def parse_date(value: Any) -> datetime | None:
    """Convert common date formats to timezone-aware datetime."""
    if value is None:
        return None

    if isinstance(value, datetime):
        if value.tzinfo is None:
            return value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc)

    raw = str(value).strip()

    if not raw:
        return None

    raw = raw.replace("Z", "+00:00")

    try:
        parsed = datetime.fromisoformat(raw)

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)

        return parsed.astimezone(timezone.utc)

    except ValueError:
        pass

    formats = [
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).replace(
                tzinfo=timezone.utc
            )
        except ValueError:
            continue

    return None


# ============================================================
# Search
# ============================================================

def search_jobs() -> list[dict[str, Any]]:
    """
    Search the ATS scraper using the configured job keywords.

    The scraper is queried once per keyword and results are combined.
    """

    results: list[dict[str, Any]] = []

    print("Starting job search...")

    for keyword in JOB_KEYWORDS:
        print(f"Searching: {keyword}")

        try:
            jobs = search(
                query=keyword,
                limit=100,
            )

            if jobs:
                results.extend(jobs)

        except Exception as exc:
            print(f"Search failed for '{keyword}': {exc}")

    print(f"Raw results collected: {len(results)}")

    return results


# ============================================================
# Filtering
# ============================================================

def is_recent(
    job: dict[str, Any],
    cutoff: datetime,
) -> bool:

    posted = first_value(
        job,
        "posted_at",
        "date_posted",
        "posted_date",
        "created_at",
    )

    posted_dt = parse_date(posted)

    if posted_dt is None:
        return False

    return posted_dt >= cutoff


def is_target_location(job: dict[str, Any]) -> bool:

    location = first_value(
        job,
        "location",
        "locations",
        "country",
        "region",
    ).lower()

    description = first_value(
        job,
        "description",
        "job_description",
        "content",
    ).lower()

    combined = f"{location} {description}"

    if any(
        country in location
        for country in TARGET_COUNTRIES
    ):
        return True

    if any(
        term in location
        for term in REMOTE_TERMS
    ):
        return True

    if any(
        term in combined
        for term in REMOTE_TERMS
    ):
        return True

    return False


def is_relevant_title(title: str) -> bool:

    title_lower = title.lower()

    if any(
        term in title_lower
        for term in EXCLUDE_TITLE_TERMS
    ):
        return False

    return any(
        keyword.lower() in title_lower
        for keyword in JOB_KEYWORDS
    )


def filter_jobs(
    raw_jobs: list[dict[str, Any]],
    cutoff: datetime,
) -> list[dict[str, Any]]:

    filtered = []

    for job in raw_jobs:

        title = first_value(
            job,
            "title",
            "job_title",
            "name",
        )

        if not title:
            continue

        if not is_recent(job, cutoff):
            continue

        if not is_target_location(job):
            continue

        if not is_relevant_title(title):
            continue

        filtered.append(job)

    print(f"After filtering: {len(filtered)}")

    return filtered


# ============================================================
# Visa detection
# ============================================================

def detect_visa_status(
    job: dict[str, Any],
) -> str:

    description = first_value(
        job,
        "description",
        "job_description",
        "content",
    ).lower()

    location = first_value(
        job,
        "location",
        "country",
    ).lower()

    combined = f"{description} {location}"

    if any(
        term in combined
        for term in VISA_NEGATIVE_TERMS
    ):
        return "No"

    if any(
        term in combined
        for term in VISA_POSITIVE_TERMS
    ):
        return "Yes / Mentioned"

    return "Not mentioned"


# ============================================================
# Normalization
# ============================================================

def normalize_job(
    job: dict[str, Any],
) -> Job:

    title = first_value(
        job,
        "title",
        "job_title",
        "name",
    )

    company = first_value(
        job,
        "company",
        "company_name",
        "employer",
    )

    location = first_value(
        job,
        "location",
        "locations",
        "region",
    )

    country = first_value(
        job,
        "country",
    )

    description = first_value(
        job,
        "description",
        "job_description",
        "content",
    )

    posted = first_value(
        job,
        "posted_at",
        "date_posted",
        "posted_date",
        "created_at",
    )

    apply_url = first_value(
        job,
        "apply_url",
        "application_url",
    )

    job_url = first_value(
        job,
        "job_url",
        "url",
        "link",
    )

    if not apply_url:
        apply_url = job_url

    remote = (
        "Yes"
        if any(
            term in f"{location} {description}".lower()
            for term in REMOTE_TERMS
        )
        else "No"
    )

    return Job(
        posted_date=posted,
        company=company or "Unknown",
        title=title,
        location=location or "Not specified",
        country=country or "Not specified",
        remote=remote,
        visa=detect_visa_status(job),
        salary=first_value(
            job,
            "salary",
            "salary_text",
        ),
        experience=first_value(
            job,
            "experience",
            "experience_level",
        ),
        ats=first_value(
            job,
            "ats",
            "source",
        ),
        apply_url=apply_url,
        job_url=job_url,
        description=description,
    )


# ============================================================
# Deduplication
# ============================================================

def deduplicate_jobs(
    jobs: list[Job],
) -> list[Job]:

    unique: dict[
        tuple[str, str, str],
        Job,
    ] = {}

    for job in jobs:

        key = (
            job.company.strip().lower(),
            job.title.strip().lower(),
            job.location.strip().lower(),
        )

        if key not in unique:
            unique[key] = job

    result = list(unique.values())

    print(
        f"After deduplication: {len(result)}"
    )

    return result


# ============================================================
# Preserve existing statuses
# ============================================================

def load_existing_statuses() -> dict[
    tuple[str, str, str],
    str,
]:
    """
    Read the previous jobs_latest.xlsx and preserve the
    user's Status values for jobs that are still present.
    """

    statuses: dict[
        tuple[str, str, str],
        str,
    ] = {}

    if not OUTPUT_FILE.exists():
        print(
            "No existing workbook found. "
            "All jobs will start as New."
        )
        return statuses

    try:

        workbook = load_workbook(
            OUTPUT_FILE,
            read_only=True,
            data_only=True,
        )

        if "Latest Jobs" not in workbook.sheetnames:
            workbook.close()
            return statuses

        sheet = workbook["Latest Jobs"]

        headers = {
            cell.value: cell.column
            for cell in sheet[1]
        }

        company_col = headers.get("Company")
        title_col = headers.get("Job Title")
        location_col = headers.get("Location")
        status_col = headers.get("Status")

        if not all(
            [
                company_col,
                title_col,
                location_col,
                status_col,
            ]
        ):
            workbook.close()
            return statuses

        for row in sheet.iter_rows(
            min_row=2,
            values_only=True,
        ):

            company = str(
                row[company_col - 1] or ""
            ).strip().lower()

            title = str(
                row[title_col - 1] or ""
            ).strip().lower()

            location = str(
                row[location_col - 1] or ""
            ).strip().lower()

            status = str(
                row[status_col - 1] or "New"
            ).strip()

            if (
                company
                and title
                and location
                and status
                in {
                    "New",
                    "Opened",
                    "Applied",
                    "Skipped",
                }
            ):
                statuses[
                    (
                        company,
                        title,
                        location,
                    )
                ] = status

        workbook.close()

        print(
            f"Preserved statuses for "
            f"{len(statuses)} existing jobs."
        )

    except Exception as exc:

        print(
            f"Could not read existing status values: {exc}"
        )

    return statuses


# ============================================================
# Excel generation
# ============================================================

def create_workbook(
    jobs: list[Job],
    existing_statuses: dict[
        tuple[str, str, str],
        str,
    ],
) -> None:

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "Latest Jobs"

    headers = [
        "Posted Date",
        "Company",
        "Job Title",
        "Location",
        "Country",
        "Remote",
        "Visa/Sponsorship",
        "Salary",
        "Experience",
        "ATS",
        "Apply",
        "Status",
        "Job Description",
    ]

    sheet.append(headers)

    # Header formatting
    for cell in sheet[1]:

        cell.font = Font(bold=True)

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    # Job rows
    for job in jobs:

        status_key = (
            job.company.strip().lower(),
            job.title.strip().lower(),
            job.location.strip().lower(),
        )

        previous_status = existing_statuses.get(
            status_key,
            "New",
        )

        row = [
            job.posted_date,
            job.company,
            job.title,
            job.location,
            job.country,
            job.remote,
            job.visa,
            job.salary,
            job.experience,
            job.ats,
            job.apply_url,
            previous_status,
            job.description,
        ]

        sheet.append(row)

        row_number = sheet.max_row

        # Clickable Apply link
        apply_cell = sheet.cell(
            row=row_number,
            column=11,
        )

        if job.apply_url:

            apply_cell.hyperlink = job.apply_url

            apply_cell.font = Font(
                color="0563C1",
                underline="single",
            )

    # Status dropdown
    status_validation = DataValidation(
        type="list",
        formula1='"New,Opened,Applied,Skipped"',
        allow_blank=False,
    )

    sheet.add_data_validation(
        status_validation
    )

    status_validation.add(
        f"L2:L{max(sheet.max_row, 2)}"
    )

    # Conditional formatting
    new_fill = PatternFill(
        fill_type="solid",
        fgColor="DDEBF7",
    )

    opened_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC",
    )

    applied_fill = PatternFill(
        fill_type="solid",
        fgColor="E2F0D9",
    )

    skipped_fill = PatternFill(
        fill_type="solid",
        fgColor="E7E6E6",
    )

    data_range = (
        f"A2:M{max(sheet.max_row, 2)}"
    )

    sheet.conditional_formatting.add(
        data_range,
        FormulaRule(
            formula=['$L2="New"'],
            fill=new_fill,
        ),
    )

    sheet.conditional_formatting.add(
        data_range,
        FormulaRule(
            formula=['$L2="Opened"'],
            fill=opened_fill,
        ),
    )

    sheet.conditional_formatting.add(
        data_range,
        FormulaRule(
            formula=['$L2="Applied"'],
            fill=applied_fill,
        ),
    )

    sheet.conditional_formatting.add(
        data_range,
        FormulaRule(
            formula=['$L2="Skipped"'],
            fill=skipped_fill,
        ),
    )

    # General worksheet settings
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    widths = {
        "A": 15,
        "B": 25,
        "C": 40,
        "D": 30,
        "E": 18,
        "F": 10,
        "G": 20,
        "H": 20,
        "I": 18,
        "J": 15,
        "K": 45,
        "L": 12,
        "M": 80,
    }

    for column, width in widths.items():
        sheet.column_dimensions[
            column
        ].width = width

    for row in sheet.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    sheet.row_dimensions[1].height = 25

    # README sheet
    readme = workbook.create_sheet(
        "README"
    )

    readme_lines = [
        "Manual Job Search Tool",
        "",
        "Purpose:",
        "This workbook contains security jobs discovered during the latest 7-day search window.",
        "",
        "How to use:",
        "1. Open the Latest Jobs sheet.",
        "2. Click the Apply link.",
        "3. Review the job manually.",
        "4. Change Status using the dropdown.",
        "",
        "Status:",
        "New = Not reviewed",
        "Opened = Job posting reviewed",
        "Applied = Application submitted",
        "Skipped = Decided not to apply",
        "",
        "Visa/Sponsorship:",
        "Yes / Mentioned = The posting contains sponsorship, visa, immigration, or relocation wording.",
        "No = The posting contains wording indicating sponsorship is unavailable.",
        "Not mentioned = No sponsorship wording was detected.",
        "",
        "Important:",
        "Visa detection is based only on the job-posting text.",
        "Not mentioned does NOT mean sponsorship is unavailable.",
        "",
        "The workbook is regenerated automatically by GitHub Actions.",
        "Existing Status values are preserved when the same job appears in the next refresh.",
    ]

    for line in readme_lines:

        readme.append([line])

    readme.column_dimensions[
        "A"
    ].width = 110

    for row in readme.iter_rows():

        for cell in row:

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    workbook.save(OUTPUT_FILE)

    print(
        f"Created: {OUTPUT_FILE}"
    )


# ============================================================
# Main
# ============================================================

def main() -> None:

    now = datetime.now(timezone.utc)

    cutoff = now - timedelta(
        days=DAYS_TO_KEEP
    )

    raw_jobs = search_jobs()

    filtered_jobs = filter_jobs(
        raw_jobs,
        cutoff,
    )

    normalized_jobs = [
        normalize_job(job)
        for job in filtered_jobs
    ]

    jobs = deduplicate_jobs(
        normalized_jobs
    )

    # Sort newest first
    jobs.sort(
        key=lambda job:
        parse_date(job.posted_date)
        or datetime.min.replace(
            tzinfo=timezone.utc
        ),
        reverse=True,
    )

    # Preserve existing user statuses
    existing_statuses = (
        load_existing_statuses()
    )

    create_workbook(
        jobs,
        existing_statuses,
    )

    print(
        "Job search completed successfully."
    )


if __name__ == "__main__":
    main()
