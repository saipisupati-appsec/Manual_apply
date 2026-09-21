MANUAL JOB SEARCH TOOL
======================

Purpose
-------
This repository automatically collects recent security job postings
for manual application.

No AI is used.
No applications are submitted automatically.


SEARCH COVERAGE
---------------

Regions:
- Europe
- United Kingdom
- Canada

Job areas include:
- Application Security
- Product Security
- Security Engineering
- Cloud Security
- DevSecOps
- Software Security
- API Security
- Security Automation
- Platform Security
- Security Architecture


DATE FILTER
-----------

Only jobs posted within the latest 7 days are included.

The jobs_latest.xlsx file is replaced during each successful run.


VISA / SPONSORSHIP
------------------

The tool searches job-posting text for indicators such as:

- Visa sponsorship
- Visa support
- Work visa
- Work permit
- Relocation
- Immigration support
- Skilled Worker
- EU Blue Card
- LMIA
- Employer sponsorship

Possible values:

Yes / Mentioned
No
Not mentioned

IMPORTANT:

"Not mentioned" does NOT mean the company does not provide
visa sponsorship.

Always verify the company's actual requirements before applying.


EXCEL WORKBOOK
--------------

jobs_latest.xlsx contains:

- Posted Date
- Company
- Job Title
- Location
- Country
- Remote
- Visa/Sponsorship
- Salary
- Experience
- ATS
- Apply
- Status
- Job Description

The Apply column contains clickable links.


APPLICATION STATUS
------------------

New
    Job has not been reviewed.

Opened
    Job posting has been reviewed.

Applied
    Application has been submitted.

Skipped
    Decided not to apply.


AUTOMATION
----------

GitHub Actions runs the scraper automatically.

The workflow:

1. Starts a GitHub-hosted runner.
2. Installs Python.
3. Installs requirements.
4. Runs search_jobs.py.
5. Creates jobs_latest.xlsx.
6. Commits the updated workbook to this repository.

The local computer does not need to run Python.


MANUAL APPLICATION
------------------

The tool only provides a job list.

You manually:

1. Open the job.
2. Review the requirements.
3. Verify location and work authorization.
4. Review sponsorship information.
5. Apply on the employer's official application page.
6. Update the Status column.


DATA RETENTION
--------------

The workbook represents the latest 7-day search window.

Older job-list data is not intentionally accumulated.


IMPORTANT
---------

Job-search results depend on the underlying job sources and their
available data.

Always verify job availability, location, salary, sponsorship,
and application requirements on the original employer posting.