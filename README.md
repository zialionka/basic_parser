# Remote Job Search Project

Small project to find and top up US remote jobs based on resume content.

## Structure

- `topup_jobs_qa.py`: main script.
- `profile_builder.py`: builds a profile from the resume.
- `config.yaml`: all filters and settings.
- `keywords.yaml`: centralized role and stack keywords.
- `tests/`: unit tests.
- `CONTRIBUTING.md`: contribution notes.
- `resume/`: resume files (gitignored).
- `results/`: outputs (`jobs_qa.xlsx`, `jobs_qa.json`, run reports; gitignored).
- `logs/`: run logs (gitignored).

## What The Script Does

- Parses the latest resume file from `resume/` and auto-builds `results/profile.yaml`.
- Merges `config.yaml` defaults with profile keywords from the resume.
- Finds remote jobs from supported sources.
- Filters for Remote + US and role relevance using runtime profile terms.
- Validates each job URL.
- Revalidates all `new` jobs on every run.
- Marks closed links as `Application Status=closed`.
- Removes duplicates by `Job URL`, `Job Key`, and `Job Title + Company`.
- Tops up the table to keep `TARGET` active `new` jobs.
- Exports JSON and writes a run report.

## Quick Start

1. Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

2. Put the resume in `resume/`.

3. Run:

```bash
python3 topup_jobs_qa.py
```

4. Check output:
- `results/jobs_qa.xlsx`
- `results/jobs_qa.json`
- `results/latest_run_report.json`
- `results/profile.yaml`

## Configuration

Defaults are in `config.yaml`:
- keyword filters (role/stack/remote/US)
- closed-job and blocked-page phrases
- request retry/backoff
- source page limits
- scoring weights

`keywords.yaml` is loaded first for `role_keywords` and `stack_keywords`.

Supported env overrides:
- `TARGET`
- `RUN_MODE` (`topup` or `revalidate`)
- `PRUNE_CLOSED` (`1` default, set `0` to keep closed rows)
- `KEYWORDS_FILE`
- `RESULTS_DIR`
- `RESUME_DIR`
- `PROFILE_FILE`
- `JOBS_XLSX`
- `JOBS_JSON`
- `CONFIG_FILE`

Example:

```bash
TARGET=150 python3 topup_jobs_qa.py
RUN_MODE=revalidate python3 topup_jobs_qa.py
```

## Before Commit

```bash
python3 -m py_compile topup_jobs_qa.py profile_builder.py
python3 -m unittest discover -s tests -v
```
