# Contributing

## Local Setup

1. Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

2. Keep resume files in `resume/` (ignored by git).
3. Build profile (optional; main script can do this automatically):

```bash
python3 profile_builder.py
```

4. Run the script:

```bash
python3 topup_jobs_qa.py
```

For daily link checks only:

```bash
RUN_MODE=revalidate python3 topup_jobs_qa.py
```

## Configuration

Defaults are in `config.yaml`:
- role/stack/remote/US keywords
- unavailable and blocked phrases
- retry/backoff settings
- source page limits
- scoring weights

`keywords.yaml` is the main place for role and stack keyword expansion.

Runtime profile is `results/profile.yaml` and is built from the latest resume file.
Job tracking columns are managed automatically:
- `Application Status`
- `Applied Date`
- `Last Checked`
- `Link Status`
- `Job Key`

## Test Checklist

Run before commit:

```bash
python3 -m py_compile topup_jobs_qa.py profile_builder.py
python3 -m unittest discover -s tests -v
```

## How To Add A New Source

1. Add a new `collect_<source>()` function in `topup_jobs_qa.py`.
2. Ensure it returns rows in the same schema as `make_record()`.
3. Pass every candidate URL through `validate_url()`.
4. Add source entry to the `collectors` list in `main()`.
5. Update `README.md` if behavior changes.
