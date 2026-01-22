# StudentConnect

Repository for `StudentConnect.py` — a small integration/sync script that connects data and services such as Mailchimp and FMCSA. This README explains how to set up, configure, run, and contribute to the project safely.

> Important: Do NOT store secrets or production data in this repository. Use environment variables and GitHub Secrets for CI.

## Table of contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Quick installation](#quick-installation)
- [Configuration (.env)](#configuration-env)
- [Usage](#usage)
- [Project layout](#project-layout)
- [Development and tests](#development-and-tests)
- [Security best practices](#security-best-practices)
- [CI / Deployment](#ci--deployment)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Overview

`StudentConnect.py` contains logic to interact with external services (for example Mailchimp) and to process or sync local datasets (for example transporter/loads data). The repository is intended to be run locally, in scheduled jobs, or in CI.

If you want function-level documentation (example inputs/outputs or endpoint docs), paste the function signatures and I will add specific examples.

## Features

- Mailchimp integration examples (sending contacts, authenticating with API key).
- Hooks for FMCSA API key usage (if applicable).
- Local data handling under `data/` (recommended to keep out of git for real data).
- Minimal dependencies so it is easy to run in a virtualenv or a container.

## Requirements

- Python 3.9+ (3.11 recommended)
- Git
- (Optional) `venv` or `virtualenv` for dependency isolation

## Quick installation

```bash
# Clone the repo (if you haven't already):
# git clone git@github.com:<your-username>/student-connect.git
# cd student-connect

# Create and activate a virtual environment (macOS / Linux):
python3 -m venv .venv
source .venv/bin/activate

# Update pip and install dependencies:
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Configuration (.env)

Copy the example `.env` and fill in your secrets. Never commit `.env` to the repo.

```bash
cp .env.example .env
# edit .env with your editor
```

Common environment variables used by the project (examples):

- `API_KEY` — generic API key used internally by the script
- `FMCSA_WEBKEY` — FMCSA API key (if the code calls FMCSA)
- `MAILCHIMP_API_KEY` — Mailchimp API key (if the code uses Mailchimp)
- `MAILCHIMP_LIST_ID` — Mailchimp list/audience ID
- `OTHER_SECRET` — placeholder for other secrets

If your project requires a JSON key file (for example a GCP service account), do not commit that file; reference its path from an environment variable and store the file securely outside the repo or in the CI secrets store.

## Usage

Run the script directly from the `src/` directory:

```bash
python src/StudentConnect.py
```

Example running with inline environment values (useful for quick tests but not recommended for production):

```bash
FMCSA_WEBKEY="your_fmcsa_key" MAILCHIMP_API_KEY="your_mailchimp_key" python src/StudentConnect.py
```

If `StudentConnect.py` exposes functions, prefer creating a small runner (e.g. `scripts/run.py`) that imports and exercises those functions. If you want, I can add a runner file.

## Project layout

Recommended layout for this repo:

```
student-connect/
├─ .env.example
├─ .gitignore
├─ README.md
├─ requirements.txt
└─ src/
   └─ StudentConnect.py
├─ data/           # ignored in git; example/sample data can live in data/example/
└─ .github/workflows/  # CI workflows (optional)
```

If the project grows, restructure into a package like `src/student_connect/` and add unit tests under `tests/`.

## Development and tests

Install dev dependencies and run linters / tests:

```bash
pip install -r requirements.txt
pip install pytest flake8

# Run lint (optional):
flake8 src || true

# Run tests (if you add tests):
pytest
```

If you want, I can add a simple `pytest` test suite skeleton for core functions.

## Security best practices

- Never commit `.env`, secret JSON files, or production `data/` into the repo.
- Add these entries to `.gitignore`: `.env`, `data/`, `*.db`, `*.sqlite`, `*.key`, `.clasprc.json`.
- Before pushing, scan for secrets with a quick grep:

```bash
grep -nE "API[_-]?KEY|SECRET|TOKEN|PASSWORD|PRIVATE[_-]?KEY|ACCESS[_-]?KEY|FMCSA|MAILCHIMP" -R src || true
```

- Use GitHub Secrets (Repository -> Settings -> Secrets) to store API keys for workflows.
- Rotate credentials if you accidentally commit a secret.

## CI / Deployment

A minimal GitHub Actions workflow that installs dependencies and runs tests could look like this (save as `.github/workflows/ci.yml`):

```yaml
name: CI
on: [push]
jobs:
  test:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Setup Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.11'
      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
      - name: Run tests
        run: |
          pip install pytest
          pytest || true
```

If you want to deploy the script to a server or a Cloud Run service, I can provide a deployment workflow and instructions — that will require storing service credentials in GitHub Secrets.

## Contributing

1. Fork the repository.
2. Create a topic branch: `git checkout -b feat/your-feature`.
3. Make small, focused commits with clear messages.
4. Push your branch and open a Pull Request describing the change.

Consider adding a `CONTRIBUTING.md` if you expect outside contributors.

## License

Add your preferred license here (for example MIT):

```
MIT License
Copyright (c) <Your Name>
```

## Contact

If you need help with any of the following, open an issue or ping me in the repo:

- Adding automated tests
- Creating CI that deploys to a target environment
- Removing accidental secrets from git history

---

If you want, I can now:

- add a `CONTRIBUTING.md` file,
- add a minimal GitHub Actions workflow file under `.github/workflows/ci.yml`,
- create a small `scripts/run.py` runner that imports `StudentConnect.py` functions, or
- add a test skeleton for `pytest`.

Tell me which of those you want and I will create them.
