# Contributing to grepxcel

Thanks for your interest in contributing!

## Branch model

| Branch | Purpose |
|---|---|
| `main` | Stable, always green. Never commit directly — changes arrive via PR from `dev`. |
| `dev` | Active development. **All pull requests target this branch.** |

## Step-by-step guide

### 1. Fork the repository

Click **Fork** on the top-right of the GitHub page. GitHub creates your own copy of the repo under your account.

### 2. Clone your fork

```bash
git clone https://github.com/<your-username>/grepxcel.git
cd grepxcel
```

### 3. Set up the development environment

```bash
python -m venv .venv

# Core dependencies + test runner
.venv/bin/pip install -r requirements.txt -r requirements-dev.txt

# Optional: install the suggest feature (heavy — compiles C++ code)
# .venv/bin/pip install -r requirements-suggest.txt
```

### 4. Create a feature branch from `dev`

Always branch off `dev`, never off `main`:

```bash
git checkout dev
git pull origin dev          # make sure you're up to date
git checkout -b my-feature
```

### 5. Make your changes

A few rules:

- No inline `python -c "..."` scripts — put any one-off code in `tmp-scripts/`
- snake_case for variables/functions, CamelCase for classes
- Maximum line length: 88 characters
- All tests must pass before opening a PR

### 6. Run the tests

```bash
.venv/bin/pytest tests/ -q
```

### 7. Push and open a Pull Request

```bash
git push origin my-feature
```

Then open a PR on GitHub. **Set the base branch to `dev`**, not `main`:

```
your-fork:my-feature  →  scpg/grepxcel:dev
```

In the PR description include:

- What the change does and why
- Any test evidence (new test added, or existing test output)

## CI

GitHub Actions runs the full test suite on every push and PR. Your PR must be green before it can be merged.
