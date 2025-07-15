# Contributing to Outlook Email Extractor

Thank you for considering contributing to Outlook Email Extractor! We welcome all contributions, including bug reports, feature requests, documentation improvements, and code contributions.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [Getting Started](#getting-started)
- [Development Workflow](#development-workflow)
- [Code Style](#code-style)
- [Testing](#testing)
- [Pull Request Process](#pull-request-process)
- [Reporting Bugs](#reporting-bugs)
- [Feature Requests](#feature-requests)

## Code of Conduct

This project and everyone participating in it is governed by our [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold this code.

## Getting Started

1. **Fork** the repository on GitHub
2. **Clone** your fork locally
   ```bash
   git clone https://github.com/your-username/outlook-extract.git
   cd outlook-extract
   ```
3. **Set up** a virtual environment
   ```bash
   python -m venv venv
   # On Windows: .\venv\Scripts\activate
   # On macOS/Linux: source venv/bin/activate
   ```
4. **Install** development dependencies
   ```bash
   pip install -r requirements-dev.txt
   pip install -e .
   ```

## Development Workflow

1. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b bugfix/issue-number-description
   ```

2. Make your changes following the code style guidelines

3. Run tests and ensure they pass
   ```bash
   pytest
   ```

4. Commit your changes with a descriptive message:
   ```bash
   git commit -m "Add feature: brief description of changes"
   ```

5. Push to your fork and open a Pull Request

## Code Style

- Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/) style guide
- Use type hints for all function parameters and return values
- Keep lines under 88 characters (Black's default line length)
- Use docstrings following Google style for all public modules, classes, and functions

### Formatting

We use `black` for code formatting and `isort` for import sorting:

```bash
black .
isort .
```

## Testing

- Write tests for new features and bug fixes
- Ensure all tests pass before submitting a PR
- Use descriptive test function names that describe what they test

Run tests with coverage:

```bash
pytest --cov=outlook_extractor tests/
```

## Pull Request Process

1. Ensure any install or build dependencies are documented in `requirements.txt`
2. Update the README.md with details of changes if needed
3. Increase the version number in `outlook_extractor/__init__.py` following [Semantic Versioning](https://semver.org/)
4. The PR will be reviewed by maintainers and may require changes
5. Once approved, a maintainer will merge the PR

## Reporting Bugs

Please open an issue and include the following information:

- A clear, descriptive title
- Steps to reproduce the issue
- Expected vs actual behavior
- Environment details (OS, Python version, etc.)
- Any relevant error messages or logs

## Feature Requests

Feature requests are welcome! Please open an issue and describe:

- The problem you're trying to solve
- Why this feature would be valuable
- Any alternative solutions you've considered

## License

By contributing, you agree that your contributions will be licensed under the project's [MIT License](LICENSE).
