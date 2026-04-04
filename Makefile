.PHONY: install test test-fast lint format typecheck ci clean benchmark docs docs-build

install:
	pip install -e ".[dev,all]"

test:
	pytest tests/ -v --tb=short --cov=messy_xlsx --cov-report=term-missing

test-fast:
	pytest tests/ -v --tb=short -m "not slow"

lint:
	ruff check src/messy_xlsx/ tests/
	ruff format --check src/messy_xlsx/ tests/

format:
	ruff check --fix src/messy_xlsx/ tests/
	ruff format src/messy_xlsx/ tests/

typecheck:
	mypy src/messy_xlsx/ --ignore-missing-imports

ci: lint typecheck test

clean:
	rm -rf dist/ build/ *.egg-info .coverage htmlcov/ .mypy_cache/ .ruff_cache/ .pytest_cache/

benchmark:
	pytest tests/test_performance/ -v --benchmark-only

docs:
	mkdocs serve

docs-build:
	mkdocs build --strict
