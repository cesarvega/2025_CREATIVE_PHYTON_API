# Tests Documentation

## 📋 Overview

This directory contains all automated tests for the CreativePythonAPI project. Tests are organized by module/service and use `pytest` as the testing framework.

## 🗂️ Test Structure

```
tests/
├── conftest.py                  # Shared fixtures and pytest configuration
├── test_constants.py            # Tests for app/constants.py
├── test_database.py             # Tests for app/config/db.py
├── test_dependencies.py         # Tests for app/api/dependencies.py
├── test_download_utils.py       # Tests for app/utils/download_utils.py
├── test_email_service.py        # Tests for app/services/email_service.py
├── test_excel_service.py        # Tests for app/services/excel_service.py
├── test_settings.py             # Tests for app/config/settings.py
├── test_utils.py                # Tests for app/utils/* modules
└── test_word_service.py         # Tests for app/services/word_service.py
```

## 🚀 Running Tests

### Run All Tests
```bash
pytest
```

### Run Specific Test File
```bash
pytest tests/test_excel_service.py
```

### Run Specific Test Class
```bash
pytest tests/test_excel_service.py::TestExcelProcessingService
```

### Run Specific Test Method
```bash
pytest tests/test_excel_service.py::TestExcelProcessingService::test_process_excel_without_groups
```

### Run Tests by Marker
```bash
# Run only unit tests
pytest -m unit

# Run only integration tests
pytest -m integration

# Run only API tests
pytest -m api

# Exclude slow tests
pytest -m "not slow"
```

### Run with Coverage
```bash
# Install pytest-cov first
pip install pytest-cov

# Run with coverage report
pytest --cov=app --cov-report=html --cov-report=term
```

### Run with Verbose Output
```bash
pytest -v
```

### Run with Print Statements
```bash
pytest -s
```

## 🏷️ Test Markers

Tests are categorized using pytest markers:

- **`@pytest.mark.unit`**: Unit tests that don't require external dependencies
- **`@pytest.mark.integration`**: Integration tests that may require database or files
- **`@pytest.mark.slow`**: Tests that take a long time to run
- **`@pytest.mark.api`**: API endpoint tests

Example usage:
```python
@pytest.mark.unit
def test_something():
    assert True
```

## 🔧 Fixtures

### Built-in Fixtures (from conftest.py)

#### `sample_excel_bytes`
Creates a sample Excel file with test data as bytes.

```python
def test_something(sample_excel_bytes):
    result = process_excel(sample_excel_bytes)
    assert result is not None
```

#### `sample_excel_with_groups`
Creates an Excel file with group data.

```python
def test_with_groups(sample_excel_with_groups):
    result = process_excel(sample_excel_with_groups, has_groups=True)
    assert result.has_groups
```

#### `sample_presentation_metadata`
Provides valid presentation metadata dictionary.

```python
def test_metadata(sample_presentation_metadata):
    assert sample_presentation_metadata["project"] == "TEST_PROJECT"
```

#### `sample_build_metadata`
Provides valid build metadata dictionary.

#### `mock_settings`
Allows mocking settings values for testing.

```python
def test_with_mock_settings(mock_settings):
    mock_settings("max_file_size_mb", 50)
    # Test with mocked setting
```

#### `temp_test_dir`
Provides a temporary directory for test files.

```python
def test_file_operations(temp_test_dir):
    test_file = temp_test_dir / "test.txt"
    test_file.write_text("content")
    assert test_file.exists()
```

## 📝 Writing New Tests

### Test File Naming
- Test files must start with `test_`
- Example: `test_my_service.py`

### Test Class Naming
- Test classes must start with `Test`
- Example: `class TestMyService:`

### Test Method Naming
- Test methods must start with `test_`
- Use descriptive names: `test_process_excel_with_groups`

### Example Test Structure

```python
"""Unit tests for my service."""

import pytest

from app.services.my_service import my_function


@pytest.mark.unit
class TestMyService:
    """Test suite for my service."""

    def test_basic_functionality(self):
        """Test basic functionality."""
        result = my_function("input")
        assert result == "expected"

    def test_with_fixture(self, sample_excel_bytes):
        """Test using a fixture."""
        result = my_function(sample_excel_bytes)
        assert result is not None

    def test_error_handling(self):
        """Test error handling."""
        with pytest.raises(ValueError):
            my_function(None)

    @pytest.mark.slow
    def test_slow_operation(self):
        """Test a slow operation."""
        # This test is marked as slow
        pass
```

## 🧪 Testing Best Practices

### 1. **One Assertion Per Test** (Ideal)
```python
def test_name_extraction(self):
    result = extract_name("John Doe")
    assert result == "John Doe"

def test_name_extraction_with_notation(self):
    result = extract_name("John (JD)")
    assert result == "John"
```

### 2. **Use Descriptive Test Names**
```python
# Good ✅
def test_process_excel_with_empty_file_returns_zero_rows(self):
    pass

# Bad ❌
def test_excel(self):
    pass
```

### 3. **Arrange-Act-Assert Pattern**
```python
def test_something(self):
    # Arrange
    input_data = create_test_data()
    
    # Act
    result = process(input_data)
    
    # Assert
    assert result.is_valid
```

### 4. **Use Fixtures for Setup**
```python
@pytest.fixture
def test_data():
    return {"key": "value"}

def test_with_fixture(test_data):
    assert test_data["key"] == "value"
```

### 5. **Mock External Dependencies**
```python
from unittest.mock import patch

@patch('app.services.email_service.smtp')
def test_send_email(mock_smtp):
    mock_smtp.send.return_value = True
    result = send_email("test@example.com")
    assert result is True
```

## 📊 Code Coverage

To generate a coverage report:

```bash
# Install coverage tool
pip install pytest-cov

# Run tests with coverage
pytest --cov=app --cov-report=html

# Open coverage report
# Windows
start htmlcov/index.html

# Linux/Mac
open htmlcov/index.html
```

## 🐛 Debugging Tests

### Run Tests with PDB Debugger
```bash
pytest --pdb
```

### Stop on First Failure
```bash
pytest -x
```

### Show Local Variables on Failure
```bash
pytest -l
```

### Verbose Output with Full Traceback
```bash
pytest -vv --tb=long
```

## ⚡ Performance

### Run Tests in Parallel
```bash
# Install pytest-xdist
pip install pytest-xdist

# Run tests in parallel
pytest -n auto
```

## 📦 Required Packages

```bash
# Core testing
pytest>=7.0.0
pytest-asyncio>=0.21.0

# Optional but recommended
pytest-cov>=4.0.0        # Coverage reporting
pytest-xdist>=3.0.0      # Parallel test execution
pytest-mock>=3.10.0      # Advanced mocking
```

## 🔍 Continuous Integration

Tests are automatically run on:
- Every pull request
- Every commit to main branch
- Nightly builds (if configured)

## 📚 Additional Resources

- [Pytest Documentation](https://docs.pytest.org/)
- [Pytest Best Practices](https://docs.pytest.org/en/stable/goodpractices.html)
- [Python Testing with pytest (Book)](https://pragprog.com/titles/bopytest/)

## ✅ Test Checklist

Before committing:
- [ ] All tests pass locally
- [ ] New code has corresponding tests
- [ ] Tests are properly marked (unit/integration/etc.)
- [ ] Test names are descriptive
- [ ] No commented-out test code
- [ ] External dependencies are mocked
- [ ] Tests are fast (< 1s per test if possible)

---

**Last Updated:** October 2025  
**Maintained by:** Development Team
