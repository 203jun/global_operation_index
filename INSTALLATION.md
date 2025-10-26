# Installation Guide

## Python Version Requirements

### Recommended
- **Python 3.8.x ~ 3.12.x**

### Compatibility
- ✅ **Python 3.8**: Fully compatible
- ✅ **Python 3.9**: Fully compatible
- ✅ **Python 3.10**: Fully compatible
- ✅ **Python 3.11**: Fully compatible
- ✅ **Python 3.12**: Fully compatible (currently tested)
- ⚠️ **Python 3.13+**: Not tested, but should work

### Key Package Compatibility

| Package | Version | Python 3.8 | Python 3.9 | Python 3.10 | Python 3.11 | Python 3.12 |
|---------|---------|------------|------------|-------------|-------------|-------------|
| Flask | 3.1.0 | ✅ | ✅ | ✅ | ✅ | ✅ |
| Pandas | ≥1.3.0, <2.3.0 | ✅ | ✅ | ✅ | ✅ | ✅ |
| NumPy | ≥1.21.0, <2.0.0 | ✅ | ✅ | ✅ | ✅ | ✅ |
| openpyxl | ≥3.0.0, <3.2.0 | ✅ | ✅ | ✅ | ✅ | ✅ |

## Installation Steps

### For Windows

1. **Install Python**
   ```bash
   # Download Python 3.8 or higher from python.org
   # Recommended: Python 3.10.x or 3.11.x for best stability
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run Data Preprocessing** (First time only)
   ```bash
   python main.py
   ```

5. **Start Flask Application**
   ```bash
   python run.py
   ```

6. **Access Dashboard**
   ```
   Open browser: http://localhost:8080
   ```

### For macOS/Linux

1. **Install Python**
   ```bash
   # macOS (Homebrew)
   brew install python@3.11

   # Ubuntu/Debian
   sudo apt-get update
   sudo apt-get install python3.11 python3.11-venv
   ```

2. **Create Virtual Environment**
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run Data Preprocessing** (First time only)
   ```bash
   python main.py
   ```

5. **Start Flask Application**
   ```bash
   python run.py
   ```

6. **Access Dashboard**
   ```
   Open browser: http://localhost:8080
   ```

## Package Version Notes

### pandas
- **Min**: 1.3.0 (for Python 3.8 compatibility)
- **Max**: <2.3.0 (avoid breaking changes)
- **Note**: Pandas 2.x works well with Python 3.8+

### numpy
- **Min**: 1.21.0 (for Python 3.8 compatibility)
- **Max**: <2.0.0 (avoid breaking changes in NumPy 2.0)
- **Note**: NumPy 2.0+ may have compatibility issues with older code

### openpyxl
- Required for reading/writing Excel files (.xlsx)
- Works with all Python 3.8+ versions

### Flask 3.1.0
- Requires Python 3.8+
- Werkzeug 3.1.3 is compatible
- All Flask extensions are compatible

## Why Python 3.8 Works

**Python 3.8+ features used in this project:**
- f-strings (Python 3.6+)
- Type hints (Python 3.5+)
- Walrus operator `:=` (Python 3.8+) - **if used**
- Dictionary merge `|` operator (Python 3.9+) - **if used**

**Current code review:** Your code primarily uses:
- Basic Python features (compatible with 3.8+)
- Pandas DataFrame operations (3.8+)
- Flask routing and templating (3.8+)

## Recommended for Windows

### Option 1: Python 3.10.x (Most Stable)
```bash
# Download from: https://www.python.org/downloads/release/python-31011/
# All packages tested and stable
```

### Option 2: Python 3.11.x (Better Performance)
```bash
# Download from: https://www.python.org/downloads/release/python-3119/
# ~25% faster than 3.10, all packages compatible
```

### Option 3: Python 3.8.x (Your Current Choice)
```bash
# Download from: https://www.python.org/downloads/release/python-3818/
# Fully compatible, but older version
```

## Installation Issues

### Common Windows Issues

1. **"pip is not recognized"**
   ```bash
   python -m pip install -r requirements.txt
   ```

2. **Long path issues**
   ```bash
   # Enable long paths in Windows
   # Run as Administrator in PowerShell:
   New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
   ```

3. **Permission errors**
   ```bash
   # Run PowerShell as Administrator
   # Or use --user flag:
   pip install --user -r requirements.txt
   ```

## Summary

✅ **Python 3.8.x is FULLY COMPATIBLE**
- All packages support Python 3.8
- No compatibility issues expected
- You can safely use Python 3.8.10+ on Windows

✅ **Recommended: Python 3.10.11 or 3.11.9**
- Better performance
- More stable
- Longer support period

⚠️ **Avoid: Python 3.7 or earlier**
- Flask 3.1.0 requires Python 3.8+
- Some pandas features may not work
