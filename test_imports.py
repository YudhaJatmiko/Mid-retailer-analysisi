#!/usr/bin/env python3
"""
Test script to verify all required libraries are properly installed
"""

print("Testing library imports...")

try:
    import pandas as pd
    print("✓ pandas imported successfully")
    print(f"  pandas version: {pd.__version__}")
except ImportError as e:
    print(f"✗ pandas import failed: {e}")

try:
    import numpy as np
    print("✓ numpy imported successfully")
    print(f"  numpy version: {np.__version__}")
except ImportError as e:
    print(f"✗ numpy import failed: {e}")

try:
    import matplotlib.pyplot as plt
    print("✓ matplotlib imported successfully")
    import matplotlib
    print(f"  matplotlib version: {matplotlib.__version__}")
except ImportError as e:
    print(f"✗ matplotlib import failed: {e}")

try:
    import seaborn as sns
    print("✓ seaborn imported successfully")
    print(f"  seaborn version: {sns.__version__}")
except ImportError as e:
    print(f"✗ seaborn import failed: {e}")

try:
    import openpyxl
    print("✓ openpyxl imported successfully")
    print(f"  openpyxl version: {openpyxl.__version__}")
except ImportError as e:
    print(f"✗ openpyxl import failed: {e}")

print("\nIf any imports failed, please run:")
print("pip install pandas numpy matplotlib seaborn openpyxl")