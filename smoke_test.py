import sys
from pathlib import Path

# Add current dir to path
sys.path.insert(0, str(Path(__file__).parent))

print("Loading app.py...")
import app

print("Loading apps.expense...")
import apps.expense

print("Loading apps.travel...")
import apps.travel

print("Loading pdf_gen_travel...")
import pdf_gen_travel

print("All imports successful!")
