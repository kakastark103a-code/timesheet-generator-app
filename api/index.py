import sys
import os

# Add parent directory to path so Flask app modules are found
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Ensure OUTPUT_DIR exists at startup (Vercel /tmp is writable)
import tempfile
_output_dir = os.path.join(tempfile.gettempdir(), 'generated_timesheets')
os.makedirs(_output_dir, exist_ok=True)

from app import app  # noqa: F401 - Vercel expects 'app' WSGI object in this file
