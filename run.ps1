# Helper script to setup the environment and run the app on Windows.

# Create a virtual environment if it doesn't exist
if (-not (Test-Path .venv)) {
    python -m venv .venv
}

# Activate the virtual environment
.venv\Scripts\Activate.ps1

# Install pypiwin32 if it doesn't exist
if (pip show pypiwin32 -q) {
    pip install pypiwin32
}

# Install the required packages
pip install -r requirements.txt

# Run the app
python app.py
