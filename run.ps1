# Helper script to setup the environment and run the app on Windows.

# Create a virtual environment if it doesn't exist
if (-not (Test-Path .venv)) {
    python -m venv .venv
}

# Activate the virtual environment
.venv\Scripts\Activate.ps1

# Install the required packages
pip install -r requirements.txt

# Install pypiwin32
pip install pypiwin32

# Run the app
python main.py
