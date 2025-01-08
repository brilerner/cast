This app has been deployed on pythonanywhere.com using Brian Lerner's personal account. Here is an overview of what was needed to set it up:

# Created a Virtual Environment:
Used python3 -m venv ~/.virtualenvs/my_project_venv.

# Installed Dependencies:

Activated the venv (source ~/.virtualenvs/my_project_venv/bin/activate) and ran pip install -r requirements.txt.

# Configured WSGI File:

Added the project directory and the virtual environment’s site-packages to sys.path in the WSGI file

# Set the venv path in the PythonAnywhere Web tab.
Restarted the Web App:
Reloaded the web app to apply changes.
