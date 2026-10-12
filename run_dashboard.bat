@echo off
cd /d "%~dp0"
echo Refreshing Merchandising data...
python import_pandas_as_pd.py
echo.
echo Opening dashboard at http://localhost:8000/  (close this window to stop)
start "" http://localhost:8000/
python -m http.server 8000
