@echo off
echo Starting PPTX2PDF in development mode...
IF EXIST .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)
python -m app.main %*
