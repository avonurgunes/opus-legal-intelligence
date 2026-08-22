@echo off
setlocal
cd /d "%~dp0"
title Opus Legal Intelligence

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set "PY=python"
  ) else (
    echo.
    echo Python bulunamadi.
    echo OLI'yi ilk kez calistirmak icin Python 3.11 veya daha yenisi gerekli.
    echo https://www.python.org/downloads/windows/
    echo Kurulumda "Add python.exe to PATH" secenegini isaretleyin.
    echo.
    pause
    exit /b 1
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo.
  echo OLI ilk kez hazirlaniyor...
  %PY% -m venv .venv
  if errorlevel 1 goto :error
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  if errorlevel 1 goto :error
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :error
)

echo.
echo Opus Legal Intelligence aciliyor...
start "" http://localhost:8501
".venv\Scripts\python.exe" -m streamlit run app.py --server.address localhost --server.port 8501
exit /b 0

:error
echo.
echo Kurulum sirasinda bir hata olustu. Bu pencerenin ekran goruntusunu ChatGPT'ye gonderebilirsiniz.
pause
exit /b 1
