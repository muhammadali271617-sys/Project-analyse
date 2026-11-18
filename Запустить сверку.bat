@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Запуск программы «Сверка с фактурами» без чёрного окна
where pyw >nul 2>nul && (start "" pyw "%~dp0start.pyw" & exit /b)
where pythonw >nul 2>nul && (start "" pythonw "%~dp0start.pyw" & exit /b)
where py >nul 2>nul && (start "" py "%~dp0start.pyw" & exit /b)
echo Не найден Python. Установите Python с python.org (галочка "Add Python to PATH").
pause
