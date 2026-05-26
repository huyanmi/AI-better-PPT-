$ErrorActionPreference = "Stop"

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onefile `
  --windowed `
  --name PPTOptimizer `
  ppt_optimizer_gui.py

Write-Host "Built dist\PPTOptimizer.exe"
