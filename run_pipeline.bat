@echo off
set PYTHONUTF8=1
cd /d "C:\Users\DELL\OneDrive\Desktop\cc-tracker"
"C:\Users\DELL\AppData\Local\Microsoft\WindowsApps\PythonSoftwareFoundation.Python.3.10_qbz5n2kfra8p0\python.exe" run_pipeline.py --bank BDO >> logs\scheduler_log.txt 2>&1
