#Compiler PyInstaller (change the paths!!! Works for Eduardo): pyinstaller --onedir --add-binary "C:\\Users\eduardo.scheffer\\.conda\\envs\\bside_clean\\Lib\\site-packages\\pywin32_system32\\pythoncom39.dll;." --paths C:\Users\\eduardo.scheffer\\.conda\\envs\\bside_clean\\Lib\\site-packages run_CR.py

# To compile executable file:
# 1. Delete older files from local paste
# 2. Copy new "CR_code"  paste on local folder
# 3. Add "." before "from formulas..." as from .formulas..." below in the "comps.py" file copy in the local folder
# 4. Change excel file_name
# 5. Go to Anaconda Prompt (using the chosen environment >>"conda activate bside_clean") up to the local folder (>>cd [local address])
# 6. Paste and run the Compiler PyInstaller code line above
# 7. Copy and paste the new generated "dist" folder into the original intranet folder (I:\GESTAO\3) Carteira Recomendada\Carteira Recomendada)

# Files needed to run the file standalone: "dist" folder, run_CR, "RUN_Excel" (Windows Batch File)


