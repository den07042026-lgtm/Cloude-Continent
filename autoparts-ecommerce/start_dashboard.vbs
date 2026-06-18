Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\Admin\Documents\Autoparts_Ecommerce"
WshShell.Run "uv run --with customtkinter,openpyxl scripts/dashboard.py", 0, False
Set WshShell = Nothing
