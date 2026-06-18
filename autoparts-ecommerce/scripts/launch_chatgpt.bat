@echo off
echo Открываю Chrome для бота (отдельный профиль, порт 9222)...
echo После открытия войди в ChatGPT - и оставь это окно открытым.
"C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="%USERPROFILE%\AppData\Local\ChromeBot" ^
  --no-first-run ^
  --no-default-browser-check ^
  https://chatgpt.com
