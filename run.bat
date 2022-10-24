node -v
IF %ERRORLEVEL% GTR 0 call winget install OpenJS.NodeJS

tasklist /fi "ImageName eq node.exe" /fo csv 2>NUL | find /I "node.exe">NUL
IF %ERRORLEVEL% EQU 0 call taskkill /f /im "node.exe"

if not exist "./node_modules" call npm install
call node scripts/main.js
PAUSE