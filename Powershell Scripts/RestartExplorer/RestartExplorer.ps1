
kill -ProcessName explorer -ErrorAction SilentlyContinue

Start-Sleep -s 2

Start-Process "explorer.exe"