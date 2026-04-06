@echo off
netsh advfirewall firewall add rule name="Node 3000" dir=in action=allow protocol=TCP localport=3000
netsh advfirewall firewall add rule name="reminder-app 8080" dir=in action=allow protocol=TCP localport=8080
echo ポート3000と8080を開放しました！
pause
