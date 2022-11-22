powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%USERPROFILE%\AppData\Roaming\Microsoft\Windows\SendTo\BulkMail.lnk');$s.TargetPath='wscript.exe';$s.Arguments = '%CD%\BulkMail.vbs';$s.Save()"
echo "Installation complete. Press any key..."
pause