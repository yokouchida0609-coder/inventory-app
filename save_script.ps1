$content = @'
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("留守電を確認してください", "リマインダー", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
'@
[System.IO.File]::WriteAllText("C:\Users\yoko7\reminder-app\remind_voicemail.ps1", $content, [System.Text.UTF8Encoding]::new($true))
