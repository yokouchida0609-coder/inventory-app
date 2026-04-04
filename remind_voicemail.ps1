Add-Type -AssemblyName System.Windows.Forms
$msg = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("55WZ5a6I6Zu744KS56K66KqN44GX44Gm44GP44Gg44GV44GE"))
$title = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("44Oq44Oe44Kk44Oz44OA44O8"))
[System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
