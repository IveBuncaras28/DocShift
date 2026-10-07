param([string]$App,[string]$Pyw)
$sh = New-Object -ComObject WScript.Shell
foreach ($folder in 'Desktop','Programs') {
  $l = $sh.CreateShortcut([Environment]::GetFolderPath($folder) + '\DocShift.lnk')
  $l.TargetPath = $Pyw; $l.Arguments = '"' + $App + '\app.py"'; $l.WorkingDirectory = $App
  $l.IconLocation = $App + '\icon.ico'; $l.Description = 'DocShift - convert documents and edit PDFs'; $l.Save()
}
