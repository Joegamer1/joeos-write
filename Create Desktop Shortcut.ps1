$ErrorActionPreference = 'Stop'
$appRoot = $PSScriptRoot
$appExe = Join-Path $appRoot 'dist-keyboard\JoeOS Write\JoeOS Write.exe'
if (-not (Test-Path -LiteralPath $appExe)) { throw 'Build the application first using Build.cmd.' }
$shortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'JoeOS Write.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $appExe
$shortcut.WorkingDirectory = Split-Path $appExe
$iconSource = Join-Path $appRoot 'assets\write-icon.ico'
$iconHash = (Get-FileHash -LiteralPath $iconSource -Algorithm SHA256).Hash.Substring(0,12)
$iconPath = Join-Path (Split-Path $appExe) ('JoeOS-Write-' + $iconHash + '.ico')
Copy-Item -LiteralPath $iconSource -Destination $iconPath -Force
$shortcut.IconLocation = $iconPath + ',0'
$shortcut.Description = 'JoeOS Write - local Word document editor'
$shortcut.Save()
$verify = $shell.CreateShortcut($shortcutPath)
if ($verify.TargetPath -ne $appExe) { throw 'Shortcut verification failed.' }
if ($verify.IconLocation -ne ($iconPath + ',0')) { throw 'Shortcut icon verification failed.' }
if ((Get-FileHash -LiteralPath $iconPath).Hash -ne (Get-FileHash -LiteralPath $iconSource).Hash) { throw 'Icon copy verification failed.' }
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class JoeShortcutRefresh {
 [DllImport("shell32.dll", CharSet=CharSet.Unicode)]
 public static extern void SHChangeNotify(uint eventId,uint flags,string item,IntPtr other);
}
'@
[JoeShortcutRefresh]::SHChangeNotify(0x2000,0x5,$shortcutPath,[IntPtr]::Zero)
Write-Output "Created and verified: $shortcutPath"
