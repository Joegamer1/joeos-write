param([Parameter(Mandatory=$true)][string]$Source)
$ErrorActionPreference = 'Stop'
$destination = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$joeFontDestination = Join-Path $env:USERPROFILE '.joeos\fonts'
New-Item -ItemType Directory -Path $joeFontDestination -Force | Out-Null
$registry = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
New-Item -Path $registry -Force | Out-Null
Add-Type @'
using System;
using System.Runtime.InteropServices;
public class JoeFontInstall {
 [DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern int AddFontResourceEx(string file,uint flags,IntPtr reserved);
 [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern IntPtr SendMessageTimeout(IntPtr h,uint m,IntPtr w,IntPtr l,uint flags,uint timeout,out IntPtr result);
}
'@
$files = @(Get-ChildItem -LiteralPath $Source -File | Where-Object Extension -In '.ttf','.otf')
foreach ($font in $files) {
 $target = Join-Path $destination $font.Name
 if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $font.FullName -Destination $target }
 elseif ((Get-FileHash -LiteralPath $target).Hash -ne (Get-FileHash -LiteralPath $font.FullName).Hash) { throw "A different version already exists: $target" }
 $joeFontTarget = Join-Path $joeFontDestination $font.Name
 if (-not (Test-Path -LiteralPath $joeFontTarget)) { Copy-Item -LiteralPath $font.FullName -Destination $joeFontTarget }
 if ((Get-FileHash -LiteralPath $joeFontTarget).Hash -ne (Get-FileHash -LiteralPath $font.FullName).Hash) { throw 'JoeOS font copy differs from source.' }
 New-ItemProperty -Path $registry -Name ($font.BaseName+' (TrueType)') -Value $target -PropertyType String -Force | Out-Null
 [JoeFontInstall]::AddFontResourceEx($target,0,[IntPtr]::Zero) | Out-Null
 if ((Get-FileHash -LiteralPath $target).Hash -ne (Get-FileHash -LiteralPath $font.FullName).Hash) { throw 'Font copy verification failed.' }
}
$result=[IntPtr]::Zero
[JoeFontInstall]::SendMessageTimeout([IntPtr]0xffff,0x001d,[IntPtr]::Zero,[IntPtr]::Zero,2,1000,[ref]$result) | Out-Null
Write-Output "Installed and verified $($files.Count) font files for this user."
