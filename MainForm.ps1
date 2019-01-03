[GC]::Collect | Out-Null

$script:ScriptDir = Split-Path -parent $MyInvocation.MyCommand.Path 

Import-Module -Name $ScriptDir\PIIFinding -Force

Show-Menu