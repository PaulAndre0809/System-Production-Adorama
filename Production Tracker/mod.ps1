[CmdletBinding()]
param(
    [switch]$Visible
)

$ErrorActionPreference = "Stop"

& (Join-Path $PSScriptRoot "embed_vba.ps1") -Visible:$Visible
