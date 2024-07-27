#Requires -Version 5.1
if ( -not ( Get-Module -ListAvailable -Name ps2exe ) ) { 
  Install-Module -Name ps2exe
}
invoke-ps2exe -inputfile sit-stand.ps1 -iconfile sit-stand.ico -noConsole -outputfile sit-stand.exe 
