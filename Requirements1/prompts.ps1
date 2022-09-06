# Tassy Estinphon , #001392883
# Scripting and Automation Task 1

[String]$Req = "C:\Users\LabAdmin\Desktop\Requirements1"
#Execution 
Try{
	Do{
$Input = Read-Host -Prompt 'Which project do you wish to start'
Switch -Regex($input)
{
#List files with Log extension and output to a txt file
1 {Get-ChildItem -Path $Req | Select -Property Name | Where-Object OutputType -EQ .log | Out-File -Append $Req\Dailylog.txt }
#Listing files with Requirements folder and outputting to a txt file 
2 {Get-ChildItem -Path $Req | Sort-Object Name | Format-Table | Out-File -Append $Req\C916context.txt }
#Shows current CPU & Memory Usage
3 {$CPU = Get-Counter '\Processor(*)\% Processor Time' | Select-Object -ExpandProperty countersamples | Select-Object -Property instancename,cookedvalue
 $MEM = Get-Counter '\Memory\Committed Bytes' | Select-Object -ExpandProperty countersamples | Select-Object -Property instancename,cookedvalue
 Write-Host "CPU Usage:" $CPU | Sort-Object -Property cookedvalue -Descending
 Write-Host "Memory Usage:" $MEM | Sort-Object -Property cookedvalue -Descending
 }
#List all current running process
4 {Get-Process | Sort-Object VirtualMemorySize -Descending | Out-GridView}
#Ends the prompt
5 {Return}
#result if input is lt 1 or gt 5
default {"N/A"}
}
#Makes the prompt run until host selects 5
}Until($input -eq 5)
}
 Catch [System.OutofMemoryException]
{Write-Host -ForegroundColor $_.exceptionmessage
}