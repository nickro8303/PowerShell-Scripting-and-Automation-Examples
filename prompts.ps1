
<#
.Synopsis
   This script prompts the user to complete the process.
.DESCRIPTION
 A.  	Create a PowerShell script named “prompts.ps1” within the “Requirements1” folder. 
	For the first line, create a comment and include your first and last name along 
	with your student ID.

	Note: The remainder of this task should be completed within the same script file,
	prompts.ps1.
 B.  	Create a “switch” statement that continues to prompt a user by doing each of the 
	following activities, until a user presses key 5:

	1.   	Using a regular expression, list files within the Requirements1 folder, with the 
		.log file extension and redirect the results to a new file called “DailyLog.txt”
		within the same directory without overwriting existing data. Each time the user 
		selects this prompt, the current date should precede the listing. (User presses key 1.)

	2.  	List the files inside the “Requirements1” folder in tabular format, sorted in 
		ascending alphabetical order. Direct the output into a new file called 
		“C916contents.txt” found in your “Requirements1” folder. (User presses key 2.)

	3. 	Use counters to list the current CPU % Processor Time and physical memory usage. 
		Collect 4 samples with each sample being 5 seconds intervals. (User presses key 3.) 

	4.  	List all the different running processes inside your system. Sort the output by 
		processor time in seconds greatest to least, and display it in grid format. 
		(User presses key 4.) 

	5.  	Exit the script execution. (User presses key 5.)

 C.  	Apply exception handling using try-catch for System.OutOfMemoryException.

#>
clear

function MyFunction {

	($param1, $param2)

    $param1 = Get-Date | Out-File -FilePath $PSScriptRoot\dailylog.txt -Append
    $param2 = Get-ChildItem $PSScriptRoot -Filter *.log | Out-File -FilePath $PSScriptRoot\dailylog.txt -Append

}


try {

clear
 
Do {
$input = (Read-Host "please enter a number from 1-4 to 

1 Get the date and add it to DailyLog.txt.
2 List the files inside the “Requirements1” folder.
3 List the current CPU % Processor Time and physical memory usage. 
4 List all the different running processes inside your system.

or enter 5 to quit" );
 
switch ($input) {
   1 {MyFunction; break}
   2 {Get-ChildItem $PSScriptRoot | Sort-Object -Property Name | Out-File -FilePath $PSScriptRoot\C916contents.txt}
   3 {Get-Counter -Counter "\Processor(_Total)\% Processor Time", "\Memory\Committed bytes" -SampleInterval 5 -MaxSamples 4}
   4 {Get-Process | Out-GridView}
   5 {" Cya!"; break}
   6 {
    "throwing a System.OutOfMemoryException"
   $exception = New-Object System.OutOfMemoryException
   throw $exception
   break
   }

   default {"Something else happened, try again!"; break}

}
} until ($input -eq 5)

}
catch [System.OutOfMemoryException] {
    Write-Host "A system is out of memory exception has occured."
}
