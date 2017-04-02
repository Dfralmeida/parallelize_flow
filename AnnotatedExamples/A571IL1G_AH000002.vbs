' #This flow has been annotated to better illustrate the parts of the script.
' #Any comment beginning with a # was not part of the original script.

' #Beginning of the modified header.
' *** Start of script for flow ExampleFlow ***

' Define constants needed for accessing files
Const ForReading = 1, ForWriting = 2, ForAppending = 8

flowStatus = 0

scriptFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\ExampleFlow.vbs"

' Update date and time variable
curDateTime = Now()

' Create timestamp used in naming the status file
timeStamp = Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime) & Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

statusFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\" & "A571IL1G_AH000002" & "_ExampleFlow_status.log"

' Initialize references to FileSystem and Shell objects
Set fileSys = Wscript.CreateObject("Scripting.FileSystemObject")
Set shell = Wscript.CreateObject("Wscript.Shell")

' Open status file
Set statusFile = fileSys.OpenTextFile(statusFilename, ForWriting, True)
' #End of the modified header.

' #Beginning of the modified job invocation


    ' Enable error handling
    On Error Resume Next

    ' Execute job
    errorLevel = shell.Run("C:\SAS\Config\Lev1\SASApp\BatchServer\sasbatch.bat -log C:\SAS\Config\Lev1\SASApp\BatchServer\Logs\ExampleFlow_Alpha_#Y.#m.#d_#H.#M.#s.log -batch -noterminal -logparm ""rollover=session""  -sysin C:\SAS\Config\Lev1\SASApp\SASEnvironment\SASCode\Jobs\Alpha.sas", , True)

    If Err.Number <> 0 Then
      status_A571IL1G_AH000002 = Err.Number
      Err.Clear
    Else
      status_A571IL1G_AH000002 = errorLevel
    End If

    ' Disable error handling
    On Error Goto 0

    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log completion of job and exit code to status file
    statusFile.WriteLine("Job Alpha_A571IL1G_AH000002 COMPLETE " & curDate & " " & curTime & " status=" & status_A571IL1G_AH000002 & ".")

    ' Set flag indicating that job has executed
    exec_A571IL1G_AH000002 = True

	' #Note that we don't need this variable, but since the subscripts do not contain a loop, leaving it in does no harm.
    dLoop = True

    ' Update flow exit code
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000002
    End If
'  #End of the modified job invocation

'  #Beginning of the injected logic to create a signal file.
simpleStatusFilename = "C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/ExampleFlow/A571IL1G_AH000002_simple_status.log"

' Open status file
Set simpleStatusFile = fileSys.OpenTextFile(simpleStatusFilename, ForWriting, True)
' Log completion of job and exit code to status file
simpleStatusFile.WriteLine(status_A571IL1G_AH000002)
' Close status file
simpleStatusFile.Close
'  #End of the injected logic to create a signal file.