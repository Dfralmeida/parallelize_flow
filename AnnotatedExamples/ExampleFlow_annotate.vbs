' #This flow has been annotated to better illustrate the parts of the script.
' #Any comment beginning with a # was not part of the original script.

' #Beginning of the script header.
' *** Start of script for flow ExampleFlow ***

' Define constants needed for accessing files
Const ForReading = 1, ForWriting = 2, ForAppending = 8

flowStatus = 0

scriptFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\ExampleFlow.vbs"

' Update date and time variable
curDateTime = Now()

' Create timestamp used in naming the status file
timeStamp = Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime) & Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

statusFilename = "C:\SAS\Config\Lev1\SchedulingServer\dkratz\ExampleFlow\" & timeStamp & "_ExampleFlow_status.log"

' Initialize references to FileSystem and Shell objects
Set fileSys = Wscript.CreateObject("Scripting.FileSystemObject")
Set shell = Wscript.CreateObject("Wscript.Shell")

' Open status file
Set statusFile = fileSys.OpenTextFile(statusFilename, ForWriting, True)

' *** Start of flow ***

' Log start of flow to status file
statusFile.WriteLine("Flow STARTING...")
' #End of the script header.

' #Beginning of the job invocations
dLoop = True

Do While dLoop = True

  dLoop = False

  If Not exec_A571IL1G_AH000002 Then

	' # If this job had dependencies, the logic for them would be controlled here.
    ' *** Begin Job Event ***

    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log start of job to status file
    statusFile.WriteLine("Job Alpha_A571IL1G_AH000002 STARTING " & curDate & " " & curTime)

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

    dLoop = True

    ' Update flow exit code
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000002
    End If

    ' *** End Job Event ***

  End If

  If Not exec_A571IL1G_AH000003 Then

    ' *** Begin Job Event ***

    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log start of job to status file
    statusFile.WriteLine("Job Beta_A571IL1G_AH000003 STARTING " & curDate & " " & curTime)

    ' Enable error handling
    On Error Resume Next

    ' Execute job
    errorLevel = shell.Run("C:\SAS\Config\Lev1\SASApp\BatchServer\sasbatch.bat -log C:\SAS\Config\Lev1\SASApp\BatchServer\Logs\ExampleFlow_Beta_#Y.#m.#d_#H.#M.#s.log -batch -noterminal -logparm ""rollover=session""  -sysin C:\SAS\Config\Lev1\SASApp\SASEnvironment\SASCode\Jobs\Beta.sas", , True)

    If Err.Number <> 0 Then
      status_A571IL1G_AH000003 = Err.Number
      Err.Clear
    Else
      status_A571IL1G_AH000003 = errorLevel
    End If

    ' Disable error handling
    On Error Goto 0

    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log completion of job and exit code to status file
    statusFile.WriteLine("Job Beta_A571IL1G_AH000003 COMPLETE " & curDate & " " & curTime & " status=" & status_A571IL1G_AH000003 & ".")

    ' Set flag indicating that job has executed
    exec_A571IL1G_AH000003 = True

    dLoop = True

    ' Update flow exit code
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000003
    End If

    ' *** End Job Event ***

  End If

  If Not exec_A571IL1G_AH000004 Then

	' #This is an example of the job dependency logic.
    If ((exec_A571IL1G_AH000002 And status_A571IL1G_AH000002 = 0) And (exec_A571IL1G_AH000003 And status_A571IL1G_AH000003 = 0)) Then

      ' *** Begin Job Event ***

      ' Update date and time variables
      curDateTime = Now()
      curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
      curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

      ' Log start of job to status file
      statusFile.WriteLine("Job Delta_A571IL1G_AH000004 STARTING " & curDate & " " & curTime)

      ' Enable error handling
      On Error Resume Next

      ' Execute job
      errorLevel = shell.Run("C:\SAS\Config\Lev1\SASApp\BatchServer\sasbatch.bat -log C:\SAS\Config\Lev1\SASApp\BatchServer\Logs\ExampleFlow_Delta_#Y.#m.#d_#H.#M.#s.log -batch -noterminal -logparm ""rollover=session""  -sysin C:\SAS\Config\Lev1\SASApp\SASEnvironment\SASCode\Jobs\Delta.sas", , True)

      If Err.Number <> 0 Then
        status_A571IL1G_AH000004 = Err.Number
        Err.Clear
      Else
        status_A571IL1G_AH000004 = errorLevel
      End If

      ' Disable error handling
      On Error Goto 0

      ' Update date and time variables
      curDateTime = Now()
      curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
      curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

      ' Log completion of job and exit code to status file
      statusFile.WriteLine("Job Delta_A571IL1G_AH000004 COMPLETE " & curDate & " " & curTime & " status=" & status_A571IL1G_AH000004 & ".")

      ' Set flag indicating that job has executed
      exec_A571IL1G_AH000004 = True

      dLoop = True

      ' Update flow exit code
      If flowStatus = 0 Then
        flowStatus = status_A571IL1G_AH000004
      End If

      ' *** End Job Event ***

    End If

  End If

Loop
' #Ending of the job invocations

' #Beginning of the script footer
' Update date and time variables
curDateTime = Now()
curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

' Log completion of flow and exit code to status file
statusFile.WriteLine("Flow ExampleFlow COMPLETE " & curDate & " " & curTime & " status=" & flowStatus & ".")

' Close status file
statusFile.Close

' Exit flow and return status
Wscript.Quit(flowStatus)

' *** End of script for flow ExampleFlow ***
' #Ending of the script footer