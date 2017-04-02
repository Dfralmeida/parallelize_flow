' #This flow has been annotated to better illustrate the parts of the script.
' #Any comment beginning with a # was not part of the original script.

' #Beginning of the script header.
' This script has been modified to support running jobs in parallel.
' *** Start of script for flow ExampleFlow ***

' Define constants needed for accessing files
Const ForReading = 1, ForWriting = 2, ForAppending = 8

flowStatus = 0

scriptFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\ExampleFlow.vbs"

' Update date and time variable
curDateTime = Now()

' Create timestamp used in naming the status file
timeStamp = Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime) & Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

statusFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\" & timeStamp & "_ExampleFlow_status.log"

' Initialize references to FileSystem and Shell objects
Set fileSys = Wscript.CreateObject("Scripting.FileSystemObject")
Set shell = Wscript.CreateObject("Wscript.Shell")

' Open status file
Set statusFile = fileSys.OpenTextFile(statusFilename, ForWriting, True)

' *** Start of flow ***

' Log start of flow to status file
statusFile.WriteLine("Flow STARTING...")

' #The following performs garbage collection on the signal files.
' #Otherwise repeated runs of the script would be short circuited to completion by the previous runs files.
If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/ExampleFlow/A571IL1G_AH000002_simple_status.log") Then

fileSys.DeleteFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/ExampleFlow/A571IL1G_AH000002_simple_status.log")

End If
If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000003_simple_status.log") Then

fileSys.DeleteFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000003_simple_status.log")

End If
If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000004_simple_status.log") Then

fileSys.DeleteFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000004_simple_status.log")

End If

' #End of the script header.

' #Beginning of the job invocations

dLoop = True

Do While dLoop = True

  dLoop = False

  If Not exec_A571IL1G_AH000002 and Not running_A571IL1G_AH000002 Then

    ' *** Begin Job Event ***
    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log start of job to status file
    statusFile.WriteLine("Job Alpha_A571IL1G_AH000002 STARTING " & curDate & " " & curTime)

running_A571IL1G_AH000002 = True

    '#Invocation of the job specific subscript
shell.Run("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/ExampleFlow/A571IL1G_AH000002.vbs")
    ' *** End Job Event ***

  End If

  If Not exec_A571IL1G_AH000003 and Not running_A571IL1G_AH000003 Then

    ' *** Begin Job Event ***
    ' Update date and time variables
    curDateTime = Now()
    curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
    curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

    ' Log start of job to status file
    statusFile.WriteLine("Job Beta_A571IL1G_AH000003 STARTING " & curDate & " " & curTime)

running_A571IL1G_AH000003 = True

shell.Run("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000003.vbs")
    ' *** End Job Event ***

  End If

  If Not exec_A571IL1G_AH000004 and Not running_A571IL1G_AH000004 Then
	' #This the logic which controls the job dependencies.
    If ((exec_A571IL1G_AH000002 And status_A571IL1G_AH000002 = 0) And (exec_A571IL1G_AH000003 And status_A571IL1G_AH000003 = 0)) Then

      ' *** Begin Job Event ***
      ' Update date and time variables
      curDateTime = Now()
      curDate = Left("00", 2 - Len(Month(curDateTime))) & Month(curDateTime) & "/" & Left("00", 2 - Len(Day(curDateTime))) & Day(curDateTime) & "/" & Left("0000", 4 - Len(Year(curDateTime))) & Year(curDateTime)
      curTime = Left("00", 2 - Len(Hour(curDateTime))) & Hour(curDateTime) & ":" & Left("00", 2 - Len(Minute(curDateTime))) & Minute(curDateTime) & ":" & Left("00", 2 - Len(Second(curDateTime))) & Second(curDateTime)

      ' Log start of job to status file
      statusFile.WriteLine("Job Delta_A571IL1G_AH000004 STARTING " & curDate & " " & curTime)

running_A571IL1G_AH000004 = True

shell.Run("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000004.vbs")
      ' *** End Job Event ***

    End If

  End If

' #Logic which scans for the signal file.  We only need to scan for jobs which have begun, and have not yet finished.
If ( ( Not exec_A571IL1G_AH000002 ) and running_A571IL1G_AH000002 ) Then

  '# We will loop while any job is running, and stop looping when no jobs are running, and no other jobs could be run.
  dLoop = True

  If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000002_simple_status.log") Then

    exec_A571IL1G_AH000002 = True
    Set sf = fileSys.OpenTextFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000002_simple_status.log", ForReading, True)
    strStatus = sf.ReadLine
    status_A571IL1G_AH000002 = CInt(strStatus)
    jobStatusFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\" & "A571IL1G_AH000002" & "_ExampleFlow_status.log"
    Set sf = fileSys.OpenTextFile( jobStatusFilename , ForReading, True)
    strLog = sf.ReadLine
    statusFile.WriteLine(strLog)
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000002
    End If
  End If
End If

If ( ( Not exec_A571IL1G_AH000003 ) and running_A571IL1G_AH000003 ) Then

  dLoop = True

  If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000003_simple_status.log") Then

    exec_A571IL1G_AH000003 = True
    Set sf = fileSys.OpenTextFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000003_simple_status.log", ForReading, True)
    strStatus = sf.ReadLine
    status_A571IL1G_AH000003 = CInt(strStatus)
    jobStatusFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\" & "A571IL1G_AH000003" & "_ExampleFlow_status.log"
    Set sf = fileSys.OpenTextFile( jobStatusFilename , ForReading, True)
    strLog = sf.ReadLine
    statusFile.WriteLine(strLog)
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000003
    End If
  End If
End If

If ( ( Not exec_A571IL1G_AH000004 ) and running_A571IL1G_AH000004 ) Then

  dLoop = True

  If fileSys.FileExists("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000004_simple_status.log") Then

    exec_A571IL1G_AH000004 = True
    Set sf = fileSys.OpenTextFile("C:/SAS/Config/Lev1/SchedulingServer/d-wise_dkratz/exampleflow/A571IL1G_AH000004_simple_status.log", ForReading, True)
    strStatus = sf.ReadLine
    status_A571IL1G_AH000004 = CInt(strStatus)
    jobStatusFilename = "C:\SAS\Config\Lev1\SchedulingServer\d-wise_dkratz\ExampleFlow\" & "A571IL1G_AH000004" & "_ExampleFlow_status.log"
    Set sf = fileSys.OpenTextFile( jobStatusFilename , ForReading, True)
    strLog = sf.ReadLine
    statusFile.WriteLine(strLog)
    If flowStatus = 0 Then
      flowStatus = status_A571IL1G_AH000004
    End If
  End If
End If

' #Pause operation of the script for 15 seconds
Wscript.Sleep(15000)
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