Option Strict On
Option Explicit On

Imports System.IO
Imports System.Configuration
Imports System.Xml

Public Class StatusFileMover

    Private mstrStatusUpdateFileFolder As String = ""
    Private mstrStatusUpdateFileSuccessFolder As String = ""
    Private mstrStatusUpdateFileFailureFolder As String = ""

    Private mstrDataFreightFileFolder As String = ""
    Private mstrDataFreightFileSuccessFolder As String = ""
    Private mstrDataFreightFileFailureFolder As String = ""

    Private mstrEventLogSource As String

    Private mblnServiceStartedSuccessfully As Boolean = True

    Private IntervalTimer As System.Threading.Timer
    Protected Overrides Sub OnStart(ByVal args() As String)

        ' Create the Event log if not existing
        Dim myLog As New EventLog()
        If Not myLog.SourceExists("StatusUpdateFileMover") Then
            myLog.CreateEventSource("StatusUpdateFileMover", "Status Update File Mover Log")
        End If

        ' Event log that the service is starting
        myLog.Source = "StatusUpdateFileMover"
        myLog.WriteEntry("Status Update File Mover Log", "Service Started on  " &
                                    Date.Today.ToShortDateString & " " &
                                    CStr(TimeOfDay),
                                    EventLogEntryType.Information)



        Try


            ' Get the app settings 
            mstrStatusUpdateFileFolder = ConfigurationManager.AppSettings.Get("StatusUpdateFileFolder")
            mstrStatusUpdateFileSuccessFolder = ConfigurationManager.AppSettings.Get("StatusUpdateSuccessFolder")
            mstrStatusUpdateFileFailureFolder = ConfigurationManager.AppSettings.Get("StatusUpdateFailureFolder")

            mstrDataFreightFileFolder = ConfigurationManager.AppSettings.Get("DataFreightFileFolder")
            mstrDataFreightFileSuccessFolder = ConfigurationManager.AppSettings.Get("DataFreightSuccessFolder")
            mstrDataFreightFileFailureFolder = ConfigurationManager.AppSettings.Get("DataFreightFailureFolder")

            mstrEventLogSource = ConfigurationManager.AppSettings.Get("EventLogSource")

            ' If any of the above are blank or empty then the service cannot function correctly
            If mstrStatusUpdateFileFolder.Trim = "" Or mstrStatusUpdateFileSuccessFolder.Trim = "" Or mstrStatusUpdateFileFailureFolder.Trim = "" Or
                    mstrDataFreightFileFolder.Trim = "" Or mstrDataFreightFileSuccessFolder.Trim = "" Or mstrDataFreightFileFailureFolder.Trim = "" Then
                mblnServiceStartedSuccessfully = False

                myLog.Source = "StatusUpdateFileMover"
                myLog.WriteEntry("Status Update File Mover Log", "App Settings are not correctly configured. Check the config file.",
                                    EventLogEntryType.Information)
            Else
                ' Start the timer event listener
                Dim tsInterval As TimeSpan = New TimeSpan(0, 0, 5)
                IntervalTimer = New System.Threading.Timer(New System.Threading.TimerCallback(AddressOf IntervalTimer_Elapsed), Nothing, tsInterval, tsInterval)

            End If


            ' Restart the timer?
            'IntervalTimer.Change(5, 5)

        Catch ex As Exception
            'IntervalTimer.Change(5, 5)
            myLog.Source = "StatusUpdateFileMover"
            myLog.WriteEntry("Status Update File Mover Log", "An error has occured!" & vbCrLf & ex.Message.ToString,
                                    EventLogEntryType.Error)

        End Try


    End Sub

    Protected Overrides Sub OnStop()

        ' Disable the timer
        If mblnServiceStartedSuccessfully = True Then
            IntervalTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)

            IntervalTimer.Dispose()
            IntervalTimer = Nothing
        End If


        ' Tell the event log we are stopping
        Dim myLog As New EventLog()
        myLog.Source = "StatusUpdateFileMover"
        myLog.WriteEntry("Status Update File Mover Log", "Service Stopped on  " &
                                    Date.Today.ToShortDateString & " " &
                                    CStr(TimeOfDay),
                                    EventLogEntryType.Information)



    End Sub

    Private Sub IntervalTimer_Elapsed(ByVal state As Object)

        Try
            ' Testing only 
            'Dim myLog As New EventLog()
            'myLog.Source = "StatusUpdateFileMover"
            'myLog.WriteEntry("Status Update File Mover Log", "Timer Event entered" & " on " &
            '                                                    Date.Today.ToShortDateString & " " &
            '                                                    CStr(TimeOfDay),
            '                                                    EventLogEntryType.Information)

            ' Pause the timer?
            IntervalTimer.Change(System.Threading.Timeout.Infinite, System.Threading.Timeout.Infinite)



            Dim di As New DirectoryInfo(mstrStatusUpdateFileFolder)

            ' Look at the status file folder for xml files.
            Dim fiArr As FileInfo() = di.GetFiles("*.xml")

            'Dim testfile As System.IO.StreamWriter
            'testfile = My.Computer.FileSystem.OpenTextFileWriter("C:\Temp\Event.log", True)


            Dim fri As FileInfo
            For Each fri In fiArr


                ' if a file exists then copy it to the data freight file folder
                System.IO.File.Copy(fri.FullName, Path.Combine(mstrDataFreightFileFolder, fri.Name), True)




                ' Event log - File copied to Data Freight folder
                'myLog = New EventLog()
                Dim myLog As New EventLog()
                myLog.Source = "StatusUpdateFileMover"
                myLog.WriteEntry("Status Update File Mover Log", "Status Update File: " & fri.Name & " moved to folder: " & mstrDataFreightFileFolder & " on " &
                                                                Date.Today.ToShortDateString & " " &
                                                                CStr(TimeOfDay),
                                                                EventLogEntryType.Information)


                ' Poll the data freight file folder until the file disappears or a time value passes.
                Dim dteStartTime As Date = DateTime.Now
                Dim seconds As Long = 0
                Dim blnFileExists As Boolean = True
                Do
                    blnFileExists = File.Exists(Path.Combine(mstrDataFreightFileFolder, fri.Name))
                    Dim dteTimeNow As Date = DateTime.Now
                    seconds = DateDiff("s", dteStartTime, dteTimeNow)

                Loop Until seconds > 30 Or blnFileExists = False


                ' If the file disappears from the data freight folder then it will be either in the success or failure folder.
                If blnFileExists = False Then
                    ' Now the file is no longer in the root folder.
                    ' It has either been moved to the failure or success folder

                    Dim folderName As String = ""
                    Dim newFolderName As String = ""
                    If File.Exists(Path.Combine(mstrDataFreightFileFailureFolder, fri.Name)) Then
                        ' Log that the file has moved to the failure folder
                        folderName = mstrDataFreightFileFailureFolder
                        newFolderName = mstrStatusUpdateFileFailureFolder
                    Else

                        ' It must be in the success folder. Log this.
                        folderName = mstrDataFreightFileSuccessFolder
                        newFolderName = mstrStatusUpdateFileSuccessFolder



                        ' The event log should be updated here

                        ' Poll the event log for an match of the file and grab the AIM ref from the log?
                        Dim strEventTitle As String = mstrEventLogSource  ' Pick this up from Config
                        Dim log As EventLog = New EventLog(strEventTitle)


                        Dim blnTestOnlyOnce As Boolean = True
                        For counter As Integer = 1 To 100


                            Dim msg As String = log.Entries(log.Entries.Count - counter).Message

                            If Strings.Left(msg, 12) <> "Notification" Then


                                Dim pos2 As Integer = InStr(msg.ToUpper, fri.Name.ToUpper)
                                'testfile.WriteLine(fri.Name.ToUpper & " - " & msg.ToUpper & vbCrLf & pos2 & vbCrLf & vbCrLf)

                                If InStr(msg.ToUpper, fri.Name.ToUpper) > 0 Then
                                    ' The file name that was imported is contained in the event viewer message
                                    ' We need to now parse this message for the Aim reference.

                                    ' Name exists in msg
                                    myLog.Source = "StatusUpdateFileMover"
                                    myLog.WriteEntry("Status Update File Mover Log", "File name " & fri.Name & " found in log " & " on " &
                                                                        Date.Today.ToShortDateString & " " &
                                                                        CStr(TimeOfDay),
                                                                        EventLogEntryType.Information)


                                    'Job | Shipment | Agent Ref: 'AIM117576 | AIM117576-1 | EPOSATGB202100000008' updated. Container Count: 1
                                    Dim pos As Integer = InStr(msg.ToUpper, "JOB | SHIP")
                                    'testfile.WriteLine("Pos: " & pos.ToString & vbCrLf)
                                    If pos > 0 Then

                                        msg = Strings.Mid(msg, pos, Len(msg))
                                        'testfile.WriteLine("msg: " & msg & vbCrLf)
                                        pos = InStr(msg.ToUpper, "FILE '")
                                        'msg = Strings.Mid(msg, pos, Len(msg))
                                        'testfile.WriteLine("pos: " & pos.ToString & vbCrLf)

                                        If pos > 0 Then
                                            msg = Strings.Left(msg, pos - 1)
                                            'testfile.WriteLine("msg: " & msg & vbCrLf)

                                            Dim msgParts() As String = Split(msg, "|")

                                            Dim aimReference As String = ""
                                            Dim eci As String = ""
                                            If msgParts.Length >= 4 Then
                                                aimReference = msgParts(3).Trim

                                                If msgParts.Length >= 5 Then
                                                    '156800054' updated. Container Count: 2
                                                    pos = InStr(msgParts(4), "'")
                                                    If pos > 0 Then
                                                        eci = Strings.Left(msgParts(4), pos - 1)
                                                    End If
                                                End If
                                            End If

                                            'testfile.WriteLine("Aim: " & aimReference & vbCrLf)
                                            'testfile.WriteLine("eci: " & eci & vbCrLf)


                                            ' Write out the xml file to the success folder
                                            ' If we get a match then we can create an xml file with the AIM Ref?
                                            ' <StatusUpdate>
                                            '   <CATTS_File>StatusUpdateCA_123456.xml</CATTS_File>
                                            '   <AIMReference>AIM123456-1</AimReference>
                                            '   <ECI>121212121</ECI>
                                            ' </StatusUpdate>

                                            ' Save this to the folder.
                                            Dim strFileName As String = Path.Combine(newFolderName, eci & ".xml")

                                            myLog.Source = "StatusUpdateFileMover"
                                            myLog.WriteEntry("Status Update File Mover Log", "Creating AIM Ref File: " & strFileName & " on " &
                                                                Date.Today.ToShortDateString & " " &
                                                                CStr(TimeOfDay),
                                                                EventLogEntryType.Information)


                                            ' Now write the data out based on the class.
                                            Dim settings As XmlWriterSettings = New XmlWriterSettings()
                                            settings.Indent = True

                                            ' Create XmlWriter.
                                            Using writer As XmlWriter = XmlWriter.Create(strFileName, settings)

                                                ' Begin writing.
                                                writer.WriteStartDocument()

                                                ' Write root element
                                                writer.WriteStartElement("StatusUpdate") ' This may not be needed


                                                ' Catts_File
                                                writer.WriteElementString("CATTS_File", fri.Name.ToString) ' Original message (outbound)?
                                                writer.WriteElementString("AIMReference", aimReference)
                                                writer.WriteElementString("ECI", eci)


                                                ' End Root element
                                                writer.WriteEndElement()


                                                ' End Document
                                                writer.WriteEndDocument()
                                            End Using

                                        End If

                                    End If

                                    ' Exit the loop
                                    Exit For

                                End If

                            End If
                        Next



                        folderName = mstrDataFreightFileSuccessFolder
                        newFolderName = mstrStatusUpdateFileSuccessFolder
                    End If

                    myLog.Source = "StatusUpdateFileMover"
                    myLog.WriteEntry("Status Update File Mover Log", "Status Update File: " & fri.Name & " moved to folder: " & folderName & " on " &
                                                                Date.Today.ToShortDateString & " " &
                                                                CStr(TimeOfDay),
                                                                EventLogEntryType.Information)


                    ' Move the original file to the status update failure or success folder.
                    If File.Exists(Path.Combine(newFolderName, fri.Name)) Then
                        ' Delete if exists
                        File.Delete(Path.Combine(newFolderName, fri.Name))
                    End If


                    fri.MoveTo(Path.Combine(newFolderName, fri.Name))

                    ' Event log - File has been archived to the Success/Failure folder
                    myLog.Source = "StatusUpdateFileMover"
                    myLog.WriteEntry("Status Update File Mover Log", "Status Update File: " & fri.Name & " moved to folder: " & newFolderName & " on " &
                                                                Date.Today.ToShortDateString & " " &
                                                                CStr(TimeOfDay),
                                                                EventLogEntryType.Information)



                Else
                    ' File has not been processed by Data Freight
                    myLog.Source = "StatusUpdateFileMover"
                    myLog.WriteEntry("Status Update File Mover Log", "Status Update File: " & fri.Name & " not processed by Data Freight (" &
                                                                Date.Today.ToShortDateString & " " &
                                                                CStr(TimeOfDay) & ")",
                                                                EventLogEntryType.Information)


                End If

            Next fri

            'testfile.Close()

            ' Restart the timer?
            Dim tsInterval As TimeSpan = New TimeSpan(0, 0, 5)
            'IntervalTimer.Change(0, 5000)
            IntervalTimer.Change(tsInterval, tsInterval)


        Catch ex As Exception
            ' Restart the timer?
            Dim tsInterval As TimeSpan = New TimeSpan(0, 0, 5)
            IntervalTimer.Change(tsInterval, tsInterval)

            Dim myLog As New EventLog()
            myLog.Source = "StatusUpdateFileMover"
            myLog.WriteEntry("Status Update File Mover Log", "An error has occured!" & vbCrLf & ex.Message.ToString,
                                    EventLogEntryType.Error)


        End Try

    End Sub


End Class
