Const logFilePath = "C:\Temp\Calendar Tests\Outlook.log"
Const sharedCalendarPrimarySmtp = "<shared calendar primary SMTP address>"
Const sharedCalendarDisplayName = "<Display name of shared calendar in Outlook>" ' This shouldn't be required, it is used for a fallback method to connect to the shared calendar

Public Sub CreateAppointments()
    ' Create a number of appointments
    Dim primaryCalendarFolder As Folder
    Dim sharedCalendarOwner As Recipient
    Dim sharedCalendarFolder As Folder
    Dim iAptCount As Integer
    Dim oInterval As Date
    Dim oStart As Date
    
    Set primaryCalendarFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar)
    
    Set sharedCalendarOwner = Application.GetNamespace("MAPI").CreateRecipient(sharedCalendarPrimarySmtp)
    sharedCalendarOwner.Resolve
    Set sharedCalendarFolder = Application.GetNamespace("MAPI").GetSharedDefaultFolder(sharedCalendarOwner, olFolderCalendar)
    
    If (sharedCalendarFolder Is Nothing) Then
        Set objExpCal = primaryCalendarFolder.GetExplorer
        Set objNavMod = objExpCal.NavigationPane.Modules.GetNavigationModule(olModuleCalendar)
        
        For Each objNavGroup In objNavMod.NavigationGroups
            For Each objNavFolder In objNavGroup.NavigationFolders
                If objNavFolder = sharedCalendarDisplayName Then
                    Set sharedCalendarFolder = objNavFolder.Folder
                    Exit For
                End If
            Next
            If Not sharedCalendarFolder Is Nothing Then Exit For
        Next
    End If
    
    If (sharedCalendarFolder Is Nothing) Then Exit Sub
    
    iAptCount = 2
    oInterval = DateAdd("n", 60, CDate(0))
    oStart = CDate(Format(Now, "dd/mmm/yyyy hh:00:00"))
    
    For i = 1 To iAptCount
        Dim oApt As AppointmentItem
        Dim subject As String
        
        subject = "Test " & i
        Set oApt = sharedCalendarFolder.Items.Add(olAppointmentItem)
        With oApt
            .Start = oStart
            .Duration = 30
            .subject = subject
            .Save
        End With
        Log " Created appointment: " + subject
        oStart = DateAdd("n", 60, oStart)
        Wait ("0:00:10")
        oApt.Delete
        Log " Deleted appointment: " + subject
        Wait ("0:00:02")
    Next
End Sub

Private Sub Wait(delay As String)
    WaitUntil = Now + TimeValue(delay)
    Do Until Now > WaitUntil
        DoEvents
    Loop
End Sub

Private Sub Log(data As String)
    Dim fso As Object 'Scripting.FileSystemObject
    Dim logStream As Object 'TextStream
    Dim logData As String
    
    logData = Format(Now, "dd/mmm/yyyy hh:nn:ss") + data
    Set fso = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set logStream = fso.OpenTextFile(logFilePath, 8, True)
    logStream.WriteLine logData
    logStream.Close
    Debug.Print logData
End Sub
