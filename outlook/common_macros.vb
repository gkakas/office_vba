Option Explicit
'published in github https://github.com/gkakas/office_vba.git

Sub UnifiedInbox()
    DoSearch ("folder:Inbox ")
End Sub
Sub UnifiedDrafts()
    DoSearch ("folder:Draft")
End Sub
Sub UnifiedArchive()
    DoSearch ("folder:Archive")
End Sub
Sub UnifiedSent()
    DoSearch ("folder:Sent")
End Sub
Sub DoSearch(ByVal terms As String)
    Dim txtSearch  As String
    Dim myOlApp As New Outlook.Application
    Set Application.ActiveExplorer.CurrentFolder = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
    txtSearch = terms
    myOlApp.ActiveExplorer.Search txtSearch, olSearchScopeAllFolders
    Set myOlApp = Nothing
End Sub
Sub UnifiedInOut()
    DoSearch ("(folder:Inbox OR folder:Sent)")
End Sub
Sub ArchiveSelectedMessages()

    Dim objFolder As Outlook.MAPIFolder
    
    Set objFolder = FindMAPIFolder("Archive")
    
    If Not objFolder Is Nothing Then
        MoveSelectedMessagesToFolder objFolder, True
    End If
    
    Set objFolder = Nothing

End Sub
Function FindMAPIFolder(ByVal name As String) As Outlook.MAPIFolder

    Dim objFolder As Outlook.MAPIFolder, objInbox As Outlook.MAPIFolder

    Dim objNS As Outlook.NameSpace
    Dim i As Integer
    
    name = UCase(name)

    Set objNS = Application.GetNamespace("MAPI")
    Set objInbox = objNS.GetDefaultFolder(olFolderInbox)

    Set objFolder = Nothing
    
    For i = 1 To objInbox.Parent.Folders.Count

        If UCase(objInbox.Parent.Folders(i).name) = name Then
        
            Set objFolder = objInbox.Parent.Folders(i)
            Exit For
        End If
    
    Next
    
    Set objInbox = Nothing
    Set objNS = Nothing
    Set FindMAPIFolder = objFolder

End Function
'move selected messages to a folder
Private Sub MoveSelectedMessagesToFolder(ByVal objFolder As Outlook.MAPIFolder, ByVal MarkAsRead As Boolean)

    Dim objItem As Object
    'Dim objMailItem As Outlook.MailItem
    If Application.ActiveExplorer.Selection.Count = 0 Then
        Exit Sub
    End If
 

    If objFolder Is Nothing Then

        MsgBox "This folder doesn’t exist!", vbOKOnly + vbExclamation, "INVALID FOLDER"
        Exit Sub

    End If
    
    For Each objItem In Application.ActiveExplorer.Selection

        If objFolder.DefaultItemType = olMailItem Then
            'If objItem.Class = olMail Then
                If MarkAsRead Then objItem.UnRead = False
                On Error Resume Next
                objItem.Move objFolder
                On Error GoTo 0
            'End If
        End If
    
    Next

    Set objItem = Nothing
    Set objFolder = Nothing

End Sub
Public Function GetCalendarView() As Outlook.CalendarView
    Dim oExpl As Outlook.Explorer
    Dim oView As Outlook.View
 
    Set oExpl = Application.ActiveExplorer
    Set oView = oExpl.CurrentView
    If oView.ViewType = olCalendarView Then
        Set GetCalendarView = oExpl.CurrentView
    Else
        Set GetCalendarView = Nothing
    End If

End Function
Public Sub ReserveCalendarTime()
    On Error Resume Next
    
    Dim oCalView As Outlook.CalendarView
    Dim oFolder As Outlook.Folder
    Dim oNameSpace As Outlook.NameSpace
    Dim oCalendar As Outlook.MAPIFolder
    Const datNull As Date = #1/1/1900#

    Dim dStart As Date
    Dim dEnd As Date
    Dim oAppointment As Outlook.AppointmentItem
    
    
    Set oNameSpace = Application.GetNamespace("MAPI")
    Set oCalendar = oNameSpace.GetDefaultFolder(olFolderCalendar)

    Set oCalView = GetCalendarView()
    If oCalView Is Nothing Then Exit Sub
    
    Set oFolder = oCalendar.Folders.Parent
    
    If oFolder Is Nothing Then
        Set oFolder = Application.ActiveExplorer.CurrentFolder
    End If
    
    dStart = oCalView.SelectedStartTime
    dEnd = oCalView.SelectedEndTime
    
    Set oAppointment = oFolder.Items.Add("IPM.Appointment")
    
    If dStart <> datNull And dEnd <> datNull Then
        oAppointment.Start = dStart
        oAppointment.End = dEnd
    End If
    oAppointment.ReminderSet = False
    
    oAppointment.Subject = "Reserved time slot"
    oAppointment.BusyStatus = olTentative
    oAppointment.Categories = "Reserved"
    oAppointment.Display
    
End Sub
Public Sub ShowMyCalendars()
    Dim xc As Folder
    Set xc = OpenCalendarDisplay
    SelectCalendar xc, "Calendar"
End Sub

Private Function OpenCalendarDisplay() As Folder
    Dim xCalendar As Folder
    Set xCalendar = Outlook.Session.GetDefaultFolder(olFolderCalendar)
    xCalendar.Display
    Set OpenCalendarDisplay = xCalendar
    
End Function
Private Sub SelectCalendar(ByVal calendar As Folder, ByVal name As String, Optional ByVal doSelect As Boolean = True)
    Dim mCalendar As CalendarModule
    Dim nGrp As NavigationGroup
    Dim nFldr As NavigationFolder
    Dim i As Integer, j As Integer
    Set mCalendar = calendar.GetExplorer().NavigationPane.Modules.GetNavigationModule(olModuleCalendar)
    For i = 1 To mCalendar.NavigationGroups.Count
        Set nGrp = mCalendar.NavigationGroups.Item(i)
        For j = 1 To nGrp.NavigationFolders.Count
            Set nFldr = nGrp.NavigationFolders.Item(j)
            If UCase(nFldr) = UCase(name) Then
                nFldr.IsSelected = doSelect
                nFldr.IsSideBySide = False
            End If
        Next
        
    Next
End Sub

