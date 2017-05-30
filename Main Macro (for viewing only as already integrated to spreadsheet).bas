Attribute VB_Name = "Cal_Export"
Option Explicit

Public Sub Main()

'''''''''''''''''''''''''''''''''
'             SETUP             '
'''''''''''''''''''''''''''''''''
   
   Debug.Print "Setting up Outlook..."
   
   Dim olApp As Outlook.Application
        Dim olAppt As Outlook.AppointmentItem
        Dim Folders As Outlook.Folder
        Dim subFolder, CalFolder As Outlook.MAPIFolder
        Dim olNS As Outlook.Namespace
    
    Dim blnCreated As Boolean
    Dim arrCal As String
    Dim i, j As Long
    Dim bool As Boolean
    Dim ImportBlanks As Boolean
    Dim SkipBlanks As Boolean
    Dim Response As Variant
    
    'Initiate instance of Outlook
    On Error Resume Next
    Set olApp = GetObject("", "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
        On Error GoTo 0
        If olApp Is Nothing Then
            MsgBox "Outlook is not available! Please open and retry again..." 'make dynamic
            Exit Sub
        End If
    End If
        
    'On Error GoTo Err_Execute
    On Error GoTo 0
    
    Set olNS = olApp.GetNamespace("MAPI")
    Set CalFolder = olNS.GetDefaultFolder(olFolderCalendar)
     
'''''''''''''''''''''''''''''''''
'           Routine             '
'''''''''''''''''''''''''''''''''
               
     i = 2 'omit title row
     SkipBlanks = False
     
     'Loop rows until Cells(i, 1) is empty
     Do Until Trim(Cells(i, 1).Value) = ""
        
        'Error handling for blank cells.
        If Cells(i, 8) = "" And ImportBlanks = False And SkipBlanks = False Then
            Response = MsgBox("You have not selected whether you want to import this item..." _
                & Chr(10) & "Import - " & Cells(i, 2).Value & "?", vbYesNo)
            
            If Response = vbYes Then
                Response = MsgBox("Do you want to import all future blank items?", vbYesNo)
                
                If Response = vbYes Then ImportBlanks = True
                GoTo ImportItem
            Else
                Response = MsgBox("Do you want to ignore all future blank items?", vbYesNo)
                
                If Response = vbYes Then SkipBlanks = True
            End If
        
        'Import Entry
        ElseIf Cells(i, 8).Value = True Or _
            (ImportBlanks = True And IsEmpty(Cells(i, 8)) = True) Then
        
ImportItem:
            
            Debug.Print "Importing - " & Cells(i, 2).Value
            
            arrCal = Cells(i, 1).Value '****Accounts get their own calendar
        
            'Conditional to add new Folders if Folder does not exist.
            bool = True
            For j = 1 To CalFolder.Folders.Count
                If CalFolder.Folders.Item(j).Name = arrCal Then
                    bool = False
                    Exit For
                End If
            Next j
            
            If bool = True Then
                Set subFolder = CalFolder.Folders.Add(arrCal, olFolderCalendar)
            Else
                Set subFolder = CalFolder.Folders(arrCal)
            End If
            
            'Add New Appointment
            Set olAppt = subFolder.Items.Add(olAppointmentItem)
                
            'Define calendar item properties
            'For more info visit:
            'https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.appointmentitem_properties.aspx
            
            With olAppt
                .Start = Cells(i, 3)
                .Subject = Cells(i, 2).Value 'Header
                .Start = Cells(i, 3).Value
                
                If IsEmpty(Cells(i, 4)) = True Then
                    .AllDayEvent = True 'End/ All day event
                Else: .End = Cells(i, 4).Value
                End If
                
                If IsEmpty(Cells(i, 5)) = True Then
                    .ReminderSet = False 'Reminder
                Else: .ReminderMinutesBeforeStart = Cells(i, 5).Value
                End If
                
                .Location = Cells(i, 6).Value
                .Body = Cells(i, 7).Value
                
                'to mark automated emails (highly recommended)
                .Categories = "Orange Category"
                .Save
            End With
        
            Cells(i, 8) = False 'Save current state
        End If
        
        
        i = i + 1
    Loop
        
    Set olAppt = Nothing
    Set olApp = Nothing
    ThisWorkbook.Save
        
    MsgBox "Outlook has been updated!", , "Export to Outlook"
    
End Sub

'OPTIONAL:
'Use this function to quickly and effectively delete all appointment items within a specified folder
Private Sub ClearAppointments(CalFolder As Outlook.MAPIFolder)
        
        Dim Appt As Object
        
Redirect:
        
        For Each Appt In CalFolder.Items
            If Appt.Class = olAppointment Then
                'Debug.Print Appt.Subject & " - Deleted..."
                Appt.Delete
            End If
        Next Appt
        
        If Not CalFolder.Items.Count = 0 Then GoTo Redirect
    
End Sub
   
