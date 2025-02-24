VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frWA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
' Module Name  : frWA
' Description  : Form module for managing Work Authorization (WA) operations.
' Child Form   :
'               - sfWA_Itm: Work and Parts
'
' Key Features :
'               - Handles Work Authorization (WA) data management.
'               - Supports searching and updating WA records.
'               - Implements data validation and user prompts for WA transactions.
'
' Events :
' 1. cbClmFnd_AfterUpdate    - Searches for a work authorization record corresponding to the specified claim number.
'                              - Retrieves the claim number from the dataset (query "qsCbClmFnd2").
'                              - Locates the corresponding record in the form's dataset.
'                              - Moves the form to the found record.
' 2. cmdExitApp_Click        - Handles the application exit process.
'                              - Prompts the user to confirm if the claim should be marked as entered in Arco.
'                              - If confirmed, updates the user and date fields, and sends an email notification.
'                              - Closes the application.
' 3. swEdit_AfterUpdate      - Toggles edit mode on and off based on the swEdit checkbox.
'                              - Enables or disables editing for key fields in the main form.
'                              - Controls the ability to add and edit records in the subform.
'                              - Updates UI elements to visually indicate the current mode.
'
' Developer    : Oleh Bondarenko
' Created      : 2012-01-25
' Last Updated : 2025-02-21 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Compare Database
Option Explicit

' ---------------------------------------------------------------------------
' Event #1        : cbClmFnd_AfterUpdate
' Purpose         : Invoice number lookup
' Behavior        :
'                  - Retrieves the claim number from the dataset using the query "qsCbClmFnd2".
'                  - Locates the corresponding record in the form's dataset.
'                  - Moves the form to the found record, if it exists.
' ---------------------------------------------------------------------------
Private Sub cbClmFnd_AfterUpdate()
' Найти акт выполненных работ, соответствующий указанной рекламации

    Dim fndAct As Long      ' - Invoice number
    Dim rst As Recordset    ' - Recordset for navigating form data
    
    fndAct = Nz(DLookup("[apAct]", "qsCbClmFnd2"))
    
    Set rst = Me.RecordsetClone         ' - Get a clone of the form's recordset for safe navigation
    rst.FindLast "[apAct] = " & fndAct  ' - Locate the record with the corresponding claim numbe
    Me.Bookmark = rst.Bookmark          ' - Move the form to the located record

'   Clean up the recordset object to free resources:
    rst.Close
    Set rst = Nothing
    
End Sub

' ---------------------------------------------------------------------------
' Event #2        : cmdExitApp_Click
' Purpose         : Handles the application exit process.
' Behavior        :
'                  - Prompts the user to confirm if the current claim should be marked as entered in Arco.
'                  - If the user confirms, updates the user and date fields (clUsr, clArcoDt) with the current user and timestamp.
'                  - Sends an email notification using the SndMail function (module: GlobalFn).
'                  - Closes the application gracefully.
' External Calls  :
'                  - SndMail() - Sends an email notification if the claim is marked as entered in Arco (module: GlobalFn).
' ---------------------------------------------------------------------------
Private Sub cmdExitApp_Click()
' Спросить пользователя считать ли данную рекламацию введенной в Арко
' Если да - сохранить юзера и дату, создать инфо e-mail
' Если нет - просто закрыть программу
    
    Dim msg As String                               ' - Text of the message box prompt
    Dim Style As Integer                            ' - Style of the message box (e.g., buttons and icon)
    Dim Title As String                             ' - Title of the message box
    Dim Responce As Integer                         ' - Variable to store the user's response
    
'   Set up the message box prompt:
    msg = "Should claim #" & Me!clNr & " be marked as entered in Arco?"
    Title = "Close the Program"
    Style = vbYesNo + vbQuestion                    ' - Yes/No buttons with a question icon
    Responce = MsgBox(msg, Style, Title)            ' - Display the message box and capture the user's response
    
    If Responce = vbYes Then                        ' - If the user confirms (clicks "Yes")
        Me!clUsr = CurrentUser()                    ' - Update the user field with the current user
        Me!clArcoDt = Now()                         ' - Update the date field with the timestamp
        Call SndMail(4)                             ' - Send an email notification ("4"-the Claim has been registered in ARCO)
    End If
    
    DoCmd.Quit                                      ' - Close the application

End Sub

' ---------------------------------------------------------------------------
' Event #3        : swEdit_AfterUpdate
' Purpose         : Toggles edit mode on and off based on the swEdit checkbox.
' Behavior        :
'                  - Enables or disables editing for key fields in the main form.
'                  - Controls the ability to add and edit records in the subform.
'                  - Updates UI elements to visually indicate the current mode.
'                  - When edit mode is enabled, unlocks fields and allows additions in the subform.
'                  - When edit mode is disabled, locks fields and prevents additions in the subform.
' ---------------------------------------------------------------------------
Private Sub swEdit_AfterUpdate()
' Включение/выключение режима правки
' Выключить блокировку полей, подкрасить поля фоном, выключить блокировку подформ
    Dim frIt As Form                    '- reference to the Items subform
    
    Set frIt = sfWA_Itm.Form
    
'   Check if edit mode is enabled (swEdit checkbox is checked):
    If Me!swEdit Then                   ' - If edit mode is enabled:
        Me!apTp.Locked = False          '   - Unlock the "apTp" field for editing
        Me!apTp.BackStyle = 1           '   - Set the background style to opaque (visible)
        Me!clDf.Locked = False          '   - Unlock the "clDf" field for editing
        Me!clDf.BackStyle = 1           '   - Set the background style to opaque (visible)
        
        sfWA_Itm.Locked = False         '   - Unlock the items subform
        frIt.AllowAdditions = True      '     and allow additions
        
        lbEdit.Visible = True           '   - Show the "Edit Mode" label to indicate that editing is enabled
        apTp.SetFocus                   '   - Set focus to the "apTp" field for user input
        
    Else                                ' - If edit mode is disabled:
        Me!apTp.Locked = True           '   - Lock the "apTp" field to prevent editing
        Me!apTp.BackStyle = 0           '   - Set the background style to transparent
        Me!clDf.Locked = True           '   - Lock the "clDf" field to prevent editing
        Me!clDf.BackStyle = 0           '   - Set the background style to transparent
        
        sfWA_Itm.Locked = True          '   - Lock the items subform and prevent additions
        frIt.AllowAdditions = False     '     and prevent additions
        
        lbEdit.Visible = False          '   - Hide the "Edit Mode" label to indicate that editing is disabled
        clNr.SetFocus                   '   - Set focus to the "clNr" field
        
        'Call CheckClmTag            ' Проверить дескриптор на предмет готовых клеймов
        
    End If
    
End Sub
