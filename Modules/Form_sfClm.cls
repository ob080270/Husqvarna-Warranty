VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfClm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================================
' Module Name  : sfClm
' Description  : Form module for managing Claims (CLM) operations within the main form (frAP).
' Parent Form  : Form_frAP (main form)
'
' Key Features :
' - Handles claim data management, including defect descriptions and technician notes.
' - Supports adding and updating defect descriptions in the lookup table (tblDf_LU).
' - Automatically updates linked fields in the parent form (frAP) when a claim is selected.
'
' Events :
' 1. clTechn_AfterUpdate       - Handles updates to the technician notes field (clTechn).
'                                - Converts the input to uppercase.
'                                - Checks if the defect description exists in the lookup table (tblDf_LU).
'                                - Prompts the user to save or update the defect description if necessary.
' 2. Form_AfterInsert          - Updates the form's Tag property with the new claim number after a new claim is inserted.
' 3. Form_Current              - Updates linked fields in the parent form (frAP) when the current claim changes.
'
' Developer    : Oleh Bondarenko
' Created      : 2011-10-14
' Last Updated : 2025-02-22 by Oleh Bondarenko - Added comments for GitHub upload
' ==============================================================================================
Option Explicit
Option Compare Database

' ---------------------------------------------------------------------------
' Event #1        : clTechn_AfterUpdate
' Purpose         : Handles updates to the technician notes field (clTechn).
' Behavior        :
'                  - Converts the input to uppercase to ensure consistency.
'                  - Checks if the defect description (clTechn) exists in the lookup table (tblDf_LU).
'                  - If the defect description does not exist, prompts the user to save it.
'                  - If the defect description exists, prompts the user to update it.
'                  - Updates the lookup table (tblDf_LU) with the new or modified description.
' ---------------------------------------------------------------------------
Private Sub clTechn_AfterUpdate()
    Dim strCriteria As String       ' - Criteria for searching the defect description in the lookup table
    Dim strOldDescr As Variant      ' - Existing defect description
    Dim intAddRec As Integer        ' - User's response to the prompt
    Dim rst As Recordset            ' - Recordset for updating the lookup table
        
    Me!clTechn = UCase(Me!clTechn)                                          ' - Convert the technician notes to uppercase for consistency
    strCriteria = "[dfCode] ='" & Me!clDf & "'"                             ' - Build the search criteria for the defect code
    strOldDescr = DFirst("[dfCode]", "tblDf_LU", strCriteria)               ' - Retrieve the existing defect description from the lookup table
    
    If IsNull(strOldDescr) Then                                             ' - If the defect description does not exist in the lookup table
        intAddRec = MsgBox("Save new defect description?" _
                            , vbQuestion + vbYesNo, "Saving")               '   - Prompt the user to save the new defect description
        
        If intAddRec = 6 Then                                               '   - If the user chooses "Yes" (vbYes = 6)
            Set rst = CurrentDb.OpenRecordset("tblDf_LU", , dbAppendOnly)   '       - Open the lookup table for appending a new record
            
            With rst
                .AddNew                 ' - Add a new record to the lookup table
                !dfCode = Me!clDf       ' - Save the defect code
                !dfDescr = Me!clTechn   ' - Save the defect description
                .Update                 ' - Commit the changes to the table
            End With
            rst.Close                   ' - Close the recordset
            Set rst = Nothing           ' - Release the recordset object
            
        End If
    Else                                                                    ' - If the defect description already exists
        If IsNull(Me!clTechn) Then                                          '   - If the new description is empty
            Exit Sub                                                        '       - don't ask prompt to change description in the tblDf_LU
        End If
        
        intAddRec = MsgBox("Change defect description?" _
                            , vbQuestion + vbYesNo, "Changing")             ' - Prompt the user to update the existing defect description
        
        If intAddRec = 6 Then                                               ' - If the user chooses "Yes" (vbYes = 6)
            Me!luDescr = Me!clTechn                                         ' - Update the description in the form
        End If
    
    End If

End Sub

' --------------------------------------------------------------------------------------------------
' Event #2        : Form_AfterInsert
' Purpose         : Handles actions after a new claim record is inserted.
' Behavior        :
'                  - Appends the new claim number to the existing list of claims in the Tag property.
'                  - Ensures that the form retains a record of all claims added during the session.
' --------------------------------------------------------------------------------------------------
Private Sub Form_AfterInsert()
    
    Me.Tag = Me.Tag & Me!clNr & "; "        ' - Append the new claim number to the form's Tag property
    
End Sub

' ---------------------------------------------------------------------------
' Event #3        : Form_Current
' Purpose         : Handles actions when the current record in the form changes.
' Behavior        :
'                  - Updates linked fields in the parent form (frAP) when the current claim changes.
'                  - Synchronizes the claim number (clNr) with the parent form's linked field (txtClmLnk).
'                  - Ensures that the parent form reflects the currently selected claim.
' ---------------------------------------------------------------------------
Private Sub Form_Current()
On Error Resume Next                    ' - Ignore errors to prevent interruptions in case of missing or invalid data
    Me.Parent("txtClmLnk") = [clNr]     ' - Update the linked field in the parent form (frAP) with the current claim number (clNr)
    'Me.Parent("txtLink2") = [dfNo]
End Sub
