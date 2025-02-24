VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfVhc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
' Module Name  : sfVhc
' Description  : Subform module for managing vehicle (VHC) data.
' Parent Form  : frCust
'
' Key Features :
' - Handles vehicle data management, including VIN and warranty information.
' - Triggers email notifications for warranty activation when a new vehicle is added.
'
' Events :
' 1. Form_AfterInsert - Handles actions after a new vehicle record is inserted.
'                       - Checks if the VIN (vhVIN) and warranty start date (vhWS) fields are filled.
'                       - If both fields are filled, triggers an email notification to the warranty administrator
'                         with a request to activate the motorcycle in the ARCO system.
'
' Developer    : Oleh Bondarenko
' Created      : 2011-10-18
' Last Updated : 2025-02-22 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Explicit
Option Compare Database

' ---------------------------------------------------------------------------
' Event #1        : Form_AfterInsert
' Purpose         : Handles actions after a new vehicle record is inserted.
' Behavior        :
'                  - Checks if the VIN (vhVIN) and warranty start date (vhWS) fields are filled.
'                  - If both fields are filled, triggers an email notification to the warranty administrator
'                    with a request to activate the motorcycle in the ARCO system.
' ---------------------------------------------------------------------------
Private Sub Form_AfterInsert()
    
    If Not IsNull(Me!vhVIN) And Not IsNull(Me!vhWS) Then    ' - If both fields aren't empty:
        Call SndMail(1)                                     '   - Send an email to the warranty administrator
    End If                                                  '     about activation the moto in the system ARCO.

End Sub
