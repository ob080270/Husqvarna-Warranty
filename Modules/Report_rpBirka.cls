VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpBirka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
' Module Name  : rpBirka
' Description  : Report module for generating and managing the Birka report.
'
' Key Features :
' - Handles the closing event of the report.
' - Ensures data consistency by refreshing the related subform (sfItm) in the main form (frAP).
'
' Events :
' 1. Report_Close - Handles actions when the report is closed.
'                   - Refreshes the sfItm subform in the frAP form to ensure data consistency.
'
' Developer    : Oleh Bondarenko
' Created      : 2012-01-17
' Last Updated : 2025-02-22 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Explicit
Option Compare Database

' ---------------------------------------------------------------------------
' Event #1        : Report_Close
' Purpose         : Handles actions when the report is closed.
' Behavior        :
'                  - Refreshes the sfItm subform in the frAP form to ensure data consistency.
'                  - Ensures that any changes made during the report session are reflected in the main form.
' ---------------------------------------------------------------------------
Private Sub Report_Close()

    Forms!frAP!sfItm.Form.Requery   ' Refresh the sfItm subform in the frAP form to ensure data consistency

End Sub
