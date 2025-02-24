VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frDfSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ================================================================================================
' Module       : Form_frDfSelect
' Description  : Handles the selection of defect codes in the warranty claims system.
'
' Key Features :
'               - Manages defect code selection interface.
'               - Processes user interactions with form elements.
'
' Events:
'   1. cmdClose_Click           - Closes the form when the Close button is clicked.
'   2. cmdFindSmp_Click         - Triggers a search operation for a sample based on input criteria.
'   3. lstSbGr_AfterUpdate      - Handles updates when a sub-group selection is changed.
'   4. lstSmp_AfterUpdate       - Handles updates when a sample selection is changed.
'   5. objTreeView_Collapse     - Handles node collapse events in the TreeView control.
'   6. objTreeView_Expand       - Handles node expand events in the TreeView control.
'   7. objTreeView_NodeClick    - Processes user clicks on nodes in the TreeView control.
'   8. Form_Load               - Initializes the form and loads necessary data.
'   9. Form_Unload             - Handles cleanup operations when the form is closed.
'   10.swRuEn_AfterUpdate       - Toggles between Russian and English interface.
'   11. cmdClrFnd_Click          - Clears the search results.

' Procedures / Functions:
'   1. LoadTreeView             - Loads the TreeView control with defect categories and items.
'   2. GetImageIndex            - Retrieves the appropriate image index for a TreeView node.
'   3. LoadImages               - Loads images/icons used in the TreeView control.
'   4. OpenSbGrNdRst            - Opens a recordset for sub-group nodes.
'   5. OpenGroupNodeRst         - Opens a recordset for main group nodes.

' Developer    : Oleh Bondarenko
' Created      : 2011-11-07
' Last Updated : 2025-02-20 by Oleh Bondarenko - Added comments for GitHub upload
' ================================================================================================

Option Compare Database
Option Explicit
 
' ----------------------------------------------------------------------------------------------------
' Declarations: Declare module-level recordsets for handling defect selection and TreeView structure.
' ----------------------------------------------------------------------------------------------------
Private mrstDfSel As Recordset       ' Recordset for storing selected defect codes.
Private mstrGrNd As Recordset        ' Recordset for storing names of group nodes in TreeView.
Private mstrSbGrNd As Recordset      ' Recordset for storing names of sub-group nodes in TreeView.
Private mrstImages As Recordset      ' Recordset for storing image data for TreeView nodes.

' -------------------------------------------------------------------
' Procedure #1 : LoadTreeView
' Purpose      : Populates the TreeView control with defect categories.
' Behavior     :
'               - Initializes the TreeView control.
'               - Adds a root node for the hierarchy.
'               - Iterates through the group nodes and adds them as child nodes.
'               - Expands the root node after loading.
' External Calls  :
'               - GetImageIndex (determines the appropriate image index for nodes).
' -------------------------------------------------------------------
Private Sub LoadTreeView()
    'Dim objImageTypes As Control
    'Dim strCountry As String

    Dim objTreeView As Control              ' TreeView control instance
    Dim objNode As Node                     ' Node object for TreeView
    
    Set objTreeView = Me.objTreeView        ' Assign the TreeView control
                
    mstrGrNd.MoveFirst                      ' Move to the first record in the group node recordset
    'mstrSbGrNd.MoveFirst
    
'   Add the root node to the TreeView:
    Set objNode = objTreeView.Nodes.Add(, , "HQ", "Structural groups", 1)
    
'   Iterate through the group node recordset and add them as child nodes:
    Do Until mstrGrNd.EOF
        Set objNode = objTreeView.Nodes.Add("HQ", tvwChild, Mid(mstrGrNd!GrNd, 2, 1), mstrGrNd!GrNd, GetImageIndex(Mid(mstrGrNd!GrNd, 2, 1)))
        mstrGrNd.MoveNext
    Loop
    
'   Expand the root node to show all group nodes by default:
    objTreeView.Nodes.Item("HQ").Expanded = True
                                
End Sub

' -----------------------------------------------------------------------------------------------
' Function #2  : GetImageIndex
' Purpose      : Retrieves the image index for a given group ID.
' Behavior     :
'               - Searches for the corresponding image index in the image recordset (mrstImages).
'               - Moves to the first record and finds the matching entry.
'               - Returns the image index value.
' Parameter    :    - strArg:   The Group ID used to locate the corresponding image index.
' Return       :    - Byte  :   The image index value from the recordset.
' -----------------------------------------------------------------------------------------------
Private Function GetImageIndex(strArg As String) As Byte

'   Move to the first record in the image recordset:
    mrstImages.MoveFirst

'   Search for the record where GrID matches the provided argument:
    mrstImages.FindFirst "GrID = " & "'" & strArg & "'"

'   Return the corresponding image index from the recordset:
    GetImageIndex = mrstImages!ImageIndex

End Function

' -------------------------------------------------------------------
' Event #1     : cmdClose_Click
' Purpose      : Closes the defect selection form.
' Behavior     :
'               - Checks if a defect code has been selected.
'               - If no code is selected, prompts the user for confirmation before closing.
'               - If a code is selected, updates the main form with the selected values.
'               - Closes the form after processing.
' External Calls  :
'               - fnDfDescr (retrieves defect description based on code).
' -------------------------------------------------------------------
Private Sub cmdClose_Click()

    Dim msg As String           ' Message text for confirmation
    Dim Style As Integer        ' MessageBox style options
    Dim Title As String         ' Title of the MessageBox
    Dim Responce As Integer     ' User response to MessageBox
    
'   Check if no defect code has been selected:
    If clcRes = "" Then
        msg = "Defect code is not selected. Close the form without selection?"
        Title = "Confirmation required..."
        Style = vbYesNo + vbExclamation + vbDefaultButton2
    
        Responce = MsgBox(msg, Style, Title)            ' Display a confirmation message
    
        If Responce = vbYes Then                        ' User wants to close the form without selecting a defect
            DoCmd.Close acForm, "frDfSelect"
        End If
    Else
'       Update the claim form with the selected defect code and description:
        Forms!frAP!sfClm.Form!clDf = Me!clcRes
        Forms!frAP!sfClm.Form!clTechn = fnDfDescr(Left([clcRes], 4), Right([clcRes], 3))

'       Close the form after processing:
        DoCmd.Close acForm, "frDfSelect"
    End If

End Sub

' -------------------------------------------------------------------
' Event #2     : cmdFindSmp_Click
' Purpose      : Searches for symptoms based on user input.
' Behavior     :
'               - Checks the selected language (Russian or English).
'               - If no input is provided, loads all available symptoms.
'               - If input is provided, filters the symptoms by matching text.
'               - Updates the list box with the search results.
' -------------------------------------------------------------------
Private Sub cmdFindSmp_Click()
    
    Dim strSQL As String    ' SQL query string for retrieving symptoms
    
'   Check if Russian language is selected:
    If swRuEn = 1 Then
'       Determine the source of strings for the symptom list, depending on the content of the search string:
        If Me!txtSrch = "" Then
'           If the search field is empty, load all symptoms:
            strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmRu FROM tblSymptm ORDER BY [smpID];"
        Else
'           Search for symptoms that match the entered text
            strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmRu" & _
                    " FROM tblSymptm" & _
                    " WHERE (((tblSymptm.smpNmRu) Like ""*" & [Forms]![frDfSelect]![txtSrch] & "*""))" & _
                    " ORDER BY tblSymptm.smpID;"
        End If

'       Update the list box with the filtered results:
        Me!lstSmp.RowSource = strSQL
        Me!lstSmp.Requery
    Else
'   If English language is selected:
        If Me!txtSrch = "" Then
            strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmEn FROM tblSymptm ORDER BY [smpID];"
        Else
            strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmEn" & _
                    " FROM tblSymptm" & _
                    " WHERE (((tblSymptm.smpNmEn) Like ""*" & [Forms]![frDfSelect]![txtSrch] & "*""))" & _
                    " ORDER BY tblSymptm.smpID;"
        End If

'       Update the list box with the filtered results:
        Me!lstSmp.RowSource = strSQL
        Me!lstSmp.Requery
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #3     : lstSbGr_AfterUpdate
' Purpose      : Updates the tooltip text of the sub-group list.
' Behavior     :
'               - Retrieves the sub-group name based on the selected defect ID.
'               - Updates the tooltip text for better user guidance.
'               - Refreshes the list to apply changes.
' -------------------------------------------------------------------
Private Sub lstSbGr_AfterUpdate()

'   Update the tooltip text to display the selected sub-group name:
    lstSbGr.ControlTipText = DLookup("[sgrNmRu]", "qsSbGrLstRu", "[DfID] ='" & lstSbGr.Value & "'")

'   Refresh the list to apply the changes:
    lstSbGr.Requery
    
End Sub

' -------------------------------------------------------------------
' Event #4     : lstSmp_AfterUpdate
' Purpose      : Updates the tooltip text of the symptom list.
' Behavior     :
'               - Retrieves the symptom name based on the selected symptom ID.
'               - Updates the tooltip text to provide more context to the user.
'               - Refreshes the list to apply changes.
' -------------------------------------------------------------------
Private Sub lstSmp_AfterUpdate()

'   Update the tooltip text to display the selected symptom name:
    lstSmp.ControlTipText = DLookup("[smpNmRu]", "tblSymptm", "[smpID] ='" & lstSmp.Value & "'")
    lstSmp.Requery

End Sub

' -----------------------------------------------------------------------------------
' Event #5     : objTreeView_Collapse
' Purpose      : Handles the collapse event of a TreeView node.
' Behavior     :
'               - Calls objTreeView_NodeClick to process the collapse action.
'               - Ensures that the necessary updates occur when a node is collapsed.
' Parameter    :
'               - Node (The TreeView node that was collapsed.)
' External Calls  :
'               - objTreeView_NodeClick (handles the node click event, see Event #7).
' -----------------------------------------------------------------------------------
Private Sub objTreeView_Collapse(ByVal Node As Object)

    Call objTreeView_NodeClick(Node)    ' Call the node click event handler to process the collapse action

End Sub

' -------------------------------------------------------------------
' Event #6     : objTreeView_Expand
' Purpose      : Handles the expand event of a TreeView node.
' Behavior     :
'               - Calls objTreeView_NodeClick to process the expand action.
'               - Ensures that necessary updates occur when a node is expanded.
' Parameters   :
'               - Node (The TreeView node that was expanded)
' External Calls  :
'               - objTreeView_NodeClick (handles the node click event, see Event #7).
' -------------------------------------------------------------------
Private Sub objTreeView_Expand(ByVal Node As Object)

    Call objTreeView_NodeClick(Node)    ' Call the node click event handler to process the expand action
    
End Sub

' -------------------------------------------------------------------
' Event #7     : objTreeView_NodeClick
' Purpose      : Handles the selection of a node in the TreeView.
' Behavior     :
'               - Extracts the selected node's key and assigns it to a form field.
'               - Determines whether to display all sub-groups or filter by selected group.
'               - Updates the list box with sub-groups based on the selected TreeView node.
' Parameters   :
'               - Node (The TreeView node that was clicked)
' -------------------------------------------------------------------
Private Sub objTreeView_NodeClick(ByVal Node As Object)

    Dim frm As Form                         ' Reference to the current form
    Dim strSQL As String                    ' SQL query to populate the sub-group list
    
    'mrstDfSel.Close
    Set frm = Forms!frDfSelect              ' Set the form reference

'   Extract the key from the selected node and assign it to a text box:
    frm!txtTreeValue = Mid(Node, 2, 1)
    
'   Determine the behavior based on the selected node value:
    Select Case frm!txtTreeValue
        Case "î"
'           Check the selection of the Ru / En switch state:
            If swRuEn = 1 Then
'               Load all sub-groups without filtering (Ru):
                strSQL = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmRu" & _
                        " FROM tblSbGr;"
            Else
'               Load all sub-groups without filtering (En):
                strSQL = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmEn" & _
                        " FROM tblSbGr;"
            End If
            frm!lstSbGr.RowSource = strSQL
            frm!lstSbGr.Requery
        Case "A" To "X"
'           Filter sub-groups based on selected group (using language setting):
            If swRuEn = 1 Then                                                              'Ru:
                strSQL = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmRu" & _
                        " FROM tblSbGr" & _
                        " WHERE (((tblSbGr.sgrGrID)=[Forms]![frDfSelect]![txtTreeValue]));"
            Else                                                                            'En:
                strSQL = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmEn" & _
                        " FROM tblSbGr" & _
                        " WHERE (((tblSbGr.sgrGrID)=[Forms]![frDfSelect]![txtTreeValue]));"
            End If
            frm!lstSbGr.RowSource = strSQL
            frm!lstSbGr.Requery
    End Select
    
    frm!clcRes.Requery
    
End Sub

' -------------------------------------------------------------------
' Event #8    : Form_Load
' Purpose      : Initializes the form and loads required data.
' Behavior     :
'               - Opens the recordset containing group node names.
'               - Loads image data for TreeView nodes.
'               - Populates the TreeView control with defect groups.
' External Calls  :
'               - OpenGroupNodeRst (opens recordset for group nodes).
'               - LoadImages (loads image data for TreeView).
'               - LoadTreeView (populates the TreeView with data).
' -------------------------------------------------------------------
Private Sub Form_Load()
     
'   Open the recordset with group node names:
    Call OpenGroupNodeRst
       
    ' Load Image Data
    Call LoadImages
        
    ' Load the TreeView control.
    Call LoadTreeView
          
End Sub

' -------------------------------------------------------------------
' Procedure #3 : LoadImages
' Purpose      : Loads image data for the TreeView nodes.
' Behavior     :
'               - Opens a recordset containing image data.
'               - Retrieves image indexes for use in the TreeView.
' -------------------------------------------------------------------
Private Sub LoadImages()
    
    Dim strSQL As String        ' SQL query to retrieve image data
    
    ' Open Images Recordset
    strSQL = "SELECT * FROM  tblImages"
    Set mrstImages = CurrentDb.OpenRecordset(strSQL, , dbReadOnly)

End Sub

' -------------------------------------------------------------------
' Event #9    : Form_Unload
' Purpose      : Cleans up memory by releasing object variables when the form is closed.
' Behavior     :
'               - Destroys object variables related to defect selection, sub-groups, and images.
'               - Ensures proper memory management by setting recordsets to Nothing.
' Parameters   :
'               - Cancel     - Integer parameter that allows canceling the unload process (not used).
' -------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    
    ' Destroy the object variables.
    Set mrstDfSel = Nothing
    Set mstrSbGrNd = Nothing
    Set mrstImages = Nothing

End Sub

' -------------------------------------------------------------------
' Procedure #4 : OpenSbGrNdRst
' Purpose      : Opens a recordset containing sub-group node names.
' Behavior     :
'               - Retrieves data from the query "qrySbGrNd".
'               - Opens the recordset in read-only mode for use in the form.
' -------------------------------------------------------------------
Private Sub OpenSbGrNdRst()

    ' Create a Recordset object with names of nodes of construction subgroups
    Set mstrSbGrNd = CurrentDb.OpenRecordset("qrySbGrNd", , dbReadOnly)
                     
End Sub

' -------------------------------------------------------------------
' Procedure #5 : OpenGroupNodeRst
' Purpose      : Opens a recordset containing group node names.
' Behavior     :
'               - Retrieves data from the query "qryGrNd".
'               - Opens the recordset in read-only mode for use in the form.
' -------------------------------------------------------------------
Private Sub OpenGroupNodeRst()

    ' Create a Recordset object with names of nodes of construction groups
    Set mstrGrNd = CurrentDb.OpenRecordset("qryGrNd", , dbReadOnly)
                     
End Sub

' -------------------------------------------------------------------
' Event #10    : swRuEn_AfterUpdate
' Purpose      : Handles language switching between Russian and English.
' Behavior     :
'               - Updates the list of sub-groups based on the selected language.
'               - Clears the search field and sets the keyboard language accordingly.
'               - Updates the list of symptoms based on the selected language.
' -------------------------------------------------------------------
Private Sub swRuEn_AfterUpdate()
'   After switching the language - change the source of the lists

    Dim strSQL1 As String             ' SQL query for sub-group list
    Dim strSQL2 As String             ' SQL query for symptom list
    
    If swRuEn = 1 Then
'   If Russian language is selected:
'       Load all sub-groups or filter them based on the selected group:
        If Forms!frDfSelect!txtTreeValue = "î" Then                                         ' - Root node "Whole motorcycle"
            strSQL1 = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmRu FROM tblSbGr;"
        Else                                                                                ' - Group EXCEPT "Whole motorcycle" selected"
            strSQL1 = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmRu" & _
                    " FROM tblSbGr" & _
                    " WHERE (((tblSbGr.sgrGrID)=[Forms]![frDfSelect]![txtTreeValue]));"
        End If
        Me!lstSbGr.RowSource = strSQL1
        
'       Clear the search field and set keyboard input to Russian:
        Me!txtSrch = ""
        Me!txtSrch.KeyboardLanguage = 27
        
'       Load the symptom list in Russian:
        strSQL2 = "SELECT tblSymptm.smpID, tblSymptm.smpNmRu FROM tblSymptm ORDER BY [smpID];"
        Me!lstSmp.RowSource = strSQL2

'       Refresh the lists:
        Me!lstSbGr.Requery
        Me!lstSmp.Requery
    Else
'   If English language is selected:
        ' All the same for the English version of the lists:
        If Forms!frDfSelect!txtTreeValue = "î" Then                                         ' - Root node "Whole motorcycle"
            strSQL1 = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmEn FROM tblSbGr;"
        Else                                                                                ' - Group EXCEPT "Whole motorcycle" selected"
            strSQL1 = "SELECT [sgrGrID] & [sgrID] AS DfID, tblSbGr.sgrNmEn" & _
                    " FROM tblSbGr" & _
                    " WHERE (((tblSbGr.sgrGrID)=[Forms]![frDfSelect]![txtTreeValue]));"
        End If
        Me!lstSbGr.RowSource = strSQL1
        
'       Clear the search field and set keyboard input to English:
        Me!txtSrch = ""
        Me!txtSrch.KeyboardLanguage = 11
        
'       Load the symptom list in English:
        strSQL2 = "SELECT tblSymptm.smpID, tblSymptm.smpNmEn FROM tblSymptm ORDER BY [smpID];"
        Me!lstSmp.RowSource = strSQL2

'       Refresh the lists:
        Me!lstSbGr.Requery
        Me!lstSmp.Requery
    End If
    
End Sub

' -------------------------------------------------------------------
' Event #11 : cmdClrFnd_Click
' Purpose   : Clears the search field and resets the symptom list.
' Behavior  :
'               - Clears the text in the search field.
'               - Reloads the full list of symptoms based on the selected language.
'               - Updates the symptom list control.
' -------------------------------------------------------------------
Private Sub cmdClrFnd_Click()
    
    Dim strSQL As String    ' - SQL query to retrieve the full symptom list
    
    Me!txtSrch = ""         ' - Clear the search field
    If swRuEn = 1 Then
'       Load the full list of symptoms based on the selected language (Ru):
        strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmRu FROM tblSymptm ORDER BY [smpID];"
    Else
'       Load the full list of symptoms based on the selected language (En):
        strSQL = "SELECT tblSymptm.smpID, tblSymptm.smpNmEn FROM tblSymptm ORDER BY [smpID];"
    End If
    
'   Update the symptom list control:
    Me!lstSmp.RowSource = strSQL
    Me!lstSmp.Requery
    
End Sub
