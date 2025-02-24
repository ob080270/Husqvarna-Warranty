VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_sfItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ===========================================================================
' Module Name  : sfItm
' Description  : Subform module within the main form frAP for managing items (parts and labor).
' Parent Form  : frAP (main form)
'
' Key Features :
' - Handles item data management, including parts and labor.
' - Supports data validation and automation for item entries.
' - Manages item descriptions, quantities, and pricing.
' - Generates reports for item tags (birka).
'
' Events :
' 1. clTechn_AfterUpdate       - Updates the technical description of an item.
'                                - Saves or updates the description in the lookup table (tblIt_LU).
'                                - Prompts the user to save or update the description if necessary.
' 2. itBirka_AfterUpdate       - Generates a report for the item tag (birka) if the item is a part.
' 3. itBirka_BeforeUpdate      - Validates if the item is a part before allowing the generation of a tag.
' 4. itCI_AfterUpdate          - Sets the default value for the credit-invoice field.<<<<<<---------------------????????????????
' 5. itNr_AfterUpdate          - Updates item details based on the item number.
'                                - Automatically sets the item type (part or labor) based on the item number.
'                                - Retrieves item description and price from the lookup table.
' 6. itQ_AfterUpdate           - Updates the charge for labor items based on the quantity of hours.
' 7. itRmb_AfterUpdate         - Converts the reimbursement price to UAH based on the exchange rate.
'
' Developer    : Oleh Bondarenko
' Created      : 2011-10-17
' Last Updated : 2025-02-22 by Oleh Bondarenko - Added comments for GitHub upload
' ===========================================================================
Option Explicit
Option Compare Database

' ---------------------------------------------------------------------------
' Event #1        : clTechn_AfterUpdate
' Purpose         : Handles the update of the technical description (clTechn) field.
' Behavior        :
'                  - Converts the entered description to uppercase.
'                  - Checks if the item number (itNr) exists in the lookup table (tblIt_LU).
'                  - If the item number does not exist in the lookup table, prompts the user to save the new description.
'                  - If the user confirms, adds the new description to the lookup table.
'                  - If the item number exists, prompts the user to update the existing description.
'                  - Updates the description in the lookup table if the user confirms.
' ---------------------------------------------------------------------------
Private Sub clTechn_AfterUpdate()
    Dim strCriteria As String       ' - Criteria for searching the item number in the lookup table
    Dim strOldDescr As Variant      ' - Existing description from the lookup table
    Dim intAddRec As Integer        ' - User's response to the confirmation dialog
    Dim rst As Recordset            ' - Recordset for updating the lookup table
        
    Me!clTechn = UCase(Me!clTechn)                                              ' - Convert the entered description to uppercase
    strCriteria = "[luItNr] ='" & Me!itNr & "'"                                 ' - Define the search criteria for the item number in the lookup table
    strOldDescr = DFirst("[luItNr]", "tblIt_LU", strCriteria)                   ' - Retrieve the existing description for the item number from the lookup table
    
    If IsNull(strOldDescr) Then                                                 ' - If the item number does not exist in the lookup table
        intAddRec = MsgBox("Save the description?" _
                            , vbQuestion + vbYesNo, "Save")                     '   - prompt the user to save the new description
        If intAddRec = 6 Then                                                   '   - If the user chooses "Yes" (vbYes = 6):
            Set rst = CurrentDb.OpenRecordset("tblIt_LU", , dbAppendOnly)       '       - Open the lookup table for appending a new record
            
            With rst
                .AddNew                     ' - add the new description to the lookup table
                !luItNr = Me!itNr           ' - Save the item number
                !luItDesc = Me!clTechn      ' - Save the new item description
                .Update                     ' - Commit the changes to the table
            End With
            
            rst.Close                       ' - Close the recordset
            Set rst = Nothing               ' - Release the recordset object
            itQ.SetFocus                    ' - Move focus to the quantity field
            
        End If
    Else                                    ' - If the defect description already exists:
        If IsNull(Me!clTechn) Then          '   - Check if the new description is empty
            Exit Sub                        '       - don't ask prompt to change description in the tblIt_LU
        End If
        
        intAddRec = MsgBox("Update description?" _
                            , vbQuestion + vbYesNo, "Update")                   ' Prompt the user to update the existing description
                            
        If intAddRec = 6 Then               ' - If the user chooses "Yes" (vbYes = 6):
            Me!luDescr = Me!clTechn         '   - Update the description in the form
        End If
    
    End If
    
End Sub

' ---------------------------------------------------------------------------
' Event #2        : itBirka_AfterUpdate
' Purpose         : Handles the update of the itBirka field.
' Behavior        :
'                  - Checks if the itBirka is marked as True.
'                  - If True, generates and opens a report (rpBirka) in preview mode for the item tag.
'                  - The report is typically used for printing tags for parts.
' ---------------------------------------------------------------------------
Private Sub itBirka_AfterUpdate()
' —генерировать отчет по бирке
    
    If itBirka Then
        DoCmd.OpenReport "rpBirka", acPreview
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #3        : itBirka_BeforeUpdate
' Purpose         : Validates the item type before allowing the update of the item label (birka) field.
' Behavior        :
'                  - Checks if the item type (itType) is not a part ("P").
'                  - If the item is not a part (e.g., it is labor or another type), cancels the update.
'                  - Prevents the generation of an item label for non-part items by undoing the change and displaying a warning message.
'                  - Ensures that item labels are only generated for parts, as labor items do not require them.
'                  - An item label (birka) is a small form or report attached to a part stored in the warranty warehouse.
'                   It contains essential information about the part, such as VIN, brand, model, mileage, warranty start date, replacement date,
'                   customer complaint, technician's conclusion, etc.
' Parameters      :
'                  - Cancel (Integer): If set to True (non-zero), cancels the update of the birka field.
'                   This is used to prevent the generation of item labels for non-part items.
' ---------------------------------------------------------------------------
Private Sub itBirka_BeforeUpdate(Cancel As Integer)

    If itType <> "P" Then                                   ' Check if the item type is not a part ("P")
        Cancel = True                                       ' Cancel the update
        Me.Undo                                             ' Undo the change to the birka field
        MsgBox "Item labels can only be printed for parts.", _
            vbExclamation, "Invalid Item Type"
    End If
    
End Sub

' ---------------------------------------------------------------------------
' Event #4        : itCI_AfterUpdate
' Purpose         : Handles the update of the credit-invoice (itCI) field.
' Behavior        :
'                  - Sets the default value of the credit-invoice field (itCI) to the newly entered value.
'                  - This ensures that the same credit-invoice value is pre-filled for subsequent records, improving data entry efficiency.
' ---------------------------------------------------------------------------
Private Sub itCI_AfterUpdate()

    itCI.DefaultValue = "'" & itCI & "'"    ' - Set the default value of the credit-invoice field to the newly entered value
    
End Sub

' ---------------------------------------------------------------------------
' Event #5        : itNr_AfterUpdate
' Purpose         : Handles the update of the item number (itNr) field.
' Behavior        :
'                  - Removes any spaces from the entered item number for consistency.
'                  - Retrieves the item description (clTechn) and price (itChrg) from the lookup table (tblIt_LU) based on the item number.
'                  - If the item number is "0", it is automatically replaced with "Ore" to indicate labor (work) type.
'                  - If the item number is "Ore", the item type (itType) is set to "W" (work), and the description is set to "Time".
'                  - For all other item numbers, the item type is set to "P" (part).
' ---------------------------------------------------------------------------
Private Sub itNr_AfterUpdate()
    
    itNr = Replace(itNr, " ", "")   ' - Remove any spaces from the item number
    clTechn = luDescr               ' - Retrieve the item description and price from the lookup table
    itChrg = luPrcAP                '   and price from the lookup table
    
    If itNr = "0" Then              ' - If the item number is "0",
        itNr = "Ore"                '   replace it with "Ore" to indicate labor type
    End If
    
    If itNr = "Ore" Then            ' - If the item number is "Ore"
        itType = "W"                '   - set the item type to "W" (work)
        clTechn = "Time"            '   - and description to "Time"
        clTechn.SetFocus            '   - Move focus to the description field
    Else                            ' - For all other cases
        itType = "P"                '  - set item type to part
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #6        : itQ_AfterUpdate
' Purpose         : Handles the update of the quantity (itQ) field.
' Behavior        :
'                  - If the item type (itType) is "W" (work), calculates the charge (itChrg) based on the quantity of hours entered.
'                  - The charge is determined by multiplying the quantity by the fixed rate for labor (apFRU) from the parent form (frAP).
'                  - This ensures that labor items are correctly priced based on the number of hours worked.
' ---------------------------------------------------------------------------
Private Sub itQ_AfterUpdate()
    
    If itType = "W" Then        ' - If the item type is work ("W")
        itChrg = Parent!apFRU   ' - Calculate charge using the fixed labor rate from the parent form
    End If

End Sub

' ---------------------------------------------------------------------------
' Event #7        : itRmb_AfterUpdate
' Purpose         : Handles the update of the reimbursement price (itRmb) field.
' Behavior        :
'                  - Converts the reimbursement price from the original currency to UAH (Ukrainian Hryvnia)
'                   using the exchange rate (aplExRt) from the parent form (frAP).
'                  - The converted amount is stored in the itRmbUAH field.
'                  - This ensures that reimbursement prices are correctly calculated in the local currency
'                   for reporting and accounting purposes.
' ---------------------------------------------------------------------------
Private Sub itRmb_AfterUpdate()
    
    If Not IsNull(itRmb) Then                       ' - If the reimbursement price (itRmb) is not null
        itRmbUAH = itRmb * Me.Parent("aplExRt")     ' - Convert to UAH using the exchange rate from the parent form
    End If

End Sub
