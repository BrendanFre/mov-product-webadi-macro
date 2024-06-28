Private Sub butDeliveryRecon_Click()
Call sub_delivery_recon
End Sub

Private Sub buttonRMA_Click()
Call rmaCheck
End Sub

Private Sub clean_Click()
    Selection.textToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, OtherChar:= _
        "", FieldInfo:=Array(1, 1)
End Sub






Private Sub CommandButton10_Click()
' Spell_Check Macro
' Keyboard Shortcut: Ctrl+o

Dim X As Range
For Each X In Selection
If Not Application.CheckSpelling(word:=X.Text) Then
X.Interior.Color = vbRed
End If
Next X
End Sub

Private Sub CommandButton11_Click()

' Keyboard Shortcut: Ctrl+j
' Clear Colour
 
 With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Private Sub CommandButton12_Click()
' Lower Macro
' Keyboard Shortcut: Ctrl+l

 For Each X In Selection
      X.value = LCase(X.value)
   Next
End Sub

Private Sub CommandButton13_Click()
' Upper Macro
' Keyboard Shortcut: Ctrl+u

 For Each X In Selection
      X.value = UCase(X.value)
   Next
End Sub

Private Sub CommandButton14_Click()
'This Checks to make sure you are about to use the correct Macro
PIPPEPRR_Check.Show

'Shows AddWordToEnd Form
'AddWordToEnd.Show

End Sub



Private Sub CommandButton16_Click()
 For Each X In Selection
      X.value = WorksheetFunction.Proper(X.value)
   Next
End Sub

Private Sub CommandButton17_Click()
Dim MDT_Report As String
Dim Hold_Report As String

answer = MsgBox("BT Postcode Check - Eisai Only", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
 Exit Sub
Else
End If

MDT_Report = ActiveWorkbook.Name
Sheets("Customer Sites Information").Select

Range("A1").Select
If ActiveSheet.AutoFilterMode Then
Else
Selection.AutoFilter
End If


ActiveSheet.Range("$A$1:$BD$51").AutoFilter Field:=8, Criteria1:="Active"
ActiveSheet.Range("$A$1:$BD$51").AutoFilter Field:=17, Criteria1:="BT*"

LR = Range("A" & Rows.count).End(xlUp).row
If LR < 2 Then

MsgBox "No BT Postcodes"
Exit Sub

End If

Range("A1:$A$" & LR).Select
count = Selection.SpecialCells(xlCellTypeVisible).count

Workbooks.Open fileName:= _
        "T:\Master_Data\A. Oracle\WebADI templates\Templates\WebADI Bulk Upload Holds Template.csv"
        
Rows("2:1000").Clear
Range("A2:$A$" & count) = "Master Data Request"
Range("B2:$B$" & count) = "Movianto UK"
Range("C2:$C$" & count) = "Item Category"
Range("D2:$D$" & count) = "Eisai."
Range("E2:$E$" & count) = "Ship to Site"
Range("G2:$G$" & count) = "Eisai no longer uses Movianto to deliver to NI"
Range("H2:$H$" & count) = "N"
Range("I2:$I$" & count) = "Y"


Hold_Report = ActiveWorkbook.Name
Windows(MDT_Report).Activate
Range("C2:$C$" & LR).Copy
Windows(Hold_Report).Activate
Range("F2").PasteSpecial Paste:=xlPasteValues
Range("A1").Select
Windows(MDT_Report).Activate
Selection.AutoFilter
Range("A1").Select
Windows(Hold_Report).Activate


MsgBox "Ready for Upload"


End Sub

Private Sub CommandButton18_Click()

Dim MDT_Report As String, Hold_Report As String, MDTSheet As Worksheet

answer = MsgBox("BT Postcode Check", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
 Exit Sub
Else
End If

MDT_Report = ActiveWorkbook.Name
Sheets("Customer Sites Information").Select
Set MDTSheet = ActiveWorkbook.Sheets("Customer Sites Information")

Range("A1").Select
If ActiveSheet.AutoFilterMode Then
Else
Selection.AutoFilter
End If


ActiveSheet.Range("$A$1:$BD$51").AutoFilter Field:=8, Criteria1:="Active"
ActiveSheet.Range("$A$1:$BD$51").AutoFilter Field:=17, Criteria1:="BT*"

LR = Range("A" & Rows.count).End(xlUp).row
If LR < 2 Then

MsgBox "No BT Postcodes"
MDTSheet.Cells.AutoFilter
Exit Sub

End If

Range("A1:$A$" & LR).Select
count = Selection.SpecialCells(xlCellTypeVisible).count
Count2 = ((count - 1) * 2) + 1

Workbooks.Open fileName:= _
        "T:\Master_Data\A. Oracle\WebADI templates\Templates\WebADI Bulk Upload Holds Template.csv"
        
Rows("2:1000").Clear
Range("A2:$A$" & Count2) = "Master Data Request"
Range("B2:$B$" & Count2) = "Movianto UK"
Range("C2:$C$" & Count2) = "Item Category"
Range("E2:$E$" & Count2) = "Ship to Site"
Range("G2:$G$" & Count2) = "Orders for NI need to be confirmed with the Client/Kam"
Range("H2:$H$" & Count2) = "N"
Range("I2:$I$" & Count2) = "Y"

ClientCat = count
Range("D2:$D$" & ClientCat) = "Eisai."
ClientCat = ClientCat + (count - 1)
ClientCate2 = count + 1
Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Baxter."

'ClientCat = ClientCat + (Count - 1)
'ClientCate2 = ClientCate2 + (Count - 1)
'Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Essential Gen."

'ClientCat = ClientCat + (Count - 1)
'ClientCate2 = ClientCate2 + (Count - 1)
'Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Essential Pharmaceuticals."

'ClientCat = ClientCat + (Count - 1)
'ClientCate2 = ClientCate2 + (Count - 1)
'Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Essential Pharma - Hosp Only."

'ClientCat = ClientCat + (Count - 1)
'ClientCate2 = ClientCate2 + (Count - 1)
'Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Chemidex."

'ClientCat = ClientCat + (Count - 1)
'ClientCate2 = ClientCate2 + (Count - 1)
'Range("$D$" & ClientCate2 & ":$D$" & ClientCat) = "Techdow."


Hold_Report = ActiveWorkbook.Name
Windows(MDT_Report).Activate
Range("C2:$C$" & LR).Copy
Windows(Hold_Report).Activate
Range("F2").PasteSpecial Paste:=xlPasteValues

LRCheck = Range("A" & Rows.count).End(xlUp).row
CountClient = count + 1
Range("$F$" & CountClient).PasteSpecial Paste:=xlPasteValues
CountClient = CountClient + count - 1

Do Until CountClient > LRCheck

Range("$F$" & CountClient).PasteSpecial Paste:=xlPasteValues
CountClient = CountClient + count - 1
Loop

Range("A1").Select
Windows(MDT_Report).Activate
Selection.AutoFilter
Range("A1").Select
Windows(Hold_Report).Activate

MsgBox "Ready for Upload"

MDTSheet.AutoFilter.ShowAllData

End Sub


Function CleanString(StrIn As String) As String

' characters, including carriage returns BUT NOT linefeeds.
' Does not remove special characters like symbols, international
' characters, etc. This function runs recursively, each call
' removing one embedded character

   Dim iCh  As Integer
   Dim Ch   As Integer      'a single character to be tested
   CleanString = StrIn
   For iCh = 1 To Len(StrIn)
      Ch = Asc(Mid(StrIn, iCh, 1))
      If Ch < 32 And Ch <> 10 Then
         'remove special character
         CleanString = Left(StrIn, iCh - 1) & CleanString(Mid(StrIn, iCh + 1))
      Exit Function
      End If
   Next iCh

End Function

Private Sub CommandButton2_Click()
'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("Held Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
 Exit Sub
Else
End If

Call heldLinesProcess

End Sub





Private Sub CommandButton22_Click()
Call marketplaceCleanUp
        
End Sub

Private Sub CommandButton23_Click()
Call update_field_reference(True)
End Sub



Private Sub CommandButton4_Click()
'Supplier_Stats Macro

'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("Supplier Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
 Exit Sub
Else
End If

'Delete Top 4 Rows
Rows("1:4").Select
Selection.Delete

'Auot Fits All Columns
Cells.Select
With Selection
.WrapText = False
End With
   Cells.EntireColumn.AutoFit
   Cells.EntireRow.AutoFit
   

'Hide Cells
Range("A:G,L:X,AC:AV").Select
Selection.EntireColumn.Hidden = True


'Add additional columns
Range("AW1").value = "Account Creation"
Range("AX1").value = "Account Amendment"
Range("AY1").value = "Site Creation"
Range("AZ1").value = "Site Amendment"
Range("BA1").value = "Result"
Range("BB1").value = "Update Method"

Range("AB1").Select

'Copy and paste format to additional columns
Range("AB:AB").Copy

    Range("AW:BB").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
'Auto Fit AW:BB
Columns("AW:BB").Select
    Columns("AW:BB").EntireColumn.AutoFit

'Add Formula to AW2
Range("AW2").Formula = "=IF(AND(H2>=J2,H2>=Y2,H2>=AA2),""YES"",""NO"")"

Dim lastRow As Long
lastRow = ActiveSheet.Cells(Rows.count, "H").End(xlUp).row

'Auto Fill All Cells in AW
Range("AW2").Select
    Selection.AutoFill Destination:=Range("AW2:AW$" & lastRow)

'Add Formula to AX2 and auto Fills Cells
Range("AX2").Formula = "=IF(AND(J2>=H2,J2>=Y2,J2>=AA2),""YES"",""NO"")"
Range("AX2").Select
    Selection.AutoFill Destination:=Range("AX2:AX$" & lastRow)
    
'Add Formula to AY2 and auto Fills Cells
Range("AY2").Formula = "=IF(AND(Y2>=H2,Y2>=J2,Y2>=AA2),""YES"",""NO"")"
Range("AY2").Select
    Selection.AutoFill Destination:=Range("AY2:AY$" & lastRow)
    
'Add Formula to AZ2 and auto Fills Cells
Range("AZ2").Formula = "=IF(AND(AA2>=H2,AA2>=J2,AA2>=Y2),""YES"",""NO"")"
Range("AZ2").Select
    Selection.AutoFill Destination:=Range("AZ2:AZ$" & lastRow)

'Add Filter
Range("BB1").Select
    Selection.AutoFilter


'Add Site Amendment Formula and Auto Fills Cells to "BA"
Range("BA2").Formula = "=IF(AND(AW2=""NO"",AX2=""NO"",AY2=""NO"",AZ2=""YES""),""SITE AMENDMENT"","""")"
Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA$" & lastRow)

'Filter For Blanks in Filed 53 (Column 53)
ActiveSheet.Range("BA1:BA$" & lastRow).AutoFilter Field:=53, Criteria1:="="

'Clear Formula From "BA"
Range("BA2:BA$" & lastRow).Select
Selection.ClearContents

'Go to next Active Cell
Range("BA1").Select
ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
    Loop

    
'Add Account Creation Formula to "BA"
ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-4]=""YES"",RC[-3]=""YES"",RC[-2]=""YES"",RC[-1]=""YES""),""ACCOUNT CREATION"","""")"
    Selection.Copy
Range("BA$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Reapply Filter and Clears formula from "BA"
ActiveSheet.AutoFilter.ApplyFilter
Range("BA2:BA$" & lastRow).Select
Selection.ClearContents

'Go to next Active Cell
Range("BA1").Select
ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
    Loop
    
'Add Account Amendment Formula
 ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-4]=""NO"",RC[-3]=""YES"",RC[-2]=""NO"",RC[-1]=""NO""),""ACCOUNT AMENDMENT"","""")"
Selection.Copy
Range("BA$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Reapply Filter and Clears formula from "BA"
ActiveSheet.AutoFilter.ApplyFilter
Range("BA2:BA$" & lastRow).Select
Selection.ClearContents

'Go to next Active Cell
Range("BA1").Select
ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
    Loop

'Add Site Creation Formula x1
ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-4]=""NO"",RC[-3]=""YES"",RC[-2]=""YES"",RC[-1]=""YES""),""SITE CREATION"","""")"
Selection.Copy
Range("BA$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Reapply Filter and clears forumla from "BA"
ActiveSheet.AutoFilter.ApplyFilter
Range("BA2:BA$" & lastRow).Select
Selection.ClearContents

'Go to next Active Cell
Range("BA1").Select
ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
    Loop

'Add Site Creation Formula x2
 ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-4]=""NO"",RC[-3]=""NO"",RC[-2]=""YES"",RC[-1]=""YES""),""SITE CREATION"","""")"
Selection.Copy
Range("BA$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Reapply Filter and Clears formula from "BA"
ActiveSheet.AutoFilter.ApplyFilter
Range("BA2:BA$" & lastRow).Select
Selection.ClearContents

'Go to next Active Cell
Range("BA1").Select
ActiveCell.Offset(1, 0).Select
    Do Until ActiveCell.EntireRow.Hidden = False
    ActiveCell.Offset(1, 0).Select
    Loop


'Add Site Amendment Formula x2
ActiveCell.FormulaR1C1 = _
        "=IF(AND(RC[-4]=""NO"",RC[-3]=""YES"",RC[-2]=""NO"",RC[-1]=""YES""),""SITE AMENDMENT"","""")"

Selection.Copy
Range("BA$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Clear Filter
Application.CutCopyMode = False
    ActiveSheet.ShowAllData

'Auto Fits All Columns
Cells.Select
With Selection
.WrapText = False
End With
   Cells.EntireColumn.AutoFit
   Cells.EntireRow.AutoFit
   
'Hide Cells
Range("A:G,L:X,AC:AV").Select
Selection.EntireColumn.Hidden = True

'Apply BB Formula
Range("BB2").Select
   ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-26]=""CONVERSION"",RC[-26]=""GSSCENTRAL""),""AUTOMATED"",""MANUAL"")"
Selection.Copy
Range("BB$" & lastRow).Select
Range(Selection, Selection.End(xlUp)).Select
ActiveSheet.Paste

'Go To Top Right
Range("H1").Select




'DP Notes
'$
'Dim lastRow As Long
'lastRow = ActiveSheet.Cells(Rows.Count, "H").End(xlUp).Row
'Range("AW" & lastRow).Select
'Application.CutCopyMode = False
'ActiveCell.FormulaR1C1 = _
'"=IF(AND(RC[-4]=""NO"",RC[-3]=""YES"",RC[-2]=""NO"",RC[-1]=""YES""),""SITE AMENDMENT"","""")"
End Sub

Private Sub CommandButton5_Click()
'Active Pricing Macro

'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("Active Pricing", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
Exit Sub
Else
End If

'Sets up Short Cuts and Format
Dim L As Integer
Dim P As Date
Dim Y As Date
P = VBA.Format(Now, "dd/mm/yyyy")
L = 2

'Adds Filter
Range("A1:K1").Select
Selection.AutoFilter


'Filters Column "H" Oldest to newest
 Worksheets(1).AutoFilter.Sort. _
        SortFields.Add Key:=Range("H1"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(1).AutoFilter. _
        Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
       End With
       
       
'Deletes any Pricing that has already expired
Y = 1
Do Until Y >= P Or Y = "00:00:00"
 
Y = Range("H" & L)
Y = VBA.Format(Y, "dd/mm/yyyy")
L = L + 1
 
Loop
 
If L > 3 Then
L = L - 2
 
Range("2:" & L).Delete Shift:=xlUp

End If

'Filters Feild 3 (Column "C") for Obsolete
ActiveSheet.Range("$A$1").AutoFilter Field:=3, Criteria1:= _
        "*OBSOLETE*"
        
'Delete Lines
Range("2:10000").SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete

ActiveSheet.ShowAllData

'Filters for £0 in column E and deletes them
ActiveSheet.Range("$A$9").AutoFilter Field:=5, Criteria1:= _
        "0"
        
Range("2:10000").SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    

'Clears Filter
ActiveSheet.ShowAllData

'Unwrap Text
Columns("C:K").Select
With Selection
        .HorizontalAlignment = xlGeneral
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
'Auto Fit Columns
Columns("C:K").Select
    Selection.ColumnWidth = 147.86
    Columns("C:K").EntireColumn.AutoFit


'Formats column "E" to be currency £
    Range("E:E").Select
Selection.NumberFormat = "$#,##0.00"

'Updates Format
Columns("G:H").Select
    Selection.NumberFormat = "dd-mmm-yyyy"
    
Range("A1").Select



End Sub

Private Sub CommandButton6_Click()

'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("MDT Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
 Exit Sub
Else
End If

'Show MDT Form
MDT.Show

End Sub

Private Sub Frame10_Click()
'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("Held Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
 Exit Sub
Else
End If

Call heldLinesProcess
End Sub

Private Sub Frame11_Click()
Call rmaCheck
End Sub

Private Sub Frame12_Click()
Call productWebadiProcess
End Sub

Private Sub Frame13_Click()
Call MHRARevalProcess
End Sub

Private Sub Frame14_Click()
Call marketplaceCleanUp
End Sub

Private Sub Frame15_Click()
Call sub_delivery_recon
End Sub

Private Sub activePricingButton_Click()
Call activePricing
End Sub

Private Sub btEisaiButton_Click()
Call btPostcodeCheckEisai
End Sub

Private Sub btPostcodeButton_Click()
Call btPostcodeCheck
End Sub

Private Sub deliveryReconButton_Click()
Call sub_delivery_recon
End Sub

Private Sub Frame16_Click()
MacroForm.Hide
excelShortCuts.Show

End Sub


Private Sub Frame6_Click()
'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("MDT Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
 Exit Sub
Else
End If

'Show MDT Form
MDT.Show

End Sub
Private Sub Label10_Click()
'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("Held Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
If answer = vbNo Then
 Exit Sub
Else
End If

Call heldLinesProcess
End Sub

Private Sub Label11_Click()
Call rmaCheck
End Sub

Private Sub Label12_Click()
Call productWebadiProcess
End Sub

Private Sub Label13_Click()
Call MHRARevalProcess
End Sub

Private Sub Label14_Click()
Call marketplaceCleanUp
End Sub

Private Sub Label15_Click()
Call sub_delivery_recon
End Sub

Private Sub heldReportButton_Click()
Call heldLinesProcess
End Sub

Private Sub Label16_Click()
MacroForm.Hide
excelShortCuts.Show
End Sub

Private Sub Label17_Click()
MacroForm.Hide
rateCardForm.Show
End Sub

Private Sub Label3_Click()
Call nhsScotland
End Sub

Private Sub Label4_Click()
Call activePricing
End Sub

Private Sub Label5_Click()
Call priceConfirmation
End Sub

Private Sub Label6_Click()
Call priceExpiry
End Sub

Private Sub Label8_Click()
Call btPostcodeCheckEisai
End Sub

Private Sub Label9_Click()
Call btPostcodeCheck
End Sub

Private Sub mdtReportLabel_Click()
'This Checks to make sure you are about to use the correct Macro
answer = MsgBox("MDT Report", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")

If answer = vbNo Then
 Exit Sub
Else
End If

'Show MDT Form
MDT.Show

End Sub

Private Sub MHRAButton_Click()
Call MHRARevalProcess
End Sub


Private Sub monthEndFrame_Click()
Call monthEndProducts
End Sub

Private Sub monthEndLabel_Click()
Call monthEndProducts

End Sub

Private Sub mdtReportButton_Click()
MDT.Show
End Sub

Private Sub mhraRevalidationButton_Click()
Call MHRARevalProcess
End Sub

Private Sub monthEndButton_Click()
Call monthEndProducts
End Sub

Private Sub mpCleanUpButton_Click()
Call marketplaceCleanUp
End Sub

Private Sub nhsImmformButton_Click()
Call nhsImmform
End Sub

Private Sub nhsScotlandButton_Click()
Call nhsScotland
End Sub

Private Sub productWebadiButton_Click()
Call productWebadiProcess
End Sub

Private Sub priceExpiryButton_Click()
Call priceExpiry
End Sub

Private Sub pricingConfirmationButton_Click()
Call priceConfirmation
End Sub

Private Sub rateCardButton_Click()
rateCardFrm.Show
End Sub



Private Sub rmaReportButton_Click()
Call rmaCheck
End Sub

Private Sub supplierReportButton_Click()
Call supplierReport
End Sub

Private Sub weeklyPricingButton_Click()
Call weeklyPricing
End Sub
