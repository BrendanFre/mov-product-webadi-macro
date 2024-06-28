Option Explicit
'Defining workbooks as either the webadi or the request form.
Public lngSetupFormItems As Long    'Count of items on the setup form.
Public lngWebadiFormItems As Long   'Count of existing lines on the webadi form.
Public lngLastProduct As Long   'Line number of the last product on the setup form.
Public lngFirstEmpty As Long    'Line number of the first empty cell on the webadi.
Public lngCurrentItem As Long   'Current line on the product setup that's being extracted.
Public i            As Long 'A counter
Public strSKUCode   As String        'Current item's SKU code.
Public strSKUDesc   As String        'Current item's SKU description.
Public strEANCode   As String        'Current item's EAN code.
Public strCountry   As String
Public strInnerPk   As String        'Current item's inner pack quantity.
Public strSalesMulti As String        'Current item's sales multiple.
Public strStorageCond As String        'Current item's storage conditions.
Public strShelfDays As String        'Current item's shelf days.
Public strQHLock    As String        'Current item's QH lock status.
Public strBatchManaged As String        'Current item's batch managed status
Public strWarehouse As String        'Current item's warehouse (V01 will always be WT)
Public strOrderability As String        'Current item's orderability
Public productClass As String
Public strActiveWarehouse As String
Public strClientName As String
Public strClientNumber As String
Public strPurchasePrice As String
Public wbkWebadiForm As Workbook        'Webadi workbook selected by user.
Public wksWebadiForm As Worksheet        'Webadi worksheet selected by user.
Public wbkProductSetup As Workbook        'Product setup form completed by client's workbook.
Public wksProductSetup As Worksheet        'Product setup form completed by client's worksheet.
Public lngCurrentCell As Long
Public blnFullService As Boolean
Public blnNHS       As Boolean
Public varFieldsWebadi As Variant

Sub productWebadiProcess()
    '*******************************************************************************
    '   Purpose: To display the ProductWebAdiForm
    '   Inputs: Mouse click
    '   Outputs: productWebAdiForm showing.
    '*******************************************************************************
    'TO DO
    'Add orderability
    
    ProductWebAdiForm.Show
End Sub

Public Sub secondStage1(client_name, client_number, webadi_sheet, product_form, warehouse, orderability)
    '*******************************************************************************
    '   Purpose: Extract data from the product setup form and copy it over to the
    '   the webadi worksheet.
    '   Inputs: Webadi Worksheet & workbook, product setup worksheet & workbook.
    '   productWebAdiForm inputs.
    '   Outputs: A completed webadi form.
    '   Last Update: 28/02/2024
    '   Last Updated by: Brendan
    '   Update Description: Fixed to work with new forms
    '*******************************************************************************
    Dim finalOutput As String
    
    Call speedUp(True)
    
    'Assigning workbooks to variables
    Set wbkWebadiForm = Workbooks(webadi_sheet)
    Set wksWebadiForm = wbkWebadiForm.Worksheets("Sheet1")
    Set wbkProductSetup = Workbooks(product_form)
    
    Call update_wksProductSetup(wbkProductSetup)
    
    'Assigning non-object variables
    lngSetupFormItems = count_items("setup", wksProductSetup)
    lngWebadiFormItems = count_items("webadi", wksWebadiForm)
    'lngLastProduct = 13 + lngSetupFormItems
    'lngCurrentItem = 13
    blnFullService = False
    
    varFieldsWebadi = update_field_reference(blnNHS)
    
   ' If (blnNHS) Then
    '    lngSetupFormItems = lngSetupFormItems - 1
   ' End If
    
    For i = 1 To lngSetupFormItems
        
        strClientName = client_name
        strClientNumber = client_number
        strOrderability = orderability
        strSKUCode = UCase(wksProductSetup.Range(varFieldsWebadi(0) & lngCurrentItem).value)
        strSKUDesc = UCase(wksProductSetup.Range(varFieldsWebadi(1) & lngCurrentItem).value)
        strCountry = (wksProductSetup.Range(varFieldsWebadi(2) & lngCurrentItem).value)
        strEANCode = UCase(wksProductSetup.Range(varFieldsWebadi(3) & lngCurrentItem).value)
        strInnerPk = UCase(wksProductSetup.Range(varFieldsWebadi(4) & lngCurrentItem).value)
        If blnNHS = True Then
            strPurchasePrice = wksProductSetup.Range(varFieldsWebadi(12) & lngCurrentItem).value
            strSalesMulti = ""
        Else
            strSalesMulti = UCase(wksProductSetup.Range(varFieldsWebadi(5) & lngCurrentItem).value)
        End If
        
        strStorageCond = define_conditions(wksProductSetup.Range(varFieldsWebadi(6) & lngCurrentItem).value)
        strShelfDays = UCase(wksProductSetup.Range(varFieldsWebadi(7) & lngCurrentItem).value)
        strQHLock = UCase(wksProductSetup.Range(varFieldsWebadi(8) & lngCurrentItem).value)
        strBatchManaged = define_managed(wksProductSetup.Range(varFieldsWebadi(9) & lngCurrentItem).value)
        productClass = define_classification(wksProductSetup.Range(varFieldsWebadi(10) & lngCurrentItem).value, wksProductSetup.Range(varFieldsWebadi(11) & lngCurrentItem).value)
        
        lngWebadiFormItems = count_items("webadi", wksWebadiForm)
        lngFirstEmpty = 4 + lngWebadiFormItems
        If warehouse(0) = "S01" Or warehouse(1) = "S02" Or warehouse(2) = "S03" Or warehouse(3) = "S04" Or warehouse(4) = "S05" Then
            blnFullService = True
            lngWebadiFormItems = count_items("webadi", wksWebadiForm)
            lngFirstEmpty = 5 + lngWebadiFormItems
            If warehouse(0) = "S01" Then
                lngWebadiFormItems = count_items("webadi", wksWebadiForm)
                lngFirstEmpty = 5 + lngWebadiFormItems
                Call fillWebAdi("S01", lngFirstEmpty)
            End If
            
            If warehouse(1) = "S02" Then
                lngWebadiFormItems = count_items("webadi", wksWebadiForm)
                lngFirstEmpty = 5 + lngWebadiFormItems
                Call fillWebAdi("S02", lngFirstEmpty)
            End If
            
            If warehouse(2) = "S03" Then
                lngWebadiFormItems = count_items("webadi", wksWebadiForm)
                lngFirstEmpty = 5 + lngWebadiFormItems
                Call fillWebAdi("S03", lngFirstEmpty)
            End If
            
            If warehouse(3) = "S04" Then
                lngWebadiFormItems = count_items("webadi", wksWebadiForm)
                lngFirstEmpty = 5 + lngWebadiFormItems
                Call fillWebAdi("S04", lngFirstEmpty)
            End If
            
            If warehouse(4) = "S05" Then

                Call fillWebAdi("S05", lngFirstEmpty)
            End If
            
        Else
            lngWebadiFormItems = count_items("webadi", wksWebadiForm)
            lngFirstEmpty = 5 + lngWebadiFormItems
            Call fillWebAdi("V01", lngFirstEmpty)
            blnFullService = False
        End If
        lngCurrentItem = lngCurrentItem + 1
        
    Next i
    
    wbkWebadiForm.Sheets(1).Activate
    ActiveSheet.Cells(1, 1).Select
    
    'Once completed display pop-up.
    
    Call speedUp(False)

    finalOutput = checkFormVersion(wksProductSetup)

    MsgBox finalOutput

End Sub

Sub fillWebAdi(currentWarehouse, first_empty)
    
    'Extract data from product setup form.
    strWarehouse = currentWarehouse
    
    wksWebadiForm.Range("B" & (first_empty)).value = "O"
    wksWebadiForm.Range("C" & (first_empty)).value = "Create"
    wksWebadiForm.Range("D" & (first_empty)).value = "Auto Numbering"
    wksWebadiForm.Range("E" & (first_empty)).value = strSKUCode
    wksWebadiForm.Range("G" & (first_empty)).value = strClientName
    wksWebadiForm.Range("K" & (first_empty)).value = strWarehouse
    wksWebadiForm.Range("H" & (first_empty)).value = strSKUDesc
    wksWebadiForm.Range("I" & (first_empty)).value = productClass
    wksWebadiForm.Range("J" & (first_empty)).value = strSalesMulti
    wksWebadiForm.Range("M" & (first_empty)).value = strInnerPk
    If (blnNHS) Then
        wksWebadiForm.Range("N" & (first_empty)).value = strPurchasePrice
    End If
    wksWebadiForm.Range("O" & (first_empty)).value = strBatchManaged
    wksWebadiForm.Range("P" & (first_empty)).value = strStorageCond
    wksWebadiForm.Range("Q" & (first_empty)).value = strShelfDays
    wksWebadiForm.Range("R" & (first_empty)).value = StrConv(strQHLock, vbProperCase)
    If (strCountry = "UK") Then
        wksWebadiForm.Range("AQ" & (first_empty)).value = "United Kingdom"
    Else
        wksWebadiForm.Range("AQ" & (first_empty)).value = strCountry
    End If
    wksWebadiForm.Range("AV" & (first_empty)).value = "Un-Owned Inventory Items"
    wksWebadiForm.Range("AY" & (first_empty)).value = "EA"
    wksWebadiForm.Range("BD" & (first_empty)).value = "Active"
    
    'Check if client is FS or WT and only populate GTIN if FS.
    If blnFullService Then
        wksWebadiForm.Range("V" & (first_empty)).value = strEANCode
    End If
    
    wksWebadiForm.Range("W" & (first_empty)).value = "OM Client"
    wksWebadiForm.Range("X" & (first_empty)).value = strClientNumber & "|" & strClientName & "|" & "OM Client"
    wksWebadiForm.Range("Y" & (first_empty)).value = "OM Order Management"
    wksWebadiForm.Range("Z" & (first_empty)).value = define_orderability(strOrderability)
    wksWebadiForm.Range("CG" & (first_empty)).value = "MFR-CLIENT"
    wksWebadiForm.Range("CI" & (first_empty)).value = strSKUCode
    
End Sub

Public Function define_classification(controlledDrug, productClass)
    If controlledDrug = "Non-CD" Then
        Select Case productClass
            Case "POM", "POM-V"
                define_classification = "Prescription only medicines"
            Case "BIOLOGICAL"
                define_classification = "Biological products"
            Case "GSL"
                define_classification = "General sales list"
            Case "HERBAL"
                define_classification = "Herbal"
            Case "immunoglobulin"
                define_classification = "Immunoglobulin"
            Case "Homeopathics"
                define_classification = "Homeopathics"
            Case "ULT"
                define_classification = "Select Ultra low temperature "
            Case "Medical Device"
                define_classification = " Medical device "
            Case "Other", "OTHER", "N/A"
                define_classification = "N/A"
        End Select
    Else
        define_classification = "CD" & controlledDrug
    End If
    
End Function

Public Function count_items(spreadsheet_name, sheet)
    Dim file_contents As Boolean
    Dim lngNumItems As Long
    Dim lngCellCounter
    
    file_contents = True
    If spreadsheet_name = "setup" Then
        lngCellCounter = 4
    ElseIf spreadsheet_name = "webadi" Then
        lngCellCounter = 5
    End If
    
    lngNumItems = 0
    
    'Check product form for number of products
    While file_contents = True
        If Not isEmpty(sheet.Cells(lngCellCounter, 2)) Then
            lngCellCounter = lngCellCounter + 1
            lngNumItems = lngNumItems + 1
        Else
            file_contents = False
        End If
    Wend
    
    count_items = lngNumItems
    
End Function
Public Function define_conditions(value)
    
    Select Case value
        Case "Ambient Temp Controlled        'c", "ATC", "Ambient Temp Controlled 15-25'c"
            define_conditions = "ATC"
        Case "Chilled        'c", "Chilled", "Chilled 2-8'c"
            define_conditions = "Chilled"
        Case "Freezer", "Freezer"
            define_conditions = "Freezer"
        Case "CD Vault        'c", "CD", "CD Vault 15-25'c"
            define_conditions = "Controlled Drug"
        Case "Ambient (Non-Temp controlled)", "Ambient"
            define_conditions = "Ambient"
        Case Else
            define_conditions = "Failed"
    End Select
    
End Function

Public Function define_managed(value)
    If value = "Yes" Or value = "Yes " Then
        define_managed = "Un-Owned Inventory Lot UK"
    ElseIf value = "No " Then
        define_managed = "Un-Owned Inventory UK"
    End If
    
End Function

Public Function define_orderability(orderability)
    If orderability = "" Then
        define_orderability = "Default.|Default|OM Order Management"
    Else
        define_orderability = orderability
    End If
    
End Function

Public Function update_field_reference(isNhs) As Variant
    
    Dim tempArray(13) As Variant
    tempArray(0) = "B"        'SKU Code
    tempArray(1) = "C"        'SKU Description
    If Not isNhs Then
        tempArray(2) = "AG"        'Country of Origin
        tempArray(3) = "D"        'EAN Code
        tempArray(4) = "AH"        'Inner Pack qty
        tempArray(5) = "AJ"        'Sales multiple
        tempArray(6) = "T"        'Storage conditions
        tempArray(7) = "Y"        'Shelf days
        tempArray(8) = "Z"        'QH Lock
        tempArray(9) = "AA"        'Batch Managed
        tempArray(10) = "U"        'Controlled Drug
        tempArray(11) = "X"        'Product Classification
    Else
        tempArray(2) = "X"        'Country of Origin
        tempArray(3) = "D"        'EAN Code
        tempArray(4) = "F"        'Inner Pack qty
        tempArray(5) = "AA"        'Sales multiple
        tempArray(6) = "G"        'Storage conditions
        tempArray(7) = "L"        'Shelf days
        tempArray(8) = "M"        'QH Lock
        tempArray(9) = "N"        'Batch Managed
        tempArray(10) = "K"        'Controlled Drug
        tempArray(11) = "H"        'Product Classification
        tempArray(12) = "E"        'Purchase Price
        
    End If
    
    Dim i           As Variant
    For Each i In tempArray()
        
    Next i
    
    update_field_reference = tempArray
    
End Function
Sub update_wksProductSetup(productSetupForm)
    Dim sht         As Worksheet
    Dim wkbProduct  As Workbook
    Dim wksProduct  As Worksheet
    
    Set wkbProduct = productSetupForm
    Set wksProduct = wkbProduct.Sheets(1)
    
    For Each sht In productSetupForm.Worksheets
        If ((InStr(sht.Name, "Product"))) Then
            Set wksProductSetup = wbkProductSetup.Worksheets(sht.Name)
        End If
    Next sht
    
    If wksProductSetup.Range("AD2").value = "SCCL Field" Then
        blnNHS = True
        lngCurrentItem = 4
        lngLastProduct = 4 + lngSetupFormItems
    Else
        blnNHS = False
        lngCurrentItem = 4
        lngLastProduct = 4 + lngSetupFormItems
        
    End If
    
    Debug.Print (wksProductSetup.Name)
    
End Sub

Function checkFormVersion(templateSheet)
    Dim headerString As String
    Dim standardForm As String
    Dim pippForm As String
    Dim standardFormVersion As String
    Dim pippFormVersion As String
    Dim outputMessage As String
    headerString = templateSheet.PageSetup.RightHeader
    standardForm = "Document No.:MO-UK-WA-002-FO-001 "
    pippForm = "Document No.: MO-UK-WA-002-FO-004"
    standardFormVersion = "5.0"
    pippFormVersion = "4.0"
    If InStr(headerString, standardForm) >= 1 Then
        If InStr(headerString, standardFormVersion) >= 1 Then
            outputMessage = "Standard product creation completed."
        Else
            outputMessage = "Standard product creation failed, not laatest" & _
            "version of product form."
        End If
    ElseIf InStr(headerString, pippForm) >= 1 Then
        If InStr(headerString, pippFormVersion) >= 1 Then
            outputMessage = "PIPP product creation completed."
        Else
            outputMessage = "PIPP product creation failed, not latest" & _
            "version of product form."
        End If
    Else
        outputMessage = "Form not recognised."
    End If

    checkFormVersion = outputMessage



End Function

