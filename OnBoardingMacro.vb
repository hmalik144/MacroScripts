Private Sub Cancel_Click()
    startDialog.hide
End Sub
Private Sub cityDropDown_DropButtonClick()
    cityDropDown.AddItem "SYD"
    cityDropDown.AddItem "MEL"
    cityDropDown.AddItem "BNE"
    cityDropDown.AddItem "ADL"
End Sub

Private Sub equipmentDropDown_DropButtonClick()
    equipmentDropDown.AddItem "Microsoft Surface Pro 4"
    equipmentDropDown.AddItem "Microsoft Surface Pro 6"
    equipmentDropDown.AddItem "HP Elitebook 820"
    equipmentDropDown.AddItem "Apple iPhone 6s"
    equipmentDropDown.AddItem "Apple iPhone 6s Plus"
    equipmentDropDown.AddItem "Apple iPhone 8"
    equipmentDropDown.AddItem "Apple iPhone 8 Plus"
End Sub



Private Sub Submit_Click()
    'Field 1'
    Dim fullNameString As Range
    Set fullNameString = ActiveDocument.Bookmarks("fullName").Range
    
    'If fullNameString.Text = "" Then
    fullNameString.Text = Me.nameTextBox.Value
    'Else
    'fullNameString.Delete
    'fullNameString.Text = Me.nameTextBox.Value
    'End If
    
    
    'Field 2'
    Dim cityString As Range
    Set cityString = ActiveDocument.Bookmarks("city").Range
    
    Dim selectedCity As String
    selectedCity = Me.cityDropDown.Value
    
    cityString.Text = selectedCity
    
    'Field 3'
    Dim equipmentString As Range
    Set equipmentString = ActiveDocument.Bookmarks("itEquipment").Range
    
    Dim selectedEquipment As String
    selectedEquipment = Me.equipmentDropDown.Value
    
    equipmentString.Text = selectedEquipment
    
    'Field 4'
    Dim assetString As Range
    Set assetString = ActiveDocument.Bookmarks("assetNumber").Range
    'conditions'
    assetString.Text = Me.assetTextBox.Value
    
     Select Case equipmentString
        Case "Microsoft Surface Pro 4"
        assetString.Text = selectedCity & "-L-" & Me.assetTextBox.Value

        Case "Microsoft Surface Pro 6"
        assetString.Text = selectedCity & "-L-" & Me.assetTextBox.Value
        
        Case "HP Elitebook 820"
        assetString.Text = selectedCity & "-L-" & Me.assetTextBox.Value
        
        Case "Apple iPhone 6s"
        assetString.Text = Me.assetTextBox.Value
        
        Case "Apple iPhone 6s Plus"
        assetString.Text = Me.assetTextBox.Value

        Case "Apple iPhone 8"
        assetString.Text = Me.assetTextBox.Value

        Case "Apple iPhone 8 Plus"
        assetString.Text = Me.assetTextBox.Value

    End Select
    
    'field 5'
    Dim serialString As Range
    Set serialString = ActiveDocument.Bookmarks("serialNumber").Range
    serialString.Text = Me.serialTextBox.Value
    
    'field 6'
    Dim dateString As Range
    Set dateString = ActiveDocument.Bookmarks("dateAccepted").Range
    dateString.Text = Date
    
    startDialog.hide
    
End Sub
