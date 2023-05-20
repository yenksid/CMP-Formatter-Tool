Attribute VB_Name = "IUSS"
Public Sub IUS()
    Dim WK As Workbook, arr As Variant, indent_level, name1, name2, designation, aadress, community, sstate, zip, text, phone As String
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Asigning the data to an array
    arr = WK.Worksheets("IUS").Range("B2").CurrentRegion
    
    'From the first item in the array to the last
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        'Class Of Service
        class_of_service = "B"
        WK.Worksheets("IUS").Range("J" & i).Value = class_of_service

        'Indent Level
        indent_level = 0
        WK.Worksheets("IUS").Range("K" & i).Value = indent_level

        'Name 1
        name1 = WK.Worksheets("IUS").Range("B" & i).Value
        WK.Worksheets("IUS").Range("L" & i).Value = Trim(Replace(name1, ",", ""))
        
'        'Address
'        aaddress = WK.Worksheets("IUS").Range("C" & i).Value
'        WK.Worksheets("IUS").Range("M" & i).Value = Trim(Replace(Replace(aaddress, "  ", " "), ",", ""))
        
        'street_number
        street_number = WK.Worksheets("IUS").Range("C" & i).Value
        If street_number <> vbNullString Then
            street_number = Left(street_number, InStr(1, street_number, " "))

            If IsNumeric(street_number) = True Then
                WK.Worksheets("IUS").Range("M" & i).Value = Trim(Replace(street_number, "  ", " "))
            Else
                WK.Worksheets("IUS").Range("M" & i).Value = vbNullString
            End If
        End If

        'street_name
        street_name = WK.Worksheets("IUS").Range("C" & i).Value
        If street_name <> vbNullString And WK.Worksheets("IUS").Range("M" & i).Value <> vbNullString Then
            street_name = Right(street_name, Len(street_name) - Len(street_number))
            WK.Worksheets("IUS").Range("N" & i).Value = Trim(Replace(street_name, "  ", " "))
        Else
            WK.Worksheets("IUS").Range("N" & i).Value = Trim(Replace(street_name, "  ", " "))
        End If
        
        'cardinales
        cardinales = street_name
        If Left(cardinales, 2) = "N " Or Left(cardinales, 2) = "E " Or Left(cardinales, 2) = "S " Or Left(cardinales, 2) = "W " Or _
            Left(cardinales, 3) = "NE " Or Left(cardinales, 3) = "NW " Or Left(cardinales, 3) = "SE " Or Left(cardinales, 2) = "SW " Then
            WK.Worksheets("IUS").Range("O" & i).Value2 = Trim(Left(cardinales, 2))
        End If
        
        'Replace street names without the cardinals
        WK.Worksheets("IUS").Range("N" & i).Value2 = Trim(Right(WK.Worksheets("IUS").Range("N" & i).Value, Len(WK.Worksheets("IUS").Range("N" & i).Value) - Len(WK.Worksheets("IUS").Range("O" & i).Value2)))

        'Community
        community = WK.Worksheets("IUS").Range("D" & i).Value
        WK.Worksheets("IUS").Range("P" & i).Value = Trim(community)

        'zip
        sstate = WK.Worksheets("IUS").Range("E" & i).Value
        WK.Worksheets("IUS").Range("Q" & i).Value = Trim(sstate)

        'state
        zip = WK.Worksheets("IUS").Range("F" & i).Value
        WK.Worksheets("IUS").Range("R" & i).Value = Trim(zip)
        
        'Right Aligned Text
        
        'Phone
        phone = WK.Worksheets("IUS").Range("G" & i).Value
        If Len(phone) > 5 Then
            WK.Worksheets("IUS").Range("U" & i).Value = Trim(Replace(Replace(Replace(phone, "-", ""), "(", ""), ")", ""))
        Else
            WK.Worksheets("IUS").Range("S" & i).Value = Trim(Replace(Replace(Replace(phone, "-", ""), "(", ""), ")", ""))
        End If
        
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "IUS Conversion Completed, Thanks!"
    
End Sub

'Remove manual sort
Public Sub ClearDataset1()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = ThisWorkbook
    
    Dim i As Long
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Take the last row from the IUS
    lastrow_sort = WK.Worksheets("IUS").Range("J" & Rows.Count).End(xlUp).Row
    
    For i = 1 To lastrow_sort Step 1
        'Remove manual sort
        If WK.Worksheets("IUS").Range("K" & i).Value = "P" Then
            WK.Worksheets("IUS").Range("K" & i).EntireRow.Delete
        End If
        'Remove empty lines
        If WK.Worksheets("IUS").Range("L" & i).Value = vbNullString Or _
            WK.Worksheets("IUS").Range("M" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("N" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("O" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("P" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("Q" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("R" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("S" & i).Value = vbNullString And _
            WK.Worksheets("IUS").Range("T" & i).Value = vbNullString Then
                WK.Worksheets("IUS").Range("A" & i).EntireRow.Delete
        End If
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub


'Clean all the data in the sheet
Sub CleanSheetIUS()
    Dim WS As Worksheet
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Get the last row
    lastrow = WK.Worksheets("IUS").Range("A1").CurrentRegion.Rows.Count
    
    'Clean data
    WK.Worksheets("IUS").Range("A1:" & "L" & lastrow).Clear
    
    'save the sheet
    WK.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

