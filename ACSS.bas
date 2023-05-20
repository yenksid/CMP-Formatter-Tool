Attribute VB_Name = "ACSS"
'Convert the ACS format to a more redable format
Public Sub ACS()
    Dim WK As Workbook, arr As Variant, indent_level, name1, name2, designation, aadress, community, sstate, zip, text, phone As String
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    RemoveSection
    
    'Asigning the data to an array
    arr = WK.Worksheets("ACS").Range("A1").CurrentRegion
    
    'From the first item in the array to the last
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        'Class Of Service
        class_of_service = Mid(arr(i, 1), 1, 1)
        WK.Worksheets("ACS").Range("A" & i).Value = class_of_service
        
        'Indent Level
        indent_level = Trim(Mid(arr(i, 1), 3, 1))
        WK.Worksheets("ACS").Range("B" & i).Value = indent_level
        
        'Name 1
        name1 = Trim(Mid(arr(i, 1), 5, 75))
        
        'Name 2
        name2 = Trim(Mid(arr(i, 1), 81, 52))
        
        'Designation
        designation = Trim(Mid(arr(i, 1), 133, 20))
        
        'Move extra lines to cell C
        WK.Worksheets("ACS").Range("C" & i).Value = Trim(name1) & " " & Trim(name2) & " " & Trim(designation)
        
'        'Address
        
        'street_number
        street_number = Trim(Mid(arr(i, 1), 154, 56))
        street_number = Left(street_number, InStr(1, street_number, " "))
        street_name = Trim(Mid(arr(i, 1), 154, 56))
            If street_number <> vbNullString Then
                street_number_limpio = RemoveLettersAndCheckNumber(street_number)
                    If IsNumeric(street_number_limpio) And street_number_limpio <> 0 Then
                        WK.Worksheets("ACS").Range("D" & i).Value = Trim(street_number)
                        street_number = ""
                    Else
                        WK.Worksheets("ACS").Range("E" & i).Value = Trim(street_name)
                    End If
            End If
   

        'street_name
        street_name = Trim(Mid(arr(i, 1), 154, 56))
        If street_name <> vbNullString Then
            If WK.Worksheets("ACS").Range("D" & i).Value2 <> vbNullString Then
                street_name = Right(street_name, Len(street_name) - Len(WK.Worksheets("ACS").Range("D" & i).Value2))
                WK.Worksheets("ACS").Range("E" & i).Value = Trim(street_name)
            Else
                WK.Worksheets("ACS").Range("E" & i).Value = Trim(street_name)
            End If
        End If
        
        'cardinales
        'Fixed to make the code read straight from the cell
        If Left(WK.Worksheets("ACS").Range("E" & i).Value, 2) = "N " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 2) = "E " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 2) = _
            "S " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 2) = "W " Or _
            Left(WK.Worksheets("ACS").Range("E" & i).Value, 3) = "NE " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 3) = _
                "NW " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 3) = "SE " Or Left(WK.Worksheets("ACS").Range("E" & i).Value, 2) = "SW " Then
                WK.Worksheets("ACS").Range("F" & i).Value2 = Trim(Left(WK.Worksheets("ACS").Range("E" & i).Value, 2))
                WK.Worksheets("ACS").Range("E" & i).Value2 = _
                Trim(Right(WK.Worksheets("ACS").Range("E" & i).Value, Len(WK.Worksheets("ACS").Range("E" & i).Value) - Len(WK.Worksheets("ACS").Range("F" & i).Value2)))
        End If
        
        'Replace street names without the cardinals
        'Moved this line to be executed alongside with the cardinals code
'        WK.Worksheets("ACS").Range("E" & i).Value2 = _
                Trim(Right(WK.Worksheets("ACS").Range("E" & i).Value, Len(WK.Worksheets("ACS").Range("E" & i).Value) - Len(WK.Worksheets("ACS").Range("F" & i).Value2)))
        
        
        'Community
        community = Trim(Mid(arr(i, 1), 210, 30))
        WK.Worksheets("ACS").Range("G" & i).Value = Trim(community)

        'Text : NSTN, See
        text = Trim(Mid(arr(i, 1), 240, 51))
        
        If Left(text, 3) = "See" Then
            WK.Worksheets("ACS").Range("C" & i).Value = WK.Worksheets("ACS").Range("C" & i).Value & " " & Trim(text)
        Else
            WK.Worksheets("ACS").Range("J" & i).Value = Trim(text)
        End If

        'Zip
        If Right(WK.Worksheets("ACS").Range("G" & i).Value, 6) Like " #####" Then
            'Get zip
            WK.Worksheets("ACS").Range("I" & i).Value = Right(WK.Worksheets("ACS").Range("G" & i).Value, 5)
            'Remove zip from the previos cell
            WK.Worksheets("ACS").Range("G" & i).Value = Left(WK.Worksheets("ACS").Range("G" & i).Value, _
                Len(WK.Worksheets("ACS").Range("G" & i).Value) - Len(WK.Worksheets("ACS").Range("I" & i).Value) - 1)
        ElseIf Right(WK.Worksheets("ACS").Range("G" & i).Value, 11) Like " #####[-]####" Then
            WK.Worksheets("ACS").Range("I" & i).Value = Right(WK.Worksheets("ACS").Cells(i, 13).Value2, 10)
            WK.Worksheets("ACS").Range("G" & i).Value = Left(WK.Worksheets("ACS").Range("G" & i).Value, _
                Len(WK.Worksheets("ACS").Range("G" & i).Value) - Len(WK.Worksheets("ACS").Range("I" & i).Value) - 1)
        ElseIf Right(WK.Worksheets("ACS").Range("G" & i).Value, 10) Like " #########" Then
            WK.Worksheets("ACS").Range("I" & i).Value = Right(WK.Worksheets("ACS").Range("G" & i).Value, 9)
            WK.Worksheets("ACS").Range("G" & i).Value = Left(WK.Worksheets("ACS").Range("G" & i).Value, _
                Len(WK.Worksheets("ACS").Range("G" & i).Value) - Len(WK.Worksheets("ACS").Range("I" & i).Value) - 1)
        ElseIf IsNumeric(WK.Worksheets("ACS").Range("G" & i).Value) And Len(WK.Worksheets("ACS").Range("G" & i).Value) = 5 Then
            WK.Worksheets("ACS").Range("I" & i).Value = WK.Worksheets("ACS").Range("G" & i).Value
            WK.Worksheets("ACS").Range("G" & i).Value = vbNullString
        End If

        'States
        StateVal = Right(WK.Worksheets("ACS").Range("G" & i).Value, 3)
            Select Case UCase(StateVal)
                Case " AL", " AK", " AZ", " AR", " CA", " CO", " CT", " DE", _
                    " DC", " FL", " GA", " HI", " ID", " IL", " IN", " IA", " KS", _
                    " KY", " LA", " ME", " MD", " MA", " MI", " MN", " MS", _
                    " MO", " MT", " NE", " NV", " NH", " NJ", " NM", " NY", " NC", _
                    " ND", " OH", " OK", " OR", " PA", " RI", " SC", " SD", _
                    " TN", " TX", " UT", " VT", " VA", " WA", " WV", " WI", " WY"
                    'Get state
                    WK.Worksheets("ACS").Range("H" & i).Value = Right(WK.Worksheets("ACS").Range("G" & i).Value, 2)
                    'Remove state from the previos cell
                    WK.Worksheets("ACS").Range("G" & i).Value = Left(WK.Worksheets("ACS").Range("G" & i).Value, _
                        Len(WK.Worksheets("ACS").Range("G" & i).Value) - 3)
                Case Else
            End Select
            
        'Right Aligned Text
        

        'Phone
        phone = Mid(arr(i, 1), 291, 10)
        WK.Worksheets("ACS").Range("L" & i).Value = Trim(Replace(Replace(Replace(phone, "-", ""), "(", ""), ")", ""))
        
    Next i
    
    RemovePLA
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "ACS-CATS Conversion Completed, Thanks!"
    
End Sub

Sub RemovePLA()
    Dim PLA As String
    Dim rng As Range
    Dim cell As Range
    
    PLA = "P" ' Change to the value you want to remove
    
    ' Set the range where you want to search for the target value
    Set rng = Range("B:B") ' Modify the range as needed
    
    ' Loop through each cell in the range
    For Each cell In rng
        If cell.Value = PLA Then
            ' Remove the entire row if the target value is found
            cell.EntireRow.Delete
        End If
    Next cell
End Sub

Sub RemoveSection()
    Dim section As String
    Dim rng As Range
    Dim cell As Range
    
    section = "E00" ' Change to the value you want to remove
    
    ' Set the range where you want to search for the target value
    Set rng = Range("A:A") ' Modify the range as needed
    
    ' Loop through each cell in the range
    For Each cell In rng
        If Left(cell.Value, 3) = section Then
            ' Remove the entire row if the target value is found
            cell.EntireRow.Delete
        End If
    Next cell
End Sub

Function RemoveLettersAndCheckNumber(ByVal word As String) As Boolean
    Dim result As String
    Dim i As Integer
    
    ' Initialize the result as an empty string
    result = ""
    
    ' Loop through each character in the word
    For i = 1 To Len(word)
        ' Check if the character is a letter or a special character
        If Not IsLetterOrSpecialCharacter(Mid(word, i, 1)) Then
            ' If it's not a letter or special character, add it to the result
            result = result & Mid(word, i, 1)
        End If
    Next i
    
    ' Check if the result is a number
    RemoveLettersAndCheckNumber = IsNumeric(result)
End Function

Function IsLetterOrSpecialCharacter(ByVal character As String) As Boolean
    ' Check if the character is a letter (A-Z or a-z) or a special character
    IsLetterOrSpecialCharacter = (Asc(UCase(character)) >= 65 And Asc(UCase(character)) <= 90) Or _
                                (Asc(character) < 48 Or Asc(character) > 57)
End Function


'Clean all the data in the sheet
Sub CleanSheetACS()
    Dim WS As Worksheet
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Get the last row
    lastrow = WK.Worksheets("ACS").Range("A1").CurrentRegion.Rows.Count
    
    'Clean data
    WK.Worksheets("ACS").Range("A1:" & "L" & lastrow).Clear
    
    'save the sheet
    WK.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
