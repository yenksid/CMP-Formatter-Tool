Attribute VB_Name = "Main"
'To add lines we need to restart lines to the previos number:
Sub Main()
    Dim WS As Worksheet, space As String, indent As String * 194, class_of_service As String * 11, street_number As String * 32, street_name As String * 70, cardinal As String * 15, community As String * 45, state As String * 18, zip As String * 13, telephone_number As String * 10, non_std_telno As String * 50, name As String * 377, right_aligned_text As String * 84
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Move captions lines below header name before start
    Call CaptionLines
    
    'Remove listings placlement
    Call ClearDataset
    
'    'Fix issue with CAINS
'    Call FixCAIN
'
'    ''Indent can't be greater than 8, system error.
'    Call FixIndent
    
    'Fix captions with different types
    Call TypeBR
    
    'Get the last row
    lastrow = WK.Worksheets("CMPFormatter").Range("A1").CurrentRegion.Rows.Count
    
    For i = 2 To lastrow
        
        'Get data
        indent = Trim(WK.Worksheets("CMPFormatter").Range("B" & i).Value2)
        class_of_service = Trim(WK.Worksheets("CMPFormatter").Range("A" & i).Value2)
        street_number = Trim(WK.Worksheets("CMPFormatter").Range("D" & i).Value2)
        street_name = Trim(WK.Worksheets("CMPFormatter").Range("E" & i).Value2)
        cardinal = Trim(WK.Worksheets("CMPFormatter").Range("F" & i).Value2)
        community = Trim(WK.Worksheets("CMPFormatter").Range("G" & i).Value2)
        state = Trim(WK.Worksheets("CMPFormatter").Range("H" & i).Value2)
        zip = Trim(WK.Worksheets("CMPFormatter").Range("I" & i).Value2)
        non_std_telno = Trim(WK.Worksheets("CMPFormatter").Range("J" & i).Value2)
        right_aligned_text = Trim(WK.Worksheets("CMPFormatter").Range("K" & i).Value2)
        telephone_number = Trim(WK.Worksheets("CMPFormatter").Range("L" & i).Value2)
        
        If indent = 0 And WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 = 0 Then
            'Only replace first instance in string
            name = Replace(WK.Worksheets("CMPFormatter").Range("C" & i).Value2, " ", "| ", , 1)
            
            'Residential listings with one name
            If InStr(1, name, "|") = False And WK.Worksheets("CMPFormatter").Range("A" & i).Value2 = "R" Then
                name = WK.Worksheets("CMPFormatter").Range("C" & i).Value2 & "|"
            End If
            
            'Xref listings
            If InStr(1, name, "See ") And indent = 0 Then
                name = Replace(Replace(name, "|", ""), " See", "| See")
            End If
            
        Else
            name = Trim(WK.Worksheets("CMPFormatter").Range("C" & i).Value2)
        End If
        
        'Insert 54 spaces
        space = "                                                      "
        
        'Concatenate the data
        WK.Worksheets("CMPFormatter").Range("N" & i).Value2 = space & indent & class_of_service & street_number & street_name & cardinal & community & state & zip & telephone_number & non_std_telno & name & right_aligned_text
        
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

'Clean all the data in the sheet
Sub CaptionLines()
    Dim WS As Worksheet
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    'Get the last row
    lastrow = WK.Worksheets("CMPFormatter").Range("A1").CurrentRegion.Rows.Count
    
    For i = lastrow To 2 Step -1
        'Find caption heads
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("B" & i - 1).Value2 = 0 Then
                'If we have data in the caption head.
                If WK.Worksheets("CMPFormatter").Range("D" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("E" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("F" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("G" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("H" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("I" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("J" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2 <> vbNullString Or _
                    WK.Worksheets("CMPFormatter").Range("L" & i - 1).Value2 <> vbNullString Then
                    
                    'Insert a new row
                    WK.Worksheets("CMPFormatter").Range("B" & i).EntireRow.Insert
                    
                    'copy data from the caption head to a line below
                    WK.Worksheets("CMPFormatter").Range("D" & i).Value2 = WK.Worksheets("CMPFormatter").Range("D" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("E" & i).Value2 = WK.Worksheets("CMPFormatter").Range("E" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("F" & i).Value2 = WK.Worksheets("CMPFormatter").Range("F" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("G" & i).Value2 = WK.Worksheets("CMPFormatter").Range("G" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("H" & i).Value2 = WK.Worksheets("CMPFormatter").Range("H" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("I" & i).Value2 = WK.Worksheets("CMPFormatter").Range("I" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("J" & i).Value2 = WK.Worksheets("CMPFormatter").Range("J" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("K" & i).Value2 = WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("L" & i).Value2 = WK.Worksheets("CMPFormatter").Range("L" & i - 1).Value2
                    
                    'Add class service and indent
                    WK.Worksheets("CMPFormatter").Range("A" & i).Value2 = WK.Worksheets("CMPFormatter").Range("A" & i - 1).Value2
                    WK.Worksheets("CMPFormatter").Range("B" & i).Value2 = 1
                    
                    'Remove data in the caption head
                    WK.Worksheets("CMPFormatter").Range("D" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("E" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("F" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("G" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("H" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("I" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("J" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2 = vbNullString
                    WK.Worksheets("CMPFormatter").Range("L" & i - 1).Value2 = vbNullString
                End If
        End If
    Next i
End Sub

'Clean all the data in the sheet
Sub CleanSheet()
    Dim WS As Worksheet
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Get the last row
    lastrow = WK.Worksheets("CMPFormatter").Range("A2").CurrentRegion.Rows.Count
    
    'Clean data
    WK.Worksheets("CMPFormatter").Range("A2:" & "L" & lastrow).Clear
    WK.Worksheets("CMPFormatter").Range("N2:" & "N" & lastrow).Clear
    
    'save the sheet
    WK.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

'Save file automaticly
Sub SaveAs_Example()
    
    Dim FilePath As String
    
    FilePath = Application.GetSaveAsFilename
    
    ActiveWorkbook.SaveAs Filename:=FilePath, FileFormat:=42
    
End Sub

'Remove manual sort
Public Sub ClearDataset()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = ThisWorkbook
    
    Dim i As Long
    'Take the last row from the CMPFormatter
    lastrow_sort = WK.Worksheets("CMPFormatter").Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 1 To lastrow_sort Step 1
        'Remove manual sort
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value = "P" Then
            WK.Worksheets("CMPFormatter").Range("B" & i).EntireRow.Delete
        End If
        
        'Remove captions lines, just with community, state, zip
        If WK.Worksheets("CMPFormatter").Range("C" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("D" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("E" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("F" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("J" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("L" & i).Value = vbNullString Then
                WK.Worksheets("CMPFormatter").Range("A" & i).EntireRow.Delete
        End If
        
        'Remove Cross References Empty lines
        If WK.Worksheets("CMPFormatter").Range("C" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("D" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("E" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("F" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("G" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("H" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("I" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("J" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i).Value = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("L" & i).Value = vbNullString Then
                WK.Worksheets("CMPFormatter").Range("A" & i).EntireRow.Delete
        End If
    Next i
End Sub

'System Issue with CAINS
Sub FixCAIN()
    Dim WS As Worksheet, PR As Workbook, WK As Workbook, i As Long
    Set WK = ThisWorkbook
    
    lastrow = WK.Worksheets("CMPFormatter").Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To lastrow
        'If its a CAIN insert 1
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("K" & i).Value2 = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2 <> vbNullString Then
            
            WK.Worksheets("CMPFormatter").Range("B" & i).Value2 = 1
        End If

        'If the line before is empty add 1
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("K" & i).Value2 = vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2 = vbNullString Then
            
            WK.Worksheets("CMPFormatter").Range("B" & i).Value2 = WK.Worksheets("CMPFormatter").Range("B" & i - 1).Value2 + 1
        End If

        'If the next phone line is not empty and the current line is empty add 1
        If WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("K" & i + 1).Value2 <> vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i).Value2 = vbNullString Then
            
            WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 = WK.Worksheets("CMPFormatter").Range("B" & i).Value2 + 1
        End If

        'if the previos phone line is not empty, take the indent from the prevous line.
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("K" & i).Value2 <> vbNullString And _
            WK.Worksheets("CMPFormatter").Range("K" & i - 1).Value2 <> vbNullString Then
            
            WK.Worksheets("CMPFormatter").Range("B" & i).Value2 = WK.Worksheets("CMPFormatter").Range("B" & i - 1).Value2
        End If
    Next i
End Sub

'Identify indent Greater than 8
Sub FixIndent()

    Dim WS As Worksheet, PR As Workbook, WK As Workbook, i As Long
    Set WK = ThisWorkbook

    lastrow = WK.Worksheets("CMPFormatter").Range("A" & Rows.Count).End(xlUp).Row

    For i = 2 To lastrow
'        'Identify indent that start with a numero different of 1 : EX 0 4 5 6 0
'        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 = 0 And WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 > 1 Then
'            WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 = 1
'        End If
'        'Identify indent that are not fallowing the sequence : EX 0 1 2 3 5 0 - missing 4
'        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 And WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 > WK.Worksheets("CMPFormatter").Range("B" & i).Value2 + 1 Then
'            WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 = WK.Worksheets("CMPFormatter").Range("B" & i + 1).Value2 - 1
'        End If

        'Indent can't be greater than 8, system error.
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value > 8 Then
            MsgBox "Please verify indent greater than 8 in the cell : " & Replace(WK.Worksheets("CMPFormatter").Range("B" & i).Address, "$", "") & vbCr & "Click OK to continue..."
        End If
    Next i
End Sub

'Fix captions with differents types
Sub TypeBR()
    Dim WS As Worksheet, PR As Workbook, WK As Workbook, i As Long
    Set WK = ThisWorkbook
    
    lastrow = WK.Worksheets("CMPFormatter").Range("A" & Rows.Count).End(xlUp).Row
    
    For i = 2 To lastrow
        If WK.Worksheets("CMPFormatter").Range("B" & i).Value2 <> 0 Then
    
            WK.Worksheets("CMPFormatter").Range("A" & i).Value2 = WK.Worksheets("CMPFormatter").Range("A" & i - 1).Value2
        End If
    Next i
End Sub

