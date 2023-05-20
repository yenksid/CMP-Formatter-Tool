Attribute VB_Name = "CMPS"
'Convert the CMP format to a more redable format
Public Sub CMP()
    Dim WK As Workbook, arr As Variant, class_of_service, phone, section As String
    Set WK = ThisWorkbook
    
    Dim caption_header As Boolean
    Dim i As Long
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Asigning the DATASET to an array
    arr = WK.Worksheets("CMP").Range("A1").CurrentRegion
    
    'From the first item in the array to the last
    For i = LBound(arr, 1) To UBound(arr, 1)
        
        On Error Resume Next
        'Captions Header
        'NEWSTART It's not flagging caption headers, DLS does.
        If Mid(arr(i, 1), 55, 1) = 0 And Mid(arr(i + 1, 1), 55, 1) <> 0 Then
            caption_header = True
            'To know if its a caption we need to review the line below each line
            'At the end of the file the next line is empty, so we need to identify
            'this line so that the macro can work properly
            If Mid(arr(i + 1, 1), 55, 1) = vbnllstring Then
                caption_header = False
            End If
        Else
            caption_header = False
        End If
        'Finish the program
        On Error GoTo 0
        
        'Indent Order
'        If indent_level = 0 And caption_header = False Then
'            WK.Worksheets("DATASET").Range("C" & i).Value = 0
'
'        ElseIf caption_header = True Then
'            WK.Worksheets("DATASET").Range("C" & i).Value = 10
'
'        ElseIf indent_level <> 0 And caption_header = False And indent_level <> "" Then
'            WK.Worksheets("DATASET").Range("C" & i).Value = WK.Worksheets("DATASET").Range("C" & i - 1).Value + 10
'        End If
            
        'Class Of Service
        'captions headers do not have the type so we take the next type and added it to the caption header.
        If Mid(arr(i, 1), 249, 1) <> " " Then
            class_of_service = Mid(arr(i, 1), 249, 1)
            WK.Worksheets("CMP").Range("A" & i).Value = class_of_service
        Else
            'Take next type
            class_of_service = Mid(arr(i + 1, 1), 249, 1)
            WK.Worksheets("CMP").Range("A" & i).Value = class_of_service
        End If
        
        'Indent level
        indent_level = Mid(arr(i, 1), 55, 1)
        WK.Worksheets("CMP").Range("B" & i).Value = indent_level
        
        '//If Caption Header Is Falso Do Things Below//
        '//Missing toll free//
        '//3 digits number//
        '//Telco code, EAS, CLEC, LOCAL//
        
        'Street number
        street_number = Trim(Mid(arr(i, 1), 260, 32))
        WK.Worksheets("CMP").Range("D" & i).Value = Trim(street_number)
        
        'Cardinals
        'DLS & NEWSTART have it in different positions
        cardinales = Trim(Mid(arr(i, 1), 362, 15))
        WK.Worksheets("CMP").Range("F" & i).Value = Trim(cardinales)
        
        'Street name
        street_name = Trim(Mid(arr(i, 1), 292, 70))
        WK.Worksheets("CMP").Range("E" & i).Value = Trim(street_name)
        
        'Comunity
        community = Trim(Mid(arr(i, 1), 377, 45))
        WK.Worksheets("CMP").Range("G" & i).Value = Trim(community)
        
        'State
        state_code = Trim(Mid(arr(i, 1), 422, 18))
        WK.Worksheets("CMP").Range("H" & i).Value = Trim(state_code)
        
        'Postal Code
        postal_code = Trim(Mid(arr(i, 1), 440, 13))
        WK.Worksheets("CMP").Range("I" & i).Value = Trim(postal_code)
        
        
        'Right Aligned Text
        
        
        'Phone
        phone = Mid(arr(i, 1), 453, 20)
        If Len(phone) > 5 Then
            WK.Worksheets("CMP").Range("L" & i).Value = Trim(Replace(phone, " ", ""))
        Else
            WK.Worksheets("CMP").Range("J" & i).Value = Trim(Replace(phone, " ", ""))
        End If
        
        'Name
        caption_name = Mid(arr(i, 1), 513, 100)
        
        '//Verify if we have differents class service or listings types//
        'OLD CODE, waiting to see the below version working before remove it.
'        'Residential listings
'        If indent_level = 0 And class_of_service = "R" And caption_header = False Then
'
'            'Last name
'            last_name = Left(caption_name, InStr(1, caption_name, "|") - 1)
'            WK.Worksheets("DATASET").Range("D" & i).Value = Trim(last_name)
'
'            'First name
'            first_name = Right(caption_name, Len(caption_name) - Len(last_name) - 2)
'            WK.Worksheets("DATASET").Range("E" & i).Value = Trim(first_name)
'            'Business listings
'        ElseIf indent_level = 0 And class_of_service = "B" And caption_header = False Then
'            'Manage cross reference listings
'            '//The macro is leaving an empty line were it find a cross refence//
'            If Left(caption_name, 3) = "See" Then
'                WK.Worksheets("DATASET").Range("H" & i - 1).Value = Trim(caption_name)
'            Else
'                caption_name = Trim(Replace(caption_name, "|", ""))
'                WK.Worksheets("DATASET").Range("D" & i).Value = Trim(caption_name)
'            End If
'            'Caption header business
'        ElseIf caption_header = True And class_of_service = "B" Then
'            WK.Worksheets("DATASET").Range("D" & i).Value = Trim(caption_name)
'            'Caption header residentials
'        ElseIf caption_header = True And class_of_service = "R" Then
'
'            'Last name
'            last_name = Left(caption_name, InStr(1, caption_name, " ") - 1)
'            WK.Worksheets("DATASET").Range("D" & i).Value = Trim(last_name)
'
'            'First name
'            first_name = Right(caption_name, Len(caption_name) - Len(last_name) - 1)
'            WK.Worksheets("DATASET").Range("E" & i).Value = Trim(first_name)
'            'Caption display text
'        ElseIf indent_level <> 0 Then
'             WK.Worksheets("DATASET").Range("D" & i).Value = Trim(caption_name)
'        End If

        'manage names
        If indent_level = 0 And caption_header = False Then
            'Manage cross reference listings
            '//The macro is leaving an empty line were it find a cross refence//
            If Left(caption_name, 4) = "See " Then
                WK.Worksheets("CMP").Range("J" & i - 1).Value = Trim(caption_name)
            Else
                caption_name = Trim(Replace(caption_name, "|", ""))
                WK.Worksheets("CMP").Range("C" & i).Value = Trim(caption_name)
            End If
            'Caption header business
        ElseIf caption_header = True Then
            WK.Worksheets("CMP").Range("C" & i).Value = Trim(Replace(Replace(caption_name, "|", ""), ",", ""))
        ElseIf indent_level <> 0 Then
             WK.Worksheets("CMP").Range("C" & i).Value = Trim(caption_name)
        End If
        
        'merge names
        'move extra lines to a new column
        'separated the comunity, state, zip
        'Remove the missing line leave by the cross references
        'Remove dots excep from the web address
        
    Next i
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "CMP Conversion Completed, Thanks!"
    
    'Insert row to provide the header
    'WK.Worksheets("DATASET").Rows(1).Insert
    'WK.Worksheets("DATASET").Range("A1:K1").Value = header_name
End Sub


'Clean all the data in the sheet
Sub CleanSheetCMP()
    Dim WS As Worksheet
    Dim lastrow As Long
    Set WK = ThisWorkbook
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    'Get the last row
    lastrow = WK.Worksheets("CMP").Range("A1").CurrentRegion.Rows.Count
    
    'Clean data
    WK.Worksheets("CMP").Range("A1:" & "L" & lastrow).Clear
    
    'save the sheet
    WK.Save
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub
