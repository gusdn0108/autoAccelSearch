Sub SearchNaverMapAndUpdate()
    Dim cell As Range
    Dim searchQuery As String
    Dim naverMapUrl As String
    Dim userInput As String
    Dim rowNum As Long
    Dim chromePath As String
    
    On Error Resume Next
    Set cell = Selection
    On Error GoTo 0
    
    If cell Is Nothing Then
        MsgBox "셀을 선택하세요.", vbExclamation
        Exit Sub
    End If
    
    searchQuery = cell.Value
    rowNum = cell.Row

    searchQuery = Replace(searchQuery, " ", "+")
    
    naverMapUrl = "https://map.naver.com/v5/search/" & searchQuery

    chromePath = """C:\Program Files\Google\Chrome\Application\chrome.exe"""
    

    Shell (chromePath & " " & naverMapUrl), vbNormalFocus
    
    userInput = InputBox("T열 셀에 입력할 값을 입력하세요:")
    
    If userInput <> "" Then
        Cells(rowNum, "T").Value = userInput
    End If
End Sub

