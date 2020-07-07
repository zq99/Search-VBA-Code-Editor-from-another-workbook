Option Explicit

Public Sub TestVBASearch()

    Dim oSearch     As New csSearchVBA
    Dim coResults   As New Collection
    Dim oResult     As csResults
    Dim lRow        As Long
    
    oSearch.clearWords
    oSearch.addWord "Hello"
    oSearch.addWord "Goodbye"
    oSearch.fileName = "C:\TEST\DummyMacro.xlsm"
    Set coResults = oSearch.GetSearchResults
    If Not coResults Is Nothing Then
        lRow = 1
        For Each oResult In coResults
            With Sheet1
                .Cells(lRow, 1).value = oSearch.fileName
                .Cells(lRow, 2).value = oResult.Module
                .Cells(lRow, 3).value = oResult.ProcName
                .Cells(lRow, 4).value = oResult.LineOfCode
                .Cells(lRow, 5).value = oResult.LineNo
                .Cells(lRow, 6).value = oResult.ColumnNo
                .Cells(lRow, 7).value = oResult.ProcStartLineNo
                .Cells(lRow, 8).value = oResult.ProcNumberOfLines
                .Cells(lRow, 9).value = oResult.SearchValue
                lRow = lRow + 1
            End With
        Next
    End If
    
    Set oSearch = Nothing
    Set coResults = Nothing
    Set oResult = Nothing
End Sub