Option Explicit

Private marrSearchValues()
Private mintWordCount        As Integer
Private mstrFile             As String
Private mappAccess           As Access.Application
Private mappWord             As Word.Application
Private mwrdDoc              As Word.Document
Private mappPT               As PowerPoint.Application
Private mpptDoc              As PowerPoint.Presentation
Private mappExcel            As Excel.Application
Private mowkb                As Workbook

Private Const CLNG_ERROR_MSG_LOCKED_PROJECT As Long = 50289

Private Sub Class_Initialize()
    mintWordCount = 0
End Sub

Property Let fileName(ByVal value As String)
    mstrFile = value
End Property

Property Get fileName() As String
    fileName = mstrFile
End Property

Public Sub clearWords()
    mintWordCount = 0
    ReDim arrSearchValues(0 To mintWordCount)
End Sub

Public Sub addWord(ByVal value As String)
    ReDim Preserve marrSearchValues(0 To mintWordCount)
    marrSearchValues(mintWordCount) = value
    mintWordCount = mintWordCount + 1
End Sub

Public Function GetSearchResults() As Collection

On Error GoTo ERR_HANDLER:

    Dim objFSO          As New FileSystemObject
    Dim objFile         As File
    Dim blnCheckFile    As Boolean
    
    Set GetSearchResults = Nothing
    If Len(Dir(mstrFile)) > 0 Then
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.GetFile(mstrFile)
        Application.StatusBar = "Checking File: " & objFile.Path & " : {" & Now() & "}"
        Set GetSearchResults = FindSearchValues(objFile)
    End If
    
EXIT_HERE:

    Application.StatusBar = Empty
    Set objFSO = Nothing
    Set objFile = Nothing
    Exit Function
    
ERR_HANDLER:

    Debug.Print Err.Description
    GoTo EXIT_HERE
    
End Function

Private Function FindSearchValues(ByVal oFile As File) As Collection
    
On Error GoTo ERR_HANDLER:

    Dim VBComp          As VBIDE.VBComponent
    Dim VBCodeMod       As VBIDE.CodeModule
    Dim VBComponentList As VBIDE.VBComponents
    Dim ProcKind        As VBIDE.vbext_ProcKind
    Dim strProcName     As String
    Dim strModuleName   As String
    Dim lngLineNum      As Long
    Dim lngNumLines     As Long
    Dim lngStartLine    As Long
    Dim lngEndLine      As Long
    Dim lngStartColumn  As Long
    Dim lngEndColumn    As Long
    Dim blnFound        As Boolean
    Dim collResults     As New Collection
    Dim oResult         As csResults
    Dim ltxtCnt         As Long

    Set FindSearchValues = Nothing
    
    Set VBComponentList = getVBAComponentList()
    If VBComponentList Is Nothing Then
        GoTo EXIT_HERE
    End If
    
    Application.VBE.MainWindow.Visible = False
    
    For Each VBComp In VBComponentList
            DoEvents
            Select Case VBComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document

                Set VBCodeMod = VBComp.CodeModule
                strModuleName = VBCodeMod.Name
                With VBCodeMod
                
                    For ltxtCnt = LBound(marrSearchValues) To UBound(marrSearchValues)
                        lngStartLine = 1
                        lngEndLine = .CountOfLines
                        lngStartColumn = 1
                        lngEndColumn = 255
                    
                        blnFound = .Find(Target:=marrSearchValues(ltxtCnt), StartLine:=lngStartLine, StartColumn:=lngStartColumn, _
                                         EndLine:=lngEndLine, EndColumn:=lngEndColumn, wholeword:=True, MatchCase:=False, patternsearch:=False)
                        Do Until Not blnFound
                            DoEvents
                            strProcName = .ProcOfLine(lngStartLine, ProcKind)
            
                            Set oResult = New csResults
                            oResult.SearchValue = marrSearchValues(ltxtCnt)
                            oResult.ProcName = strProcName
                            oResult.Module = strModuleName
                            oResult.LineOfCode = Trim(VBComp.CodeModule.Lines(lngStartLine, 1))
                            oResult.LineNo = lngStartLine
                            oResult.ColumnNo = lngStartColumn
                            oResult.ProcStartLineNo = .ProcStartLine(strProcName, ProcKind)
                            oResult.ProcNumberOfLines = .ProcCountLines(strProcName, ProcKind) + .ProcStartLine(strProcName, ProcKind)
                            collResults.Add oResult
                                      
                            lngEndLine = .CountOfLines
                            lngStartColumn = lngEndColumn + 1
                            lngEndColumn = 255
                            blnFound = .Find(Target:=marrSearchValues(ltxtCnt), StartLine:=lngStartLine, StartColumn:=lngStartColumn, _
                                             EndLine:=lngEndLine, EndColumn:=lngEndColumn, wholeword:=True, MatchCase:=False, patternsearch:=False)
                        Loop
                                    
                    Next ltxtCnt
                End With
            End Select
    Next VBComp
    Application.VBE.MainWindow.Visible = True
    
    Set FindSearchValues = collResults

EXIT_HERE:

    Call CloseAll
    Application.VBE.MainWindow.Visible = True
    Set VBComp = Nothing
    Set VBCodeMod = Nothing
    Set VBComponentList = Nothing
    Set oResult = Nothing
    Set collResults = Nothing
    Exit Function
    
ERR_HANDLER:

    Select Case Err.Number
    Case CLNG_ERROR_MSG_LOCKED_PROJECT
        Debug.Print "Locked VBA code"
    Case Else
        Debug.Print Err.Description
    End Select
    GoTo EXIT_HERE

End Function

Private Function getVBAComponentList() As VBIDE.VBComponents
    Dim strFileExt As String
    
    Set getVBAComponentList = Nothing
    
    If Len(mstrFile) > 0 Then
    
        strFileExt = GetExtension(mstrFile)
        
        Select Case strFileExt
        Case "xls", "xlt", "xla", "xlb", "xlb", "xlc", "xld", "xlk", "xll", "xlm", "xlv", "xlw", "xlsm", "xlsb", "xltm", "xlam"
            Set mappExcel = CreateObject("Excel.Application")
            Set mowkb = mappExcel.Workbooks.Open(fileName:=mstrFile, UpdateLinks:=False, ReadOnly:=True, local:=True)
            Set getVBAComponentList = mowkb.VBProject.VBComponents
        
        Case "mdb", "accdb", "accdt", "mdn"
            Set mappAccess = CreateObject("Access.Application")
            mappAccess.Visible = False
            mappAccess.OpenCurrentDatabase mstrFile, False
            Set getVBAComponentList = mappAccess.VBE.ActiveVBProject.VBComponents
            
        Case "doc", "docm"
            Set mappWord = CreateObject("Word.Application")
            mappWord.Visible = False
            Set mwrdDoc = mappWord.Documents.Open(fileName:=mstrFile, ReadOnly:=True)
            Set getVBAComponentList = mappWord.VBE.ActiveVBProject.VBComponents
            
        Case "ppt"
            Set mappPT = CreateObject("PowerPoint.Application")
            mappPT.Visible = False
            Set mpptDoc = mappPT.Presentations.Open(fileName:=mstrFile, ReadOnly:=msoTrue)
            Set getVBAComponentList = mappPT.VBE.ActiveVBProject.VBComponents
        End Select
    End If
    
End Function

Private Sub CloseAll()
    If Not mowkb Is Nothing Then
        mowkb.Close False
    End If
    Set mowkb = Nothing

    If Not mappExcel Is Nothing Then
        mappExcel.Quit
    End If
    Set mappExcel = Nothing

    If Not mappAccess Is Nothing Then
        mappAccess.CloseCurrentDatabase
        mappAccess.Quit
    End If
    Set mappAccess = Nothing
    
    If Not mappWord Is Nothing Then
        If Not mwrdDoc Is Nothing Then
            mwrdDoc.Close False
        End If
        Set mwrdDoc = Nothing
        mappWord.Quit
        Set mappWord = Nothing
    End If
    
    If Not mappPT Is Nothing Then
        mappPT.Quit
        Set mappPT = Nothing
    End If
End Sub

Private Function GetExtension(strFileName As String) As String
    Dim i       As Integer
    Dim strChar As String
    Dim intPos  As Integer
    
    For i = Len(Trim(strFileName)) To 1 Step -1
      strChar = Mid(Trim(strFileName), i, 1)
      If strChar = "." Then
        intPos = i + 1
        Exit For
      End If
    Next
    If intPos = 0 Then
        GetExtension = "#N/A"
    Else
        GetExtension = Mid(strFileName, intPos, Len(strFileName) - intPos + 1)
    End If
End Function