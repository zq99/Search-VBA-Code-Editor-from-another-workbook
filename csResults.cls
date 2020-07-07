Option Explicit

Private strSearchValue          As String
Private strModule               As String
Private strProcName             As String
Private strLineOfCode           As String
Private lngLineNumber           As Long
Private lngColumnNumber         As Long
Private lngProcStartLineNumber  As Long
Private lngProcNumberOfLines    As Long

Property Let Module(ByVal value As String)
    strModule = value
End Property

Property Get Module() As String
    Module = strModule
End Property

Property Let ProcName(ByVal value As String)
    strProcName = value
End Property

Property Get ProcName() As String
    ProcName = strProcName
End Property

Property Let LineOfCode(ByVal value As String)
    strLineOfCode = value
End Property

Property Get LineOfCode() As String
    LineOfCode = strLineOfCode
End Property

Property Let LineNo(ByVal value As Long)
    lngLineNumber = value
End Property

Property Get LineNo() As Long
    LineNo = lngLineNumber
End Property

Property Let ColumnNo(ByVal value As Long)
    lngColumnNumber = value
End Property

Property Get ColumnNo() As Long
    ColumnNo = lngColumnNumber
End Property

Property Let ProcStartLineNo(ByVal value As Long)
    lngProcStartLineNumber = value
End Property

Property Get ProcStartLineNo() As Long
    ProcStartLineNo = lngProcStartLineNumber
End Property

Property Let ProcNumberOfLines(ByVal value As Long)
    lngProcNumberOfLines = value
End Property

Property Get ProcNumberOfLines() As Long
    ProcNumberOfLines = lngProcNumberOfLines
End Property

Property Let SearchValue(ByVal value As String)
    strSearchValue = value
End Property

Property Get SearchValue() As String
    SearchValue = strSearchValue
End Property