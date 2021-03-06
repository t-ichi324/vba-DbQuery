VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------

Public ModeDebugLog As Boolean
Public ModeErrorLog As Boolean
Public LastError As String  '最後に発生したErrorを格納
Private mRs As DAO.Recordset
Private mFind As Boolean
Private mReaded As Boolean
Private mEOF As Boolean
Private mCount As Long


'■ログ保存フォルダ
Private Property Get LOG_DIR(): LOG_DIR = Application.CurrentProject.FullName: End Property
'//////////////////////////////////////
Private Sub Class_Initialize()
    ModeDebugLog = False    'True：SQLログを作成
    ModeErrorLog = True    'True：Errorログを作成
    mFind = False
End Sub

Private Sub Class_Terminate()
On Error GoTo ex
    mFind = False
    If Not mRs Is Nothing Then
        mRs.Close
        Set mRs = Nothing
    End If
ex:
End Sub
Private Sub ErrorProc(ByVal Key As String)
    Dim log As String
    LastError = """" & Key & """ " & Err.Description & vbCrLf
    log = "# " & Now & vbCrLf
    log = log & "[DbReader]ErrorProc ======" & vbCrLf
    log = log & LastError
    log = log & "========================="
    Debug.Print log
    If ModeErrorLog Then: AppendLog "error", log
End Sub
Private Sub AppendLog(ByVal fName As String, ByVal log As String)
On Error GoTo ex:
    Dim fso As Object
    Dim f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(LOG_DIR & "." & fName & ".log", 8, True)
    Call f.WriteLine(log & vbCrLf)
ex:
    If Not fso Is Nothing Then: Set fso = Nothing
    If Not f Is Nothing Then: f.Close: Set f = Nothing
End Sub

Private Function mGet(ByVal Key As String) As Variant
On Error GoTo ex:
    mGet = mRs.fields(Key)
Exit Function
ex: Call ErrorProc(Key)
End Function


Public Sub SetRecordset(ByRef rs As DAO.Recordset, Optional ByVal autoReadFirst As Boolean = False)
    Set mRs = rs
    mReaded = False
    mFind = False
    mCount = 0
    If Not mRs Is Nothing Then
        If Not mRs.EOF Then
            mFind = True
            mCount = mRs.RecordCount
        End If
    End If
    
    If autoReadFirst Then: Me.ReadLine
End Sub

Public Property Get IsFind() As Boolean
    IsFind = mFind
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Private Function EOF() As Boolean
    If Not IsFind Then
        EOF = True
    Else
        EOF = mRs.EOF
    End If
End Function

Public Function ReadLine() As Boolean
    If EOF Then
        ReadLine = False
    Else
        If mReaded Then
            mRs.MoveNext
        Else
            mReaded = True
        End If
        ReadLine = Not mRs.EOF
    End If
End Function

Public Function Sets(ByVal Key As String, ByRef ctrl As Control, Optional ByVal nullVal As Variant = "")
    Dim cn As String: cn = TypeName(ctrl)
    Dim v As String: v = GetStr(Key, nullVal)
    If cn = "Label" Then
        ctrl.text = v
    ElseIf cn = "TextBox" Then
        ctrl.Value = v
    ElseIf cn = "CheckBox" Then
        If Nz(v, "") = "" Then
            ctrl.Value = False
        Else
            ctrl.Value = CBool(v)
        End If
    Else
        ctrl.Value = v
    End If
    Set Sets = Me
End Function

Public Function GetVal(ByVal Key As String, Optional ByVal nullVal As Variant = Null) As Variant
    If EOF Then
        GetVal = nullVal
    Else
        GetVal = Nz(mGet(Key), nullVal)
    End If
End Function

Public Function GetStr(ByVal Key As String, Optional ByVal nullVal As String = "") As String
    If EOF Then
        GetStr = nullVal
    Else
        GetStr = ToString(mGet(Key), nullVal)
    End If
End Function

Public Function GetInt(ByVal Key As String, Optional ByVal nullVal As Integer = 0) As Integer
    If EOF Then
        GetInt = nullVal
    Else
        GetInt = ToInt(mGet(Key), nullVal)
    End If
End Function

Public Function GetLong(ByVal Key As String, Optional ByVal nullVal As Long = 0) As Long
    If EOF Then
        GetLong = nullVal
    Else
        GetLong = ToLong(mGet(Key), nullVal)
    End If
End Function

Public Function GetDouble(ByVal Key As String, Optional ByVal nullVal As Double = 0) As Double
    If EOF Then
        GetDouble = nullVal
    Else
        GetDouble = ToDouble(mGet(Key), nullVal)
    End If
End Function
Public Function GetBoole(ByVal Key As String, Optional ByVal nullVal As Boolean = False) As Boolean
    If EOF Then
        GetBoole = nullVal
    Else
        Dim v As String: v = StrConv(ToString(mGet(Key)), vbUpperCase)
        GetBoole = v = "TRUE"
    End If
End Function

'==============================================================================
'Variantを強制的に文字型へ変換する
'==============================================================================
Private Function ToString(ByVal v As Variant, Optional ByVal nullVal As String = "") As String
    If IsNull(v) Then
        ToString = nullVal
        Exit Function
    End If
    If IsEmpty(v) Then
        ToString = nullVal
        Exit Function
    End If
    If v = "" Then
        ToString = nullVal
        Exit Function
    End If

    ToString = CStr(v)
End Function

'==============================================================================
'Int型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Private Function ToInt(ByVal v As Variant, Optional ByVal nullVal As Integer = 0) As Integer
On Error GoTo ex:
    Dim str As String: str = ToString(v)
    If str = "" Then
        ToInt = nullVal
        Exit Function
    End If
    ToInt = CInt(str)
    Exit Function
ex:
    ToInt = nullVal
End Function

'==============================================================================
'Long型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Private Function ToLong(ByVal v As Variant, Optional ByVal nullVal As Long = 0) As Long
On Error GoTo ex:
    Dim str As String: str = ToString(v)
    If str = "" Then
        ToLong = nullVal
        Exit Function
    End If
    ToLong = CLng(str)
    Exit Function
ex:
    ToLong = nullVal
End Function

'==============================================================================
'Double型へ変換する、失敗した場合はnullValを返す
'==============================================================================
Private Function ToDouble(ByVal v As Variant, Optional ByVal nullVal As Double = 0) As Double
On Error GoTo ex:
    Dim str As String: str = ToString(v)
    If str = "" Then
        ToDouble = nullVal
        Exit Function
    End If
    ToDouble = CDbl(str)
    Exit Function
ex:
    ToDouble = nullVal
End Function
