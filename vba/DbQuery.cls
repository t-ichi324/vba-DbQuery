VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'-----------------------
' auth : t.ichi
'-----------------------

Public ModeDebug As Boolean
Public ModeDebugLog As Boolean
Public ModeErrorLog As Boolean
Public LastError As String  '最後に発生したErrorを格納
Public AllowNull As Boolean
Private mLastSql As String

Private mCon As Dao.Database
Private mWork As Dao.Workspace
Private mTable As String
Private mWhere As String
Private mOrders As String
Private mGroups As String
Private mInsFields As String
Private mInsVals As String
Private mUpdFields As String
Private mJoin As String
Private mClearFlg As Boolean

'■ログ保存フォルダ
Private Property Get LOG_DIR(): LOG_DIR = Application.CurrentProject.FullName: End Property
'//////////////////////////////////////
Private Sub Class_Initialize()
    ModeDebug = False      'True：コンソールへDebugLogにSQLを出力
    ModeDebugLog = False    'True：SQLログを作成
    ModeErrorLog = True    'True：Errorログを作成
    mClearFlg = True
    AllowNull = True
    Call DoClear
    Call Me.db(CurrentDb)
End Sub
Private Sub Class_Terminate()
On Error GoTo ex
    If ModeDebug Then Debug.Print "DbQuery-Terminate."
    Me.TransRollback
    If Not mCon Is Nothing Then
        mCon.Close
        Set mCon = Nothing
    End If
ex:
End Sub
Public Sub DoClose()
    Call Clear
    Call Class_Terminate
End Sub

Public Function Clear() As DbQuery
    mTable = ""
    mWhere = ""
    mOrders = ""
    mGroups = ""
    mInsFields = ""
    mInsVals = ""
    mUpdFields = ""
    mJoin = ""
    mClearFlg = True
    Set Clear = Me
End Function
Public Function NoClear() As DbQuery
    mClearFlg = False
    Set NoClear = Me
End Function

Private Sub DoClear()
    If mClearFlg Then
        Me.Clear
    End If
    mClearFlg = True
End Sub

Public Function Clone() As DbQuery
    Dim c As New DbQuery
    Call c.Extend(mTable, mWhere, mOrders, mGroups, mInsFields, mInsVals, mUpdFields, mJoin, mClearFlg)
    Set Clone = c
End Function
Public Sub Extend(pTable, pWhere, pOrders, pGroups, pInsFields, pInsVals, pUpdFields, pJoin, pClearFlg)
    mTable = pTable
    mWhere = pWhere
    mOrders = pOrders
    mGroups = pGroups
    mInsFields = pInsFields
    mInsVals = pInsVals
    mUpdFields = pUpdFields
    mJoin = pJoin
    mClearFlg = pClearFlg
End Sub

'//////////////////////////////////////
Public Property Get LastSql() As String: LastSql = mLastSql: End Property
Private Property Let LastSql(ByVal sql As String)
    mLastSql = sql
    If ModeDebug Or ModeDebugLog Then
        Dim log As String
        log = "# " & Now & vbCrLf
        log = log & "[DbQuery]DebugMode ======" & vbCrLf
        log = log & sql & vbCrLf
        log = log & "========================="
        If ModeDebug Then: Debug.Print log
        If ModeDebugLog Then: AppendLog "debug", log
    End If
End Property
Private Sub ErrorProc(ByVal sql As String)
    Dim log As String
    LastError = err.Description & vbCrLf & "-------------------------" & vbCrLf & sql & vbCrLf
    log = "# " & Now & vbCrLf
    log = log & "[DbQuery]ErrorProc ======" & vbCrLf
    log = log & LastError
    log = log & "========================="
    Debug.Print log
    If ModeErrorLog Then: AppendLog "error", log
    Clear
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

'//////////////////////////////////////
Public Function TransBegin()
    If mWork Is Nothing Then
        Set mWork = DBEngine.Workspaces(0)
        mWork.BeginTrans
    End If
End Function
Public Function TransCommit()
    If Not mWork Is Nothing Then
        mWork.CommitTrans
        Set mWork = Nothing
    End If
End Function
Public Function TransRollback()
    If Not mWork Is Nothing Then
        mWork.Rollback
        Set mWork = Nothing
    End If
End Function

'//////////////////////////////////////
'Connection
Public Function db(ByRef con As Dao.Database) As DbQuery
    Set mCon = CurrentDb
    Set db = Me
End Function
'Table
Public Function Table(tableName As String, Optional ByVal alias As String = "") As DbQuery
    mTable = tableName & IIf(alias = "", "", " AS " & alias)
    Set Table = Me
End Function
'TableJojn
Public Function LeftJoin(ByVal query As String, ByVal alias As String, ByVal comopreOn As String) As DbQuery
    mJoin = mJoin & vbCrLf & " LEFT JOIN " & query & IIf(alias = "", "", " AS " & alias) & " ON " & comopreOn
    Set LeftJoin = Me
End Function
Public Function RightJoin(ByVal query As String, ByVal alias As String, ByVal comopreOn As String) As DbQuery
    mJoin = mJoin & vbCrLf & " RIGHT JOIN " & query & IIf(alias = "", "", " AS " & alias) & " ON " & comopreOn
    Set RightJoin = Me
End Function

'Orders
Public Function OrderBy(fieldName As String) As DbQuery
    If fieldName = "" Then
        mOrders = ""
    Else
        mOrders = " ORDER BY " & fieldName
    End If
    Set OrderBy = Me
End Function
'Orders
Public Function GroupBy(fieldName As String) As DbQuery
    If fieldName = "" Then
        mGroups = ""
    Else
        mGroups = " GROUP BY " & fieldName
    End If
    Set GroupBy = Me
End Function

'Sets
Public Function Sets(fieldName As String, Val As Variant) As DbQuery
    Dim aVal As String
    If Not mInsFields = "" Then
        mInsFields = mInsFields & ","
        mInsVals = mInsVals & ","
        mUpdFields = mUpdFields & ","
    End If
    aVal = AntInj(Val)
    mInsFields = mInsFields & fieldName
    mInsVals = mInsVals & aVal
    mUpdFields = mUpdFields & fieldName & "=" & aVal
    Set Sets = Me
End Function
'Sets
Public Function SetsBlankToNull(fieldName As String, Val As Variant) As DbQuery
    Call Sets(fieldName, IIf(Nz(Val, "") = "", Null, Val))
    Set SetsBlankToNull = Me
End Function
Public Function SetsInt(fieldName As String, Val As Variant, Optional ByVal nullVal As Variant = Null)
    If Nz(Val, "") = "" Then
        Call Sets(fieldName, nullVal)
    Else
        'Call Sets(fieldName, CInt(val))
        Call Sets(fieldName, CLng(Val))
    End If
    Set SetsInt = Me
End Function
Public Function SetsLong(fieldName As String, Val As Variant, Optional ByVal nullVal As Variant = Null)
    If Nz(Val, "") = "" Then
        Call Sets(fieldName, nullVal)
    Else
        Call Sets(fieldName, CLng(Val))
    End If
    Set SetsLong = Me
End Function
Public Function SetsDouble(fieldName As String, Val As Variant, Optional ByVal nullVal As Variant = Null)
    If Nz(Val, "") = "" Then
        Call Sets(fieldName, nullVal)
    Else
        Call Sets(fieldName, CDbl(Val))
    End If
    Set SetsDouble = Me
End Function
'Sets
Public Function SetsRaw(fieldName As String, PAram As String) As DbQuery
    Dim aVal As String
    If Not mInsFields = "" Then
        mInsFields = mInsFields & ","
        mInsVals = mInsVals & ","
        mUpdFields = mUpdFields & ","
    End If
    mInsFields = mInsFields & fieldName
    mInsVals = mInsVals & PAram
    mUpdFields = mUpdFields & fieldName & "=" & PAram
    Set SetsRaw = Me
End Function

'//////////////////////////////////////

'Where
Public Function WhereRaw(ParamArray params()) As DbQuery
    Dim str As String
    Dim i As Long
    str = ""
    For i = LBound(params) To UBound(params)
        str = str & CStr(params(i))
    Next i
    mWhere = mWhere & str
    Set WhereRaw = Me
End Function

Public Function Wheres(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    mWhere = mWhere & MakeWheresRaw(fieldName, Val, operator)
    Set Wheres = Me
End Function
Public Function WheresNot(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    WheresNot = mWhere & " NOT " & MakeWheresRaw(fieldName, Val, operator)
    Set WheresNot = Me
End Function

Public Function Ands(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " AND "
    Set Ands = Me.Wheres(fieldName, Val, operator)
End Function
Public Function AndsNot(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " AND NOT "
    Set AndsNot = Me.Wheres(fieldName, Val, operator)
End Function

Public Function Ors(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " OR "
    Set Ors = Me.Wheres(fieldName, Val, operator)
End Function
Public Function OrsNot(fieldName As String, Val As Variant, Optional operator As String = "=") As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " OR NOT "
    Set OrsNot = Me.Wheres(fieldName, Val, operator)
End Function

Public Function AndsLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " AND "
    Set AndsLikes = Me.WheresLikes(fieldNames, vals, compare, pre_hit, aftr_hit)
End Function
Public Function AndsNotLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " AND NOT "
    Set AndsNotLikes = Me.WheresLikes(fieldNames, vals, compare, pre_hit, aftr_hit)
End Function

Public Function OrsLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " OR "
    Set OrsLikes = Me.WheresLikes(fieldNames, vals, compare, pre_hit, aftr_hit)
End Function

Public Function OrsNotLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " OR NOT "
    Set OrsNotLikes = Me.WheresLikes(fieldNames, vals, compare, pre_hit, aftr_hit)
End Function

Public Function WheresNotLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    If Not mWhere = "" Then: mWhere = mWhere & " NOT "
    Set WheresNotLikes = Me.WheresLikes(fieldNames, vals, compare, pre_hit, aftr_hit)
End Function

Public Function WheresLikes(fieldNames() As String, vals() As String, Optional compare As String = "AND", Optional pre_hit As Boolean = True, Optional aftr_hit As Boolean = True) As DbQuery
    Dim i, j As Integer
    Dim v, f As Variant
    i = 0
    mWhere = mWhere & "("
    For Each v In vals
        Dim vl As String
        vl = CStr(v)
        If Not vl = "" Then
            vl = IIf(pre_hit, "*", "") & CStr(vl) & IIf(aftr_hit, "*", "")
            If i > 0 Then: mWhere = mWhere & " " & compare & " "
            
            j = 0
            mWhere = mWhere & "("
            For Each f In fieldNames
                If j > 0 Then: mWhere = mWhere & " OR "
                Call Me.Wheres(CStr(f), vl, "LIKE")
                j = j + 1
            Next
            i = i + 1
            mWhere = mWhere & ")"
        End If
    Next
    If i = 0 Then
        mWhere = mWhere & "TRUE"
    End If
    mWhere = mWhere & ")"
    
    Set WheresLikes = Me
End Function

Private Function GetWhereSql() As String
    If mWhere = "" Then
        GetWhereSql = ""
    Else
        GetWhereSql = " WHERE " & mWhere
    End If
End Function
Private Function MakeWheresRaw(fieldName As String, Val As Variant, operator As String) As String
    Dim v As String: v = AntInj(Val)
    If v = "()" Then
        MakeWheresRaw = " FALSE "
        Exit Function
    End If
    If v = "NULL" Then
        If operator = "=" Then: operator = "IS"
        If operator = "<>" Then: operator = "IS NOT"
    End If
    MakeWheresRaw = fieldName & " " & operator & " " & v
End Function

'//////////////////////////////////////

Public Function ExequteRaw(ByVal sql As String) As Boolean
On Error GoTo ex
    ExequteRaw = False
    LastSql = sql
    Call mCon.Execute(LastSql)
    Call DoClear
    ExequteRaw = True
Exit Function
ex: Call ErrorProc(LastSql)
End Function
Public Function GetRecordsetRaw(ByVal sql As String) As Recordset
On Error GoTo ex
    Dim rs As Recordset
    LastSql = sql
    Set GetRecordsetRaw = mCon.OpenRecordset(LastSql)
    Call DoClear
Exit Function
ex: Call ErrorProc(LastSql)
End Function

Public Function Insert() As Boolean
    Insert = ExequteRaw(SqlInserts)
End Function
Public Function InsertIDENTITY() As Integer
    If Insert() Then
        InsertIDENTITY = Val(mCon.OpenRecordset("SELECT @@IDENTITY")(0))
    Else
        InsertIDENTITY = 0
    End If
End Function

Public Function Update() As Boolean
    Update = ExequteRaw(SqlUpdates)
End Function
Public Function Delete() As Boolean
    Delete = ExequteRaw(SqlDeletes)
End Function
Public Function SelectInsert(ByVal insertTable As String, Optional ByVal insertFields = "", Optional ByVal selectFields As String = "*") As Boolean
    SelectInsert = ExequteRaw(SqlInsertSelects(insertTable, insertFields, selectFields))
End Function
Public Function GetRecordset(Optional ByVal fieldName As String = "*") As Recordset
    Set GetRecordset = GetRecordsetRaw(SqlSelect(fieldName))
End Function
Public Function GetReader(Optional ByVal fieldName As String = "*") As DbReader
    Dim r As New DbReader
    Call r.SetRecordset(Me.GetRecordset(fieldName))
    Set GetReader = r
End Function
Public Function GetReaderOne(Optional ByVal fieldName As String = "*") As DbReader
    Dim r As New DbReader
    Call r.SetRecordset(Me.GetRecordset(fieldName), True)
    Set GetReaderOne = r
End Function

'Get String
Public Function GetField(ByVal fieldName As String, Optional ByVal nullVal As String = "") As String
On Error GoTo ex
    GetField = GetRsVal(GetRecordset(fieldName), "", nullVal)
Exit Function
ex: Call ErrorProc("GetField(""" & fieldName & """)")
End Function

'Get String
Public Function GetFields(ByVal fieldName As String, Optional ByVal nullVal As String = "") As String()
On Error GoTo ex
    Dim rs As Recordset
    Dim cnt As Long
    Dim preClear As Boolean
    preClear = mClearFlg
    cnt = Me.NoClear.GetCount()
    mClearFlg = preClear

    Dim r() As String
    If cnt > 0 Then
        ReDim r(cnt - 1)
        Set rs = Me.GetRecordset(fieldName)
        Dim i As Long
        i = 0
        Do While Not rs.EOF
            r(i) = Nz(rs.fields(0), nullVal)
            i = i + 1
            rs.MoveNext
        Loop
    Else
        DoClear
    End If
    GetFields = r
Exit Function
ex: Call ErrorProc("GetFields(""" & fieldName & """)")
End Function

'
Public Function GetCount() As Integer
On Error GoTo ex
    Dim tOrders As String: tOrders = mOrders
    mOrders = ""
    GetCount = CInt(GetRsVal(GetRecordset("COUNT(*) AS cnt"), "", "0"))
    mOrders = tOrders
Exit Function
ex: Call ErrorProc("GetCount()")
End Function

'
Public Function GetSum(ByVal fieldName As String) As Integer
On Error GoTo ex
    Dim tOrders As String: tOrders = mOrders
    mOrders = ""
    GetSum = CInt(GetRsVal(GetRecordset("SUM(" & fieldName & ") AS cnt"), "", "0"))
    mOrders = tOrders
Exit Function
ex: Call ErrorProc("GetCount()")
End Function

Public Function IsExists() As Boolean
    IsExists = Me.GetCount() > 0
End Function

Public Function GetMax(Optional ByVal fieldName As String = "*") As Long
On Error GoTo ex
    GetMax = CLng(GetRsVal(GetRecordset("MAX(" & fieldName & ") AS cnt"), "cnt", "0"))
Exit Function
ex: Call ErrorProc("GetMax()")
End Function

Public Function GetSql(Optional ByVal fieldName As String = "*") As String
    GetSql = SqlSelect(fieldName)
End Function

'//////////////////////////////////////
Private Function SqlSelect(Optional ByVal fields As String = "*") As String
    SqlSelect = "SELECT " & fields & " FROM " & mTable & " " & mJoin & " " & GetWhereSql & " " & mOrders & mGroups
End Function
Private Function SqlInserts() As String
    SqlInserts = "INSERT INTO " & mTable & "(" & mInsFields & ")VALUES(" & mInsVals & ")"
End Function
Private Function SqlUpdates() As String
    SqlUpdates = "UPDATE " & mTable & " " & mJoin & " SET " & mUpdFields & " " & GetWhereSql
End Function
Private Function SqlDeletes() As String
    SqlDeletes = "DELETE FROM " & mTable & " " & mJoin & " " & GetWhereSql
End Function
Private Function SqlInsertSelects(ByVal insertTable As String, Optional ByVal insertFields = "", Optional ByVal selectFields As String = "*") As String
    If insertFields = "" Then
        SqlInsertSelects = "INSERT INTO " & insertTable & " SELECT * FROM (" & SqlSelect(selectFields) & ") AS EXT"
    Else
        SqlInsertSelects = "INSERT INTO " & insertTable & "(" & insertFields & ")" & " SELECT * FROM (" & SqlSelect(selectFields) & ") AS EXT"
    End If
End Function

'//////////////////////////////////////

Private Function AntInj(ByVal v As Variant) As String
    Dim t As String: t = TypeName(v)
    Dim s As String
    
    If IsNull(v) Then
        s = "NULL"
        AntInj = s
        Exit Function
    End If
    
    If IsArray(v) Then
        Dim inParm As String
        Dim i As Integer
        inParm = ""
On Error GoTo inEx
        If UBound(v) >= 0 Then
            For i = LBound(v) To UBound(v)
                If Not inParm = "" Then inParm = inParm & ","
                inParm = inParm & AntInj(v(i))
            Next
            AntInj = "(" & inParm & ")"
            Exit Function
        End If
inEx:
        AntInj = "()"
        Exit Function
    End If
    
    Select Case t
        Case "Nothing"
            s = "NULL"
        Case "Date"
            s = "#" & CStr(v) & "#"
        Case "Boolean"
            If v = True Then
                s = "True"
            Else
                s = "False"
            End If
        Case "CheckBox"
            If CBool(v) = True Then
                s = "True"
            Else
                s = "False"
            End If
        Case "Integer"
            s = CStr(v)
        Case "Long"
            s = CStr(v)
        Case "Double"
            s = CStr(v)
        Case Else
            s = "'" & Replace(v, "'", "''") & "'"
    End Select
    AntInj = s
End Function

Private Function GetRsVal(ByRef rs As Recordset, fieldName As String, Optional ByVal nullVal As Variant) As Variant
    If rs Is Nothing Then
        GetRsVal = nullVal
        Exit Function
    End If
    If rs.EOF Then
        GetRsVal = nullVal
        Exit Function
    End If
    If fieldName = "" Then
        GetRsVal = Nz(rs.fields(0), nullVal)
    Else
        GetRsVal = Nz(rs.fields(fieldName), nullVal)
    End If
End Function
