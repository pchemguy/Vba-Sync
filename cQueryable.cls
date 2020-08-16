VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cQueryable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Requires a refernce to the Microsoft ActiveX Data Objects 6.1 Library (or equivalent)

Private WithEvents mASyncConn As ADODB.Connection
Attribute mASyncConn.VB_VarHelpID = -1
Private mSyncConn As ADODB.Connection
Private mConn As ADODB.Connection
Private mComm As ADODB.Command
Private mSql As String
Private mProcedureAfterQuery As String
Private mAsync As Boolean
Private mConnectionString As String

Private Const mSyncExecute As Long = -1

Private Sub Class_Initialize()
    Set mComm = New ADODB.Command
    Set mConn = New ADODB.Connection
End Sub

Public Property Let Sql(value As String)
    mSql = value
End Property

Public Property Get Sql() As String
    Sql = mSql
End Property

Public Property Let ConnectionString(value As String)
    mConnectionString = value
End Property

Public Property Get ConnectionString() As String
    ConnectionString = mConnectionString
End Property

Public Property Let procedureAfterQuery(value As String)
    mProcedureAfterQuery = value
End Property

Public Property Get procedureAfterQuery() As String
    procedureAfterQuery = mProcedureAfterQuery
End Property

Public Sub createParam(pName As String, pType As DataTypeEnum, pValue As Variant, Optional pDirection As ParameterDirectionEnum = adParamInput, Optional pSize As Long = 0)
    Dim pm As ADODB.Parameter
    With mComm
       Set pm = .CreateParameter(name:=pName, Type:=pType, direction:=pDirection, value:=pValue, size:=pSize)
       .Parameters.Append pm
    End With
End Sub

Public Function SyncExecute()
    Set mSyncConn = mConn
    If connectionSuccessful Then
        With mComm
            .CommandText = mSql
            Set .ActiveConnection = mSyncConn
            Set SyncExecute = .execute(Options:=mSyncExecute)
        End With
    End If
End Function

Public Sub AsyncExecute()
    Set mASyncConn = mConn
    If connectionSuccessful Then
        With mComm
            .CommandText = mSql
            Set .ActiveConnection = mASyncConn
            .execute Options:=adAsyncExecute
        End With
    End If
End Sub

Private Function connectionSuccessful() As Boolean
    If mConn.State = adStateClosed Then
        mConn.ConnectionString = mConnectionString
    End If
    
    On Error GoTo errHandler
        If mConn.State = adStateClosed Then
            mConn.Open
        End If
    
        connectionSuccessful = (mConn.State = adStateOpen)
    On Error GoTo 0
    
    Exit Function
errHandler:
    Debug.Print "Error: Connection unsuccessful"
    connectionSuccessful = False
End Function

Private Sub mASyncConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    If mProcedureAfterQuery <> "" Then
        Call Application.Run(mProcedureAfterQuery, pRecordset)
    End If
End Sub
