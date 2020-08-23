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
Private mAsyncProcedure As String
Private mConnectionString As String

Private Const mSyncExecute As Long = -1
Private Const mConnectionUnsuccessfulError As Long = vbObjectError + 513
Private Const mEmptyAsyncProcedureError As Long = vbObject + 514

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

Public Property Let AsyncProcedure(value As String)
    mAsyncProcedure = value
End Property

Public Property Get AsyncProcedure() As String
    AsyncProcedure = mAsyncProcedure
End Property

Public Sub createParam(pName As String, pType As DataTypeEnum, pValue As Variant, Optional pSize As Long = 0, Optional pDirection As ParameterDirectionEnum = adParamInput)
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
    If Me.AsyncProcedure <> "" Then
        Set mASyncConn = mConn
        If connectionSuccessful Then
            With mComm
                .CommandText = mSql
                Set .ActiveConnection = mASyncConn
                .execute Options:=adAsyncExecute
            End With
        End If
    Else
        Call Err.Raise(Number:=mEmptyAsyncProcedureError, Description:="AsyncProcedure property not set")
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
    connectionSuccessful = False
    Call Err.Raise(Number:=mConnectionUnsuccessfulError, Description:="Database connection unsuccessful")
End Function

Private Sub mASyncConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    Call Application.Run(mAsyncProcedure, pRecordset)
End Sub
