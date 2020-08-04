Imports System.ComponentModel

Public MustInherit Class CTargetObjWithView農地台帳
    Inherits HimTools2012.TargetSystem.CTargetObjWithView

    Public Property DatabaseTableName As String = ""

    Public ReadOnly Property DataTable As DataTable
        Get
            If mvarRow IsNot Nothing Then
                Return mvarRow.Table
            Else
                Return Nothing
            End If
        End Get
    End Property

    Public Overrides Function SaveMyself() As Boolean
        Return MyBase.SaveBase(DatabaseTableName)
    End Function

    Public Sub New(ByRef pRow As DataRow, ByVal bAddNew As Boolean, ByRef pKey As HimTools2012.TargetSystem.DataKey, ByVal sDatabaseTableName As String)
        MyBase.New(pRow, bAddNew, pKey)

        DatabaseTableName = sDatabaseTableName
    End Sub

    <Browsable(False)>
    Public Overrides ReadOnly Property MyNamespace As String
        Get
            Return sLRDB
        End Get
    End Property

    Public Overrides Function InitDataViewNext(ByRef pDB As HimTools2012.TargetSystem.CDataViewCollection, Optional InterfaceName As String = "") As Boolean
        If Me.DataViewPage Is Nothing Then
            Me.DataViewPage = New DataViewNext農地台帳(Me, DatabaseTableName, IIf(InterfaceName = "", DatabaseTableName, InterfaceName), Me.DataTableWithUpdateList, pDB)
        End If
        Return True
    End Function
    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case sCommand
            Case "その他操作"
                Return Nothing
            Case "農地を呼ぶ"
                Return Nothing
            Case "変更を取り消す"
                If Me.Row IsNot Nothing Then
                    Me.CancelUpdate()
                End If
                Return Nothing

            Case Else
                Return MyBase.DoCommand(sCommand, sParams)
        End Select
    End Function


   
    Public Overrides Function MySystem() As HimTools2012.System管理.CSystemBase
        Return SysAD
    End Function


End Class
