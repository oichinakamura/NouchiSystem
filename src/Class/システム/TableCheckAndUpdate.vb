Imports HimTools2012.Data


Public MustInherit Class TableCheckAndUpdate
    Protected mvarTarget As DataTable
    Protected mvarLog As DataTable
    Protected mvarTableName As String
    Public MustOverride Function CheckStart() As CheckResult
    Protected mvarALTERTABLESQL As New System.Text.StringBuilder
    Private mvarDB As CommonDataBase

    Public Enum CheckResult
        CompleteUpdate = 1
        NoUpdate = 0
        UpdateFailed = -1
    End Enum

    Public Sub New(ByRef pDB As CommonDataBase, ByVal sTableName As String, ByRef pTarget As DataTable, ByRef pUpdateLog As HimTools2012.Data.CUpdateLog)
        mvarTarget = pTarget
        mvarLog = pUpdateLog
        mvarTableName = sTableName
        mvarDB = pDB

        If mvarLog.PrimaryKey Is Nothing OrElse mvarLog.PrimaryKey.Length = 0 Then
            Try
                mvarLog.PrimaryKey = New DataColumn() {mvarLog.Columns("Key")}
            Catch ex As Exception
                MsgBox("更新ログを正しく取得できませんでした。")
            End Try
        End If
    End Sub

    Protected Overloads Sub MakeTBLFieldModifySQL(ByVal sColumnName As String, ByVal pType As System.Data.OleDb.OleDbType)
        If mvarTarget.Columns.Contains(sColumnName) Then
            Return
        Else
            Select Case pType.ToString
                Case OleDb.OleDbType.Integer
                    mvarALTERTABLESQL.AppendLine(String.Format("ALTER TABLE [{0}] ADD [{1}] LONG;", mvarTableName, sColumnName))
                Case Else
                    Stop
            End Select
        End If
        Return
    End Sub

    Protected Sub MakeTBLFieldDropSQL(ByVal sColumnName As String)
        If Not mvarTarget.Columns.Contains(sColumnName) Then
            Return
        Else
            mvarALTERTABLESQL.AppendLine(String.Format("ALTER TABLE [{0}] DROP [{1}]", mvarTableName, sColumnName))
        End If
        Return
    End Sub


    Protected Overloads Function MakeTBLFieldModifySQL(ByVal sColumnName As String, ByVal sType As String) As Boolean
        If mvarTarget.Columns.Contains(sColumnName) Then
            Return False
        Else
            mvarALTERTABLESQL.AppendLine(String.Format("ALTER TABLE [{0}] ADD [{1}] {2};", mvarTableName, sColumnName, sType))
            Return True
        End If
    End Function


    Protected Function CheckLog(ByVal sKey As String) As Boolean
        Dim sKey2 As String = New StackFrame(1).GetMethod.Name
        Dim bResult As Boolean = (mvarLog.Rows.Find(MyClass.GetType.FullName & "_" & sKey) IsNot Nothing)
        Return bResult
    End Function

    Protected Function AddLog(ByVal sKey As String, bBeforeCheck As CheckResult) As CheckResult
        Dim sRes As String = ""
        If mvarALTERTABLESQL.Length > 0 Then
            Dim Va As String() = Split(mvarALTERTABLESQL.ToString, vbCrLf)
            For Each SX As String In Va
                If Len(SX) Then

                    sRes = mvarDB.ExecuteSQL(SX)
                    If sRes = "OK" OrElse sRes = "" Then
                    Else
                        'Stop
                    End If

                End If

            Next

            mvarALTERTABLESQL.Clear()
        End If
        If sRes.Length = 0 OrElse sRes.IndexOf("OK") > -1 Then
            Dim sWriteKey As String = String.Format("{0}_{1}", MyClass.GetType.FullName, sKey)
            Dim sRes2 As String = mvarDB.ExecuteSQL(String.Format("INSERT INTO [S_システム更新履歴]([KEY],[Update]) VALUES('{0}_{1}',Now);", MyClass.GetType.FullName, sKey))
            mvarLog.Rows.Add({sWriteKey, Now()})
            If sRes2 = "" OrElse sRes2.IndexOf("OK") > -1 Then
                Return CheckResult.CompleteUpdate
            Else
                Return CheckResult.UpdateFailed
            End If
        Else
            Return CheckResult.UpdateFailed
        End If
    End Function


End Class



