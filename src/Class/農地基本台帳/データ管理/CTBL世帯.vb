Imports HimTools2012.CommonFunc

Public Class CTBL世帯
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(DSet As DataSet, pTable As DataTable)
        MyBase.New(pTable, sLRDB)
        

        'Dim mvarTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect(sSQL)
        'Dim pCheck As New 世帯InfoUpdate("D:世帯Info", mvarTBL)

        'pCheck.CheckStart()

        SetTableEnv(DSet, "")
    End Sub

    Public Sub SetTableEnv(ByRef DSet As DataSet, ByVal StUp As String)
        Try
            DSet.Relations.Add(StUp & "世帯主", DSet.Tables("D:個人Info").Columns("ID"), Columns("世帯主ID"), False)
            Columns.Add(New DataColumn("世帯主氏名", GetType(String), "Parent(" & StUp & "世帯主).氏名"))
            Columns.Add(New DataColumn("フリガナ", GetType(String), "Parent(" & StUp & "世帯主).フリガナ"))
            Columns.Add(New DataColumn("世帯主行政区ID", GetType(String), "Parent(" & StUp & "世帯主).行政区ID"))
            Columns.Add(New DataColumn("世帯主行政区名", GetType(String), "Parent(" & StUp & "世帯主).行政区名"))
            Columns.Add(New DataColumn("世帯主郵便番号", GetType(String), "Parent(" & StUp & "世帯主).郵便番号"))
            Columns.Add(New DataColumn("世帯主電話番号", GetType(String), "Parent(" & StUp & "世帯主).電話番号"))
            Columns.Add(New DataColumn("検索フリガナ", GetType(String), "Parent(" & StUp & "世帯主).検索フリガナ"))
            Me.Columns.Add(New DataColumn("住所", GetType(String), "Parent(世帯主).住所"))
            'Columns.Add(New DataColumn("あっせん種別名", GetType(String), "IIF([あっせん希望種別]=1,'農業委員会',IIF([あっせん希望種別]=2,'中間管理機構',IIF([あっせん希望種別]=3,'農委・機構','なし')))"))

            DataInitAfter(DSet)
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Public Function Get世帯主IDList() As List(Of Decimal)
        Dim pList As New List(Of Decimal)

        For Each pRow As DataRow In Me.Rows
            If Not IsDBNull(pRow.Item("世帯主ID")) AndAlso pRow.Item("世帯主ID") <> 0 Then
                Dim nID As Decimal = pRow.Item("世帯主ID")
                Dim pRowK As DataRow = App農地基本台帳.TBL個人.Rows.Find(nID)
                If pRowK Is Nothing AndAlso Not pList.Contains(nID) Then
                    pList.Add(nID)
                End If
            End If
        Next

        Return pList
    End Function

    Public Overrides Sub MergePlus(ByVal pTable As DataTable, Optional ByVal preserveChanges As Boolean = False, Optional ByVal pAction As System.Data.MissingSchemaAction = MissingSchemaAction.Add)
        SyncLock Me
            MyBase.MergePlus(pTable, preserveChanges, pAction)
            '世帯主の抽出
            Dim sX As List(Of Decimal) = Me.Get世帯主IDList()

            Dim sB As New System.Text.StringBuilder
            For Each nID As Decimal In sX
                sB.Append("," & nID)

                If sB.Length > 128 Then
                    Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & Mid$(sB.ToString, 2) & ")")
                    App農地基本台帳.TBL個人.MergePlus(pAddK, True)
                    sB.Clear()
                End If
            Next
            If sB.Length > 0 Then
                Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & Mid$(sB.ToString, 2) & ")")
                App農地基本台帳.TBL個人.MergePlus(pAddK, True)
                sB.Clear()
            End If
        End SyncLock
    End Sub


    Public Function Update(ByVal mvarUpdateRow As HimTools2012.Data.UpdateRow, bAddNew As Boolean) As Boolean
        If bAddNew Then
            Stop
            Return True
        Else
            Return SysAD.DB(sLRDB).UpdateRecord(Me.Body, mvarUpdateRow)
        End If
    End Function

    Public Function GetObject(sKey As String) As HimTools2012.TargetSystem.CTargetObjWithView
        Select Case GetKeyHead(sKey)
            Case Else
                Dim pRow As DataRow = Me.Rows.Find(GetKeyCode(sKey))

                If pRow Is Nothing Then
                    Return Nothing
                Else
                    Return New CObj個人(pRow, False)
                End If
        End Select
        Return Nothing
    End Function



    Public Overrides Function GetDataView(ByVal sWhere As String, ByVal sFilter As String, ByVal sOrderBy As String) As System.Data.DataView
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:世帯Info] WHERE " & sWhere)
        Me.MergePlus(pTBL)
        Return New DataView(Me.Body, sFilter, sOrderBy, DataViewRowState.CurrentRows)
    End Function
End Class

Public Class 世帯InfoUpdate
    Inherits TableCheckAndUpdate

    Public Sub New(ByVal sTableName As String, ByRef pTarget As DataTable)
        MyBase.New(SysAD.DB(sLRDB), sTableName, pTarget, SysAD.DB(sLRDB).UpdateLog)
    End Sub

    Public Overrides Function CheckStart() As TableCheckAndUpdate.CheckResult
        If Check個人Info申告納税方式20150129(CheckResult.CompleteUpdate) = CheckResult.NoUpdate Then
        ElseIf Check世帯Info基本20150320(CheckResult.CompleteUpdate) = CheckResult.NoUpdate Then
        End If

        Return CheckResult.NoUpdate
    End Function

    Public Function Check世帯Info基本20150320(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            'MakeTBLFieldDropSQL("認定農業者")
            'MakeTBLFieldModifySQL("農年加入受給種別", "LONG")
            MakeTBLFieldModifySQL("あっせん希望種別", "LONG")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    ''' <summary>
    ''' 2014/6農地法改正による管理項目の追加
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check個人Info申告納税方式20150129(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            'MakeTBLFieldModifySQL("青色申告", "LONG")
            'MakeTBLFieldModifySQL("青色申告開始年", "LONG")

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

End Class