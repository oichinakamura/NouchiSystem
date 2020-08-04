
Imports System.ComponentModel

Public Class CTBL個人
    Inherits HimTools2012.Data.DataTableWith

    Private BackPanel As HimTools2012.TabPages.BackGroundPage

    Public Sub New(DSet As DataSet, pTable As DataTable)
        MyBase.New(pTable, sLRDB)

        If Not SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID]=0").Columns.Contains("システム確認") Then
            Try
                Dim sResult As String = SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [D:個人Info] ADD システム確認 DATETIME")
                If sResult.Length > 0 Then
                    Stop
                Else
                    Me.MergePlus(SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [ID]=0"))
                End If
            Catch ex As Exception

            End Try
        Else
            If Not SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_削除個人] WHERE [ID]=0").Columns.Contains("システム確認") Then
                Dim sResult As String = SysAD.DB(sLRDB).ExecuteSQL("ALTER TABLE [D_削除個人] ADD システム確認 DATETIME")
            End If

        End If
        DataInitAfter(DSet)
    End Sub

    Public Overrides Sub MergePlus(pTable As System.Data.DataTable, Optional preserveChanges As Boolean = False, Optional pAction As System.Data.MissingSchemaAction = System.Data.MissingSchemaAction.Add)
        SyncLock Me
            If pTable IsNot Nothing Then
                MyBase.MergePlus(pTable)
            End If
        End SyncLock
    End Sub

    Public Sub 検索文字初期化()
        BackPanel = New HimTools2012.TabPages.BackGroundPage(False, True, "検索文字のバックグラウンド処理", "検索文字のバックグラウンド処理", True, "バックグラウンドで検索文字を初期化しています。操作が終了すれば自動的に閉じます。")
        SysAD.MainForm.MainTabCtrl.AddPage(BackPanel)
        BackPanel.Start(AddressOf bgWorker_DoWork)
    End Sub

    Public Overrides Function GetDataView(ByVal sWhere As String, ByVal sFilter As String, ByVal sOrderBy As String) As System.Data.DataView
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE " & sWhere)
        Me.MergePlus(pTBL)
        Return New DataView(Me.Body, sFilter, sOrderBy, DataViewRowState.CurrentRows)
    End Function

    Private Sub bgWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        BackPanel.MaxValue = 100
        BackPanel.Value = 0

        If Me.Body.Columns.Contains("システム確認") Then
            Dim pDate As DateTime = CDate(SysAD.DB(sLRDB).DBProperty("検索CHKDate", Now.Date.ToShortDateString))
            SysAD.DB(sLRDB).DBProperty("検索CHKDate") = Now.Date.ToString()

            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE [システム確認] Is Null Or [システム確認]<=#{1}/{2}/{0}#", pDate.Year, pDate.Month, pDate.Day)

            pTBL.Columns.Add("検索CHK", GetType(String))
            For Each pRow As DataRow In pTBL.Rows
                If IsDBNull(pRow.Item("検索フリガナ")) Then
                    pRow.Item("検索フリガナ") = ""
                End If
                pRow.Item("検索CHK") = Replace(pRow.Item("フリガナ").ToString, " ", "")
            Next

            Dim pView As New DataView(pTBL, "not [検索CHK]=[検索フリガナ]", "", DataViewRowState.CurrentRows)
            BackPanel.MaxValue = pView.Count
            BackPanel.Value = 0


            For Each pRow As DataRowView In pView
                If Not IsDBNull(pRow.Item("フリガナ")) AndAlso Trim(pRow.Item("フリガナ").ToString).Length > 0 Then
                    If Not pRow.Item("検索フリガナ").ToString = Replace(pRow.Item("フリガナ").ToString, " ", "") Then
                        pRow.Item("検索フリガナ") = Replace(pRow.Item("フリガナ").ToString, " ", "")
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [検索フリガナ]='{0}',[システム確認]=Now WHERE [ID]={1}", Replace(pRow.Item("フリガナ").ToString, " ", ""), pRow.Item("ID"))
                    Else
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:個人Info] SET [システム確認]=Now WHERE [ID]={1}", Replace(pRow.Item("フリガナ").ToString, " ", ""), pRow.Item("ID"))
                    End If
                End If
                If BackPanel._Cancel Then
                    Exit For
                End If
                BackPanel.IncrementProgress()
            Next
        End If
        e.Result = "Compleate"
    End Sub
End Class

Public Class 検索文字作成
    Inherits HimTools2012.clsAccessor
    Private mvarTable As DataTable

    Public Sub New(pTBL As DataTable)
        MyBase.new()
        mvarTable = pTBL
    End Sub

    Public Overrides Sub Execute()
        Message = "検索文字作成.."
        Value = 0
        Maximum = mvarTable.Rows.Count
        Dim sSQL As New System.Text.StringBuilder

        For Each pRow As DataRow In mvarTable.Rows
            sSQL.AppendLine("UPDATE [D:個人Info] SET [検索フリガナ]='" & Get検索文字(pRow.Item("フリガナ")) & "' WHERE [ID]=" & pRow.Item("ID"))

            If sSQL.Length > 1024 Then
                SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
                sSQL.Clear()
            End If
            Value += 1
            Application.DoEvents()
        Next
        If sSQL.Length > 0 Then
            SysAD.DB(sLRDB).ExecuteSQL(sSQL.ToString)
            sSQL.Clear()
        End If
    End Sub

    Private Function Get検索文字(ByVal St As String) As String
        St = Replace(St, " ", "")

        Return St
    End Function
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

Public Class 個人InfoUpdate
    Inherits TableCheckAndUpdate

    Public Sub New(ByVal sTableName As String, ByRef pTarget As DataTable)
        MyBase.New(SysAD.DB(sLRDB), sTableName, pTarget, SysAD.DB(sLRDB).UpdateLog)
    End Sub

    Public Overrides Function CheckStart() As TableCheckAndUpdate.CheckResult

        Dim bCheck As CheckResult = CheckResult.NoUpdate
        bCheck = bCheck Or Check個人Info基本20150101_01(bCheck)
        bCheck = bCheck Or Check個人Info基本20150202_01(bCheck)
        bCheck = bCheck Or Check個人Info基本20150203_01(bCheck)
        bCheck = bCheck Or Check個人Info日置特殊2014_01(bCheck)
        Return bCheck

    End Function

    Public Function Check個人Info基本20150203_01(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then

            MakeTBLFieldModifySQL("送付先住所", "VARCHAR(255)")
            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function



    Public Function Check個人Info基本20150202_01(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then

            MakeTBLFieldModifySQL("送付先郵便番号", "VARCHAR(255)")
            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    Public Function Check個人Info日置特殊2014_01(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("別居", "BIT")
            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function



    Public Function Check個人Info基本20150101_01(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then

            'MakeTBLFieldDropSQL("認定農業者")
            'MakeTBLFieldDropSQL("担い手")
            'MakeTBLFieldDropSQL("登録斡旋")

            MakeTBLFieldModifySQL("検索フリガナ", "VARCAHR(255)")
            MakeTBLFieldModifySQL("登載申請関連者", "BIT")
            MakeTBLFieldModifySQL("農年加入受給種別", "LONG")

            MakeTBLFieldDropSQL("田所有")
            MakeTBLFieldDropSQL("畑所有")
            MakeTBLFieldDropSQL("樹所有")
            MakeTBLFieldDropSQL("田自作")
            MakeTBLFieldDropSQL("畑自作")
            MakeTBLFieldDropSQL("田遊休")
            MakeTBLFieldDropSQL("畑遊休")
            MakeTBLFieldDropSQL("樹遊休")
            MakeTBLFieldDropSQL("住所Work")
            MakeTBLFieldDropSQL("世帯連携")
            MakeTBLFieldDropSQL("作業者・作業PC")

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
    Public Function Check個人Info世帯員および就業20150129(bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then


            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

End Class

Public Class CTBL削除個人
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(DSet As DataSet, pTable As DataTable)
        MyBase.New(pTable, sLRDB)
    End Sub

    Public Overrides Sub MergePlus(pTable As System.Data.DataTable, Optional preserveChanges As Boolean = False, Optional pAction As System.Data.MissingSchemaAction = System.Data.MissingSchemaAction.Add)
        MyBase.MergePlus(pTable)
    End Sub
End Class