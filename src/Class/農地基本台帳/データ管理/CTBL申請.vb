'20160401霧島


Public Class CTBL申請
    Inherits HimTools2012.Data.DataTableWith

    Public Sub New(DSet As DataSet, pTable As DataTable)
        MyBase.New(pTable, sLRDB)
        Try
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [ID]=0")
            pTBL.TableName = "D_申請"
            Dim bReLoad As Integer = 0
            bReLoad += AlterTableADD(pTBL, "調査年月日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "農業委員1", "LONG")
            bReLoad += AlterTableADD(pTBL, "農業委員2", "LONG")
            bReLoad += AlterTableADD(pTBL, "農業委員3", "LONG")
            bReLoad += AlterTableADD(pTBL, "代理人住所", "VARCHAR(255)")
            bReLoad += AlterTableADD(pTBL, "農地の広がり", "LONG")
            bReLoad += AlterTableADD(pTBL, "土地改良事業の有無", "BIT")
            bReLoad += AlterTableADD(pTBL, "土地改良区の意見書の有無", "BIT")
            bReLoad += AlterTableADD(pTBL, "支払方法CD", "LONG")
            bReLoad += AlterTableADD(pTBL, "単位面積支払額", "LONG")
            bReLoad += AlterTableADD(pTBL, "経由法人ID", "LONG")
            bReLoad += AlterTableADD(pTBL, "申請地目安", "TEXT")
            bReLoad += AlterTableADD(pTBL, "関連ファイル", "LONGTEXT")
            bReLoad += AlterTableADD(pTBL, "申請後農地分類", "LONG")
            bReLoad += AlterTableADD(pTBL, "土地改良区の意見書の有不用", "LONG")
            bReLoad += AlterTableADD(pTBL, "解除条件付きの農地の貸借", "BIT")
            bReLoad += AlterTableADD(pTBL, "営農類型", "LONG")
            bReLoad += AlterTableADD(pTBL, "各筆申請情報", "LONGTEXT")
            bReLoad += AlterTableADD(pTBL, "10a金額", "Currency")
            bReLoad += AlterTableADD(pTBL, "不許可例外", "VARCHAR(255)")
            bReLoad += AlterTableADD(pTBL, "機構配分計画中間管理権取得日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画意見回答日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画知事公告日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画認可通知日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画権利設定内容", "LONG")
            bReLoad += AlterTableADD(pTBL, "機構配分計画利用配分計画始期日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画利用配分計画終期日", "DATETIME")
            bReLoad += AlterTableADD(pTBL, "機構配分計画利用配分計画借賃額", "LONG")
            bReLoad += AlterTableADD(pTBL, "機構配分計画利用配分計画10a賃借料", "Currency")
            bReLoad += AlterTableADD(pTBL, "意見聴取案件", "LONG")
            bReLoad += AlterTableADD(pTBL, "行政書士", "VARCHAR(255)")
            bReLoad += AlterTableADD(pTBL, "自由項目1", "TEXT")

            Select Case New 申請InfoUpdate("D_申請", pTBL).CheckStart
                Case TableCheckAndUpdate.CheckResult.CompleteUpdate
                    bReLoad = True
                Case TableCheckAndUpdate.CheckResult.UpdateFailed
                    MsgBox("データベースを更新しようと試みましたが、他のユーザーが使用中の為施工できません。他のユーザーの作業が終了したのち再起動してください。", MsgBoxStyle.Information)
                    End
            End Select

            If bReLoad <> 0 Then
                pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_申請] WHERE [ID]=0")
            End If

            Me.MergePlus(pTBL)
            SetTableEnv(DSet, "")
            DataInitAfter(DSet)
        Catch ex As Exception
            MsgBox(ex.Source & vbCrLf & ex.Message, , "例外処理")
        End Try
    End Sub

    Public Sub SetTableEnv(ByRef DSet As DataSet, ByVal StUp As String)
        DSet.Relations.Add(New DataRelation("申請者A", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("申請者A"), False))
        DSet.Relations.Add(New DataRelation("申請者B", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("申請者B"), False))
        DSet.Relations.Add(New DataRelation("経由法人", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("経由法人ID"), False))
        DSet.Relations.Add(New DataRelation("代理人A", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("代理人A"), False))
        'DSet.Relations.Add(New DataRelation("管理者", App農地基本台帳.TBL個人.Columns("ID"), Me.Columns("管理者"), False))

        DSet.Relations.Add(New DataRelation("申請時農地区分", DSet.Tables("申請時農地区分").Columns("ID"), Me.Columns("農地区分"), False))
        DSet.Relations.Add(New DataRelation("申請時農振区分", DSet.Tables("申請時農振区分").Columns("ID"), Me.Columns("農振区分"), False))
        DSet.Relations.Add(New DataRelation("申請時都計区分", DSet.Tables("申請時都計区分").Columns("ID"), Me.Columns("都市計画区分"), False))
        DSet.Relations.Add(New DataRelation("農業委員01", DSet.Tables("農業委員").Columns("ID"), Me.Columns("農業委員1"), False))
        DSet.Relations.Add(New DataRelation("農業委員02", DSet.Tables("農業委員").Columns("ID"), Me.Columns("農業委員2"), False))
        DSet.Relations.Add(New DataRelation("農業委員03", DSet.Tables("農業委員").Columns("ID"), Me.Columns("農業委員3"), False))


        Try
            Me.Columns.Add(New DataColumn("申請者情報氏名A", GetType(String), "Parent(申請者A).氏名"))
            Me.Columns.Add(New DataColumn("申請者情報郵便番号A", GetType(String), "Parent(申請者A).郵便番号"))
            Me.Columns.Add(New DataColumn("申請者情報氏名B", GetType(String), "Parent(申請者B).氏名"))
            Me.Columns.Add(New DataColumn("申請者情報郵便番号B", GetType(String), "Parent(申請者B).郵便番号"))
            Me.Columns.Add(New DataColumn("経由法人名", GetType(String), "Parent(経由法人).氏名"))

            'Me.Columns.Add(New DataColumn("管理者氏名", GetType(String), "Parent(管理者).氏名"))
            'Me.Columns.Add(New DataColumn("管理者郵便番号", GetType(String), "Parent(管理者).郵便番号"))
            'Me.Columns.Add(New DataColumn("管理者住所", GetType(String), "Parent(管理者).住所"))

            'me.Columns.Add(New DataColumn("許可年", GetType(String), "Convert(許可年月日,'System.Strig')"))
            Me.Columns.Add(New DataColumn("代理人情報氏名A", GetType(String), "Parent(代理人A).氏名"))
            Me.Columns.Add(New DataColumn("代理人情報住所A", GetType(String), "Parent(代理人A).住所"))

            Me.Columns.Add(New DataColumn("農地区分名称", GetType(String), "Parent(申請時農地区分).名称"))
            Me.Columns.Add(New DataColumn("農振区分名称", GetType(String), "Parent(申請時農振区分).名称"))
            Me.Columns.Add(New DataColumn("都計区分名称", GetType(String), "Parent(申請時都計区分).名称"))

            Me.Columns.Add(New DataColumn("農業委員01氏名", GetType(String), "Parent(農業委員01).名称"))
            Me.Columns.Add(New DataColumn("農業委員02氏名", GetType(String), "Parent(農業委員02).名称"))
            Me.Columns.Add(New DataColumn("農業委員03氏名", GetType(String), "Parent(農業委員03).名称"))
            Me.Columns.Add(New DataColumn("再設定表示", GetType(String), "IIF([再設定],'再設定','新規')"))

            Dim n中間管理機構ID As Integer = Val(SysAD.DB(sLRDB).DBProperty("中間管理機構ID"))
            If Not n中間管理機構ID = 0 Then
                Me.Columns.Add(New DataColumn("中間管理Flag", GetType(Integer), "IIF(経由法人ID=" & n中間管理機構ID & ",0,1)"))
            Else
                Me.Columns.Add(New DataColumn("中間管理Flag", GetType(Integer), "0"))
            End If

        Catch ex As Exception
            Stop
        End Try

    End Sub

    Public Overrides Sub MergePlus(pTable As System.Data.DataTable, Optional preserveChanges As Boolean = False, Optional pAction As System.Data.MissingSchemaAction = System.Data.MissingSchemaAction.Add)
        MyBase.MergePlus(pTable, preserveChanges, pAction)
        Try
            Dim nID As New List(Of String)
            For Each pRow As DataRow In pTable.Rows
                If Not IsDBNull(pRow.Item("申請者A")) AndAlso App農地基本台帳.TBL個人.Rows.Find(pRow.Item("申請者A")) Is Nothing AndAlso pRow.Item("申請者A") <> 0 AndAlso Not nID.Contains(Val(pRow.Item("申請者A")).ToString) Then
                    nID.Add(Val(pRow.Item("申請者A")).ToString)
                End If
                If Not IsDBNull(pRow.Item("申請者B")) AndAlso App農地基本台帳.TBL個人.Rows.Find(pRow.Item("申請者B")) Is Nothing AndAlso pRow.Item("申請者B") <> 0 AndAlso Not nID.Contains(Val(pRow.Item("申請者B")).ToString) Then
                    nID.Add(Val(pRow.Item("申請者B")).ToString)
                End If
                If Not IsDBNull(pRow.Item("申請者C")) AndAlso App農地基本台帳.TBL個人.Rows.Find(pRow.Item("申請者C")) Is Nothing AndAlso pRow.Item("申請者C") <> 0 AndAlso Not nID.Contains(Val(pRow.Item("申請者C")).ToString) Then
                    nID.Add(Val(pRow.Item("申請者C")).ToString)
                End If
                If Not IsDBNull(pRow.Item("経由法人ID")) AndAlso App農地基本台帳.TBL個人.Rows.Find(pRow.Item("経由法人ID")) Is Nothing AndAlso pRow.Item("経由法人ID") <> 0 AndAlso Not nID.Contains(Val(pRow.Item("経由法人ID")).ToString) Then
                    nID.Add(Val(pRow.Item("経由法人ID")).ToString)
                End If
                If Not IsDBNull(pRow.Item("代理人A")) AndAlso App農地基本台帳.TBL個人.Rows.Find(pRow.Item("代理人A")) Is Nothing AndAlso pRow.Item("代理人A") <> 0 AndAlso Not nID.Contains(Val(pRow.Item("代理人A")).ToString) Then
                    nID.Add(Val(pRow.Item("代理人A")).ToString)
                End If
                If nID.Count > 50 Then
                    App農地基本台帳.TBL個人.FindRowBySQL(String.Format("[ID] IN ({0})", Join(nID.ToArray, ",")))
                    nID.Clear()
                End If
            Next

            If nID.Count > 0 Then
                App農地基本台帳.TBL個人.FindRowBySQL(String.Format("[ID] IN ({0})", Join(nID.ToArray, ",")))
            End If


        Catch ex As Exception
            If Not SysAD.IsClickOnceDeployed Then
                Stop
            End If
        End Try
    End Sub

End Class


Public Class C申請関連情報取り込み
    Inherits HimTools2012.clsAccessor
    Private mvarBody As DataTable

    Public Sub New(ByVal pBody As DataTable)
        MyBase.new()
        mvarBody = pBody
    End Sub

    Public Overrides Sub Execute()
        Dim pDic As New List(Of String)

        Message = "申請者取得.."

        Dim p申請者A = From q In mvarBody Where Not IsDBNull(q.Item("申請者A")) AndAlso Not q.Item("申請者A") = 0 AndAlso IsDBNull(q.Item("申請者情報氏名A")) = True
        For Each pRow As DataRow In p申請者A
            If Not pDic.Contains(pRow.Item("申請者A").ToString) Then
                pDic.Add(pRow.Item("申請者A").ToString)
            End If
        Next
        Dim p申請者B = From q In mvarBody Where Not IsDBNull(q.Item("申請者B")) AndAlso Not q.Item("申請者B") = 0 AndAlso IsDBNull(q.Item("申請者情報氏名B")) = True
        For Each pRow As DataRow In p申請者B
            If Not pDic.Contains(pRow.Item("申請者B").ToString) Then
                pDic.Add(pRow.Item("申請者B").ToString)
            End If
        Next


        Value = 0
        Maximum = pDic.Count
        Message = String.Format("データベースより関係者名取得..({0}/{1})", Value, Maximum)

        Dim sB As New System.Text.StringBuilder
        Dim sC As String = ""
        For i As Integer = 0 To pDic.Count - 1
            sB.Append(sC & pDic.Item(i))


            If (i Mod 64) = 63 AndAlso sB.Length > 0 Then
                Value = i
                Message = String.Format("データベースより関係者名取得..({0}/{1})", i, pDic.Count)
                Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & sB.ToString & ")")
                App農地基本台帳.TBL個人.MergePlus(pAddK)
                sB.Clear()
                sC = ""
            Else
                sC = ","
            End If
        Next

        If sB.Length > 0 Then
            Dim pAddK As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:個人Info] WHERE ID In (" & sB.ToString & ")")
            App農地基本台帳.TBL個人.MergePlus(pAddK)
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

Public Class 申請InfoUpdate
    Inherits TableCheckAndUpdate

    Public Sub New(ByVal sTableName As String, ByRef pTarget As DataTable)
        MyBase.New(SysAD.DB(sLRDB), sTableName, pTarget, SysAD.DB(sLRDB).UpdateLog)
    End Sub

    Public Overrides Function CheckStart() As TableCheckAndUpdate.CheckResult

        Dim bCheck As CheckResult = CheckResult.NoUpdate
        bCheck = Check申請共通項目20150620(bCheck)
        bCheck = Check申請耕作目的追加項目20150617(bCheck)
        bCheck = Check申請貸借終了追加項目20150617(bCheck)
        bCheck = Check申請農地転用追加項目20150617(bCheck)

        Return bCheck
    End Function

    ''' <summary>
    ''' 共通項目-農地の権利移動・借賃等調査
    ''' </summary>
    ''' <param name="bCheck"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check申請共通項目20150620(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("調査適用法令", "LONG")         'カ
            MakeTBLFieldModifySQL("調査整理番号", "LONG")         'キ
            MakeTBLFieldModifySQL("調査許可等年月日", "DATETIME") 'ク

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    ''' <summary>
    ''' 様式1-耕作目的の権利の設定・移転
    ''' </summary>
    ''' <param name="bCheck"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check申請耕作目的追加項目20150617(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("調査権利の種類", "LONG")         'サ
            MakeTBLFieldModifySQL("調査農法3条2項5号", "LONG")      '１
            MakeTBLFieldModifySQL("調査農法3条2項124号", "LONG")    '２
            MakeTBLFieldModifySQL("調査個人法人の別A", "LONG")      '６
            MakeTBLFieldModifySQL("調査法人の形態別", "LONG")       '７
            MakeTBLFieldModifySQL("調査経営改善計画の有無", "LONG") '８
            MakeTBLFieldModifySQL("調査個人法人の別B", "LONG")      '９

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    ''' <summary>
    ''' 様式2-貸借の終了
    ''' </summary>
    ''' <param name="bCheck"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check申請貸借終了追加項目20150617(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("調査許可等の根拠条項", "LONG")           '２４
            MakeTBLFieldModifySQL("調査基盤法満了農地状況", "LONG")         '２５
            MakeTBLFieldModifySQL("調査中間管理事業法満了農地状況", "LONG") '２６

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function

    ''' <summary>
    ''' 様式3-農地等の転用
    ''' </summary>
    ''' <param name="bCheck"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Check申請農地転用追加項目20150617(ByVal bCheck As CheckResult) As CheckResult
        If Not CheckLog(Reflection.MethodBase.GetCurrentMethod.Name) Then
            MakeTBLFieldModifySQL("調査許可等の除外条項", "LONG")         '３１
            MakeTBLFieldModifySQL("調査土地利用計画区域区分", "LONG")     '３２
            MakeTBLFieldModifySQL("調査転用に伴う農用地区域除外", "LONG") '３３
            MakeTBLFieldModifySQL("調査転用主体", "LONG")                 '３４
            MakeTBLFieldModifySQL("調査転用用途", "LONG")                 '３５
            MakeTBLFieldModifySQL("調査一時転用該当有無", "LONG")         '３６
            MakeTBLFieldModifySQL("調査転用農地区分", "LONG")             '３７
            MakeTBLFieldModifySQL("調査優良農地許可判断根拠", "LONG")     '３８

            Return AddLog(Reflection.MethodBase.GetCurrentMethod.Name, bCheck)
        Else
            Return bCheck
        End If
    End Function
End Class
