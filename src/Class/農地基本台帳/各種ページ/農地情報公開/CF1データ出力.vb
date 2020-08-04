
Public Class CF1データ出力
    Inherits C各フェーズCSV出力

    Public DSET As New DataSet
    Public 出力TBL As DataTable
    Public TBL公開用個人 As DataTable

    Public Sub New()
        Me.Start(True, True)
    End Sub

    Public Overrides Sub Execute()
        都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
        市町村CD = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)
        市町村名 = SysAD.DB(sLRDB).DBProperty("市町村名")

        Message = "重複している耕作者番号の削除中..."

        Dim pDelTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_公開用個人.[PID], D_公開用個人.[AutoID] FROM D_公開用個人 WHERE (((D_公開用個人.[PID]) In (SELECT [PID] FROM [D_公開用個人] As Tmp GROUP BY [PID] HAVING Count(*)>1 ))) ORDER BY D_公開用個人.[PID];")
        Do
            Dim nID As New List(Of String)
            For Each pRow As DataRow In pDelTBL.Rows
                nID.Add(CStr(pRow.Item("AutoID")))
            Next
            If nID.Count > 0 Then
                SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_公開用個人] WHERE [AutoID] IN (" & Join(nID.ToArray, ",") & ")")
            End If
            pDelTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT D_公開用個人.[PID], D_公開用個人.[AutoID] FROM D_公開用個人 WHERE (((D_公開用個人.[PID]) In (SELECT [PID] FROM [D_公開用個人] As Tmp GROUP BY [PID] HAVING Count(*)>1 ))) ORDER BY D_公開用個人.[PID];")
        Loop Until pDelTBL.Rows.Count = 0


        Message = "データ初期化中...数分かかります"
        出力TBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].ID AS 筆ID, [D:農地Info].所有者ID, [D:個人Info].氏名 AS 所有者氏名, [D:農地Info].大字ID, V_大字.大字, [D:農地Info].小字ID, V_小字.小字, [D:農地Info].一部現況, [D:農地Info].地番, [D:農地Info].現況地目, V_現況地目.名称 AS 現況地目名, [D:農地Info].登記簿地目, V_地目.名称 AS 登記簿地目名, [D:農地Info].登記簿面積, [D:農地Info].実面積, [D:農地Info].農振法区分, [D:農地Info].都市計画法区分, [D:農地Info].生産緑地法, [D:農地Info].所有者農地意向, [D:農地Info].自小作別, [D:農地Info].小作形態," & _
                                        " [D:農地Info].小作開始年月日, [D:農地Info].小作終了年月日, [D:農地Info].利用状況調査日, [D:農地Info].利用状況調査農地法, [D:農地Info].利用意向調査日, [D:農地Info].利用意向根拠条項, [D:農地Info].利用意向意思表明日, [D:農地Info].利用意向意向内容区分, [D:農地Info].利用意向権利関係調査区分, [D:農地Info].利用意向公示年月日, [D:農地Info].利用意向通知年月日, [D:農地Info].農地法35の1通知日, [D:農地Info].農地法35の2通知日, [D:農地Info].農地法35の3通知日, [D:農地Info].勧告年月日, [D:農地Info].勧告内容, [D:農地Info].中間管理勧告日, [D:農地Info].再生利用困難農地," & _
                                        " [D:農地Info].農地法40裁定公告日, [D:農地Info].農地法43裁定公告日, [D:農地Info].農地法44の1裁定公告日, [D:農地Info].農地法44の3裁定公告日, [D:農地Info].中間管理権取得日, [D:農地Info].権利設定内容, [D:農地Info].利用配分計画設定期間年, [D:農地Info].利用配分計画設定期間月, [D:農地Info].利用配分計画始期日, [D:農地Info].利用配分計画終期日, [D:農地Info].管理者ID, [D:個人Info_1].氏名 AS 管理者氏名, [D:農地Info].借受人ID, [D:個人Info_2].氏名 AS 借受人氏名 " & _
                                        " FROM (((((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID) LEFT JOIN [D:個人Info] ON [D:農地Info].所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON [D:農地Info].管理者ID = [D:個人Info_1].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_2] ON [D:農地Info].借受人ID = [D:個人Info_2].ID" & _
                                        " WHERE ((([D:農地Info].大字ID)>0) AND (([D:農地Info].都市計画法区分) Is Null Or ([D:農地Info].都市計画法区分)<>1));")
        With 出力TBL
            .Columns.Add(New DataColumn("本番区分", GetType(String)))
            .Columns.Add(New DataColumn("本番", GetType(Integer)))
            .Columns.Add(New DataColumn("枝番区分", GetType(String)))
            .Columns.Add(New DataColumn("枝番", GetType(Integer)))
            .Columns.Add(New DataColumn("孫番区分", GetType(String)))
            .Columns.Add(New DataColumn("孫番", GetType(Integer)))
            .Columns.Add(New DataColumn("耕作者整理番号", GetType(String)))
            .Columns.Add(New DataColumn("小作設定期間年", GetType(Integer)))
            .Columns.Add(New DataColumn("小作設定期間月", GetType(Integer)))
            .Columns.Add(New DataColumn("市町村コード", GetType(Integer), Val(都道府県ID) & HimTools2012.StringF.Left(市町村CD & "0000", 4)))
            .Columns.Add(New DataColumn("市町村名", GetType(String), "'" & SysAD.市町村.市町村名 & "'"))
        End With

        TBL公開用個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_公開用個人")
        TBL公開用個人.PrimaryKey = New DataColumn() {TBL公開用個人.Columns("PID")}

        TBL個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].* AS D, * FROM [D:個人Info] WHERE ((([D:個人Info].氏名) Is Not Null Or ([D:個人Info].氏名)<>''));")
        TBL個人.PrimaryKey = {TBL個人.Columns("ID")}

        Dim TBL登記簿地目 As New 登記簿地目変換()
        TBL登記簿地目.Init()
        Dim TBL現況地目 As New 現況地目変換()
        TBL現況地目.Init()

        出力TBL.TableName = "対象農地明細"
        TBL登記簿地目.TableName = "登記簿地目変換TBL"
        TBL現況地目.TableName = "現況地目変換TBL"
        DSET.Tables.Add(出力TBL)
        DSET.Tables.Add(TBL登記簿地目)
        DSET.Tables.Add(TBL現況地目)

        DSET.Relations.Add("登記", TBL登記簿地目.Columns("名称"), 出力TBL.Columns("登記簿地目名"), False)
        出力TBL.Columns.Add("変換登記簿地目", GetType(Integer), "Parent(登記).ID")
        DSET.Relations.Add("現況", TBL現況地目.Columns("名称"), 出力TBL.Columns("現況地目名"), False)
        出力TBL.Columns.Add("変換現況地目", GetType(Integer), "Parent(現況).ID")

        For Each pRow As DataRow In 出力TBL.Rows
            Conv地番(pRow)
            Sub小作形態(pRow)
            Sub日付関連(pRow)
        Next

        CSVデータ出力()
    End Sub

    Private Sub Sub小作形態(ByRef pRow As DataRow)
        Select Case pRow.Item("小作形態").ToString
            Case "1" : pRow.Item("小作形態") = 5
            Case "2" : pRow.Item("小作形態") = 4
            Case "3", "7", "8", "9" : pRow.Item("小作形態") = 6
            Case "4" : pRow.Item("小作形態") = 1
            Case "5" : pRow.Item("小作形態") = 2
            Case "6" : pRow.Item("小作形態") = 3
            Case Else
        End Select
    End Sub

    Private Sub Sub日付関連(ByRef pRow As DataRow)
        Dim 設定年 As Integer = 0
        Dim 設定月 As Integer = 0
        'Debug.Print(pRow.Item("小作開始年月日").ToString)
        If Val(pRow.Item("自小作別").ToString) = 0 Then
            pRow.Item("小作開始年月日") = DBNull.Value
            pRow.Item("小作終了年月日") = DBNull.Value
            pRow.Item("小作設定期間年") = DBNull.Value
            pRow.Item("小作設定期間月") = DBNull.Value
            pRow.Item("利用配分計画始期日") = DBNull.Value
            pRow.Item("利用配分計画終期日") = DBNull.Value
            pRow.Item("利用配分計画設定期間年") = DBNull.Value
            pRow.Item("利用配分計画設定期間月") = DBNull.Value
        Else
            If Not IsDBNull(pRow.Item("小作開始年月日")) And Not IsDBNull(pRow.Item("小作終了年月日")) Then
                If Year(pRow.Item("小作開始年月日")) >= 1945 And Year(pRow.Item("小作終了年月日")) >= 1945 Then
                    設定年 = Format(CDate(pRow.Item("小作終了年月日")), "yyyyMMdd") - Format(CDate(pRow.Item("小作開始年月日")), "yyyyMMdd")
                    設定年 = Math.Floor(設定年 / 100)
                    If HimTools2012.StringF.Right(設定年, 2) >= 1 And HimTools2012.StringF.Right(設定年, 2) <= 12 Then
                        設定月 = HimTools2012.StringF.Right(設定年, 2)
                    ElseIf HimTools2012.StringF.Right(設定年, 2) > 0 Then
                        設定月 = 12 - (100 - HimTools2012.StringF.Right(設定年, 2))
                    Else
                        設定月 = 0
                    End If
                    設定年 = Math.Floor(設定年 / 100)
                    pRow.Item("小作設定期間年") = 設定年
                    pRow.Item("小作設定期間月") = 設定月
                Else
                    pRow.Item("小作開始年月日") = DBNull.Value
                    pRow.Item("小作終了年月日") = DBNull.Value
                    pRow.Item("小作設定期間年") = DBNull.Value
                    pRow.Item("小作設定期間月") = DBNull.Value
                End If
            Else
                pRow.Item("小作開始年月日") = DBNull.Value
                pRow.Item("小作終了年月日") = DBNull.Value
                pRow.Item("小作設定期間年") = DBNull.Value
                pRow.Item("小作設定期間月") = DBNull.Value
            End If
            If Not IsDBNull(pRow.Item("利用配分計画始期日")) And Not IsDBNull(pRow.Item("利用配分計画終期日")) Then
                If Year(pRow.Item("利用配分計画始期日")) >= 1945 And Year(pRow.Item("利用配分計画終期日")) >= 1945 Then
                    設定年 = Format(CDate(pRow.Item("利用配分計画終期日")), "yyyyMMdd") - Format(CDate(pRow.Item("利用配分計画始期日")), "yyyyMMdd")
                    設定年 = Math.Floor(設定年 / 100)
                    If HimTools2012.StringF.Right(設定年, 2) >= 1 And HimTools2012.StringF.Right(設定年, 2) <= 12 Then
                        設定月 = HimTools2012.StringF.Right(設定年, 2)
                    ElseIf HimTools2012.StringF.Right(設定年, 2) > 0 Then
                        設定月 = 12 - (100 - HimTools2012.StringF.Right(設定年, 2))
                    Else
                        設定月 = 0
                    End If
                    設定年 = Math.Floor(設定年 / 100)
                    pRow.Item("利用配分計画設定期間年") = 設定年
                    pRow.Item("利用配分計画設定期間月") = 設定月
                Else
                    pRow.Item("利用配分計画始期日") = DBNull.Value
                    pRow.Item("利用配分計画終期日") = DBNull.Value
                    pRow.Item("利用配分計画設定期間年") = DBNull.Value
                    pRow.Item("利用配分計画設定期間月") = DBNull.Value
                End If
            Else
                pRow.Item("利用配分計画始期日") = DBNull.Value
                pRow.Item("利用配分計画終期日") = DBNull.Value
                pRow.Item("利用配分計画設定期間年") = DBNull.Value
                pRow.Item("利用配分計画設定期間月") = DBNull.Value
            End If
        End If
    End Sub

    Public Sub CSVデータ出力()
        Dim sCSV As New StringBEx("", EnumCnv.設定無)

        Message = "CSV出力中.."
        Dim p出力View As DataView = New DataView(出力TBL, "", "[市町村コード], [大字ID], [小字ID], [本番区分], [本番], [枝番区分], [枝番], [孫番区分], [孫番], [一部現況]", DataViewRowState.CurrentRows)
        Me.Maximum = p出力View.Count

        Dim TBL公開用個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_公開用個人]")
        TBL公開用個人.PrimaryKey = New DataColumn() {TBL公開用個人.Columns("PID")}

        Me.Value = 0
        Dim s市町村CD As String = HimTools2012.StringF.Right("000" & 都道府県ID, 2) & HimTools2012.StringF.Left(市町村CD & "0000", 4)

        For Each pRow As DataRowView In p出力View
            Try
                Dim pLine As New StringBEx("""" & s市町村CD & """", EnumCnv.半角)
                Me.Value += 1
                Message = "CSV出力中(" & Me.Value & "/" & Me.Maximum & ").."
                '基本的事項－基本的事項
                With pLine
                    .CnvData(pRow.Item("市町村名").ToString, EnumCnv.全角)
                    .CnvData(pRow.Item("大字ID").ToString, EnumCnv.半角)
                    .CnvData(pRow.Item("大字").ToString, EnumCnv.全角)
                    .CnvData(pRow.Item("小字ID").ToString, EnumCnv.字コード)
                    .CnvData(pRow.Item("小字").ToString, EnumCnv.外字)
                    .CnvData(pRow.Item("本番区分").ToString, EnumCnv.全角)
                    .CnvData(pRow.Item("本番").ToString, EnumCnv.半角)
                    .CnvData(pRow.Item("枝番区分").ToString, EnumCnv.全角)
                    .CnvData(pRow.Item("枝番").ToString, EnumCnv.半角)
                    .CnvData(pRow.Item("孫番区分").ToString, EnumCnv.全角)
                    .CnvData(pRow.Item("孫番").ToString, EnumCnv.半角)
                    .CnvData(IIf(Val(pRow.Item("一部現況").ToString) > 0, Val(pRow.Item("一部現況").ToString), ""), EnumCnv.全角)
                    .CnvData(pRow.Item("変換登記簿地目").ToString, EnumCnv.登記簿地目)
                    .CnvData(pRow.Item("変換現況地目").ToString, EnumCnv.現況地目)
                    .CnvData(Val(pRow.Item("登記簿面積").ToString), EnumCnv.面積)
                    .CnvData(Val(pRow.Item("実面積").ToString), EnumCnv.面積)
                    .CnvData(Val(pRow.Item("農振法区分").ToString), EnumCnv.半角)
                    If SysAD.市町村.市町村名 = "日置市" Then : .CnvData("", EnumCnv.半角)
                    Else : .CnvData(Val(pRow.Item("都市計画法区分").ToString), EnumCnv.半角)
                    End If

                    .CnvData(IIf(pRow.Item("生産緑地法").ToString = True, 2, 1), EnumCnv.半角)
                    .CnvData(IIf(Val(pRow.Item("所有者農地意向").ToString) > 5, 5, Val(pRow.Item("所有者農地意向").ToString)), EnumCnv.半角)
                    .CnvData(.Fnc耕作者整理番号(pRow, TBL公開用個人, 2), EnumCnv.半角)
                    '農地等の借地権等の設定状況－農地等の借地権等の設定状況
                    If Val(pRow.Item("自小作別").ToString) > 0 AndAlso IsDate(pRow.Item("小作終了年月日")) AndAlso pRow.Item("小作終了年月日") > Now Then
                        .CnvData(Val(pRow.Item("小作形態").ToString), EnumCnv.半角)
                        .CnvData(Val(pRow.Item("小作設定期間年").ToString), EnumCnv.半角)
                        .CnvData(Val(pRow.Item("小作設定期間月").ToString), EnumCnv.半角)
                        .CnvData(IIf(pRow.Item("小作開始年月日").ToString = "", "", pRow.Item("小作開始年月日")), EnumCnv.日付)
                        .CnvData(IIf(pRow.Item("小作終了年月日").ToString = "", "", pRow.Item("小作終了年月日")), EnumCnv.日付)
                    Else
                        .CnvData(0, EnumCnv.半角)
                        .CnvData("", EnumCnv.半角)
                        .CnvData("", EnumCnv.半角)
                        .CnvData("", EnumCnv.日付)
                        .CnvData("", EnumCnv.日付)
                    End If
                    '農地中間管理権と利用配分計画等－機構が農地中間管理権を取得した年月日
                    .CnvData(IIf(pRow.Item("中間管理権取得日").ToString = "", "", pRow.Item("中間管理権取得日")), EnumCnv.日付)
                    '農地中間管理権と利用配分計画等－利用配分計画
                    .CnvData(Val(pRow.Item("権利設定内容").ToString), EnumCnv.半角)
                    .CnvData(Val(pRow.Item("利用配分計画設定期間年").ToString), EnumCnv.半角)
                    .CnvData(Val(pRow.Item("利用配分計画設定期間月").ToString), EnumCnv.半角)
                    .CnvData(IIf(pRow.Item("利用配分計画始期日").ToString = "", "", pRow.Item("利用配分計画始期日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("利用配分計画終期日").ToString = "", "", pRow.Item("利用配分計画終期日")), EnumCnv.日付)
                    '利用状況調査－利用状況調査
                    .CnvData(IIf(pRow.Item("利用状況調査日").ToString = "", "", pRow.Item("利用状況調査日")), EnumCnv.日付)
                    .CnvData(IIf(Val(pRow.Item("利用状況調査農地法").ToString) > 0, Val(pRow.Item("利用状況調査農地法").ToString), 3), EnumCnv.半角)
                    '利用意向調査－利用意向調査
                    .CnvData(IIf(pRow.Item("利用意向調査日").ToString = "", "", pRow.Item("利用意向調査日")), EnumCnv.日付)
                    .CnvData(Val(pRow.Item("利用意向根拠条項").ToString), EnumCnv.半角)
                    '利用意向調査－調査結果
                    .CnvData(IIf(pRow.Item("利用意向意思表明日").ToString = "", "", pRow.Item("利用意向意思表明日")), EnumCnv.日付)
                    .CnvData(Val(pRow.Item("利用意向意向内容区分").ToString), EnumCnv.半角)
                    '利用意向調査－所有者が確知できない農地
                    .CnvData(IIf(Val(pRow.Item("利用意向権利関係調査区分").ToString) > 2, 0, Val(pRow.Item("利用意向権利関係調査区分").ToString)), EnumCnv.半角)
                    .CnvData(IIf(pRow.Item("利用意向公示年月日").ToString = "", "", pRow.Item("利用意向公示年月日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("利用意向通知年月日").ToString = "", "", pRow.Item("利用意向通知年月日")), EnumCnv.日付)
                    '農地中間管理機構等との協議等－農地中間管理機構との協議
                    .CnvData(IIf(pRow.Item("農地法35の1通知日").ToString = "", "", pRow.Item("農地法35の1通知日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("農地法35の2通知日").ToString = "", "", pRow.Item("農地法35の2通知日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("農地法35の3通知日").ToString = "", "", pRow.Item("農地法35の3通知日")), EnumCnv.日付)
                    '農地中間管理機構等との協議等－農地所有者への勧告
                    .CnvData(IIf(pRow.Item("勧告年月日").ToString = "", "", pRow.Item("勧告年月日")), EnumCnv.日付)
                    .CnvData(Val(pRow.Item("勧告内容").ToString), EnumCnv.半角)
                    .CnvData(IIf(pRow.Item("中間管理勧告日").ToString = "", "", pRow.Item("中間管理勧告日")), EnumCnv.日付)
                    '農地中間管理機構等との協議等－再生利用困難な農地
                    .CnvData(Val(pRow.Item("再生利用困難農地").ToString), EnumCnv.半角)
                    '裁定－裁定公告の状況
                    .CnvData(IIf(pRow.Item("農地法40裁定公告日").ToString = "", "", pRow.Item("農地法40裁定公告日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("農地法43裁定公告日").ToString = "", "", pRow.Item("農地法43裁定公告日")), EnumCnv.日付)
                    '措置命令－措置命令の内容
                    .CnvData(IIf(pRow.Item("農地法44の1裁定公告日").ToString = "", "", pRow.Item("農地法44の1裁定公告日")), EnumCnv.日付)
                    .CnvData(IIf(pRow.Item("農地法44の3裁定公告日").ToString = "", "", pRow.Item("農地法44の3裁定公告日")), EnumCnv.日付)
                End With

                sCSV.Body.AppendLine(pLine.Body.ToString)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            If _Cancel Then
                Throw New Exception("Cancel")
                Exit Sub
            End If
        Next

        '↓保存場所の確認だけ設定
        'Shift-JISで保存します。
        Dim dtNow As DateTime = DateTime.Now
        Dim dtToday As DateTime = dtNow.Date
        Dim StToday As String = dtToday.ToString("yyyyMMdd")
        Dim sPath As String = ""

        '/***名前を付けて保存***/
        With New SaveFileDialog
            .FileName = 都道府県ID & 市町村CD & "_" & StToday & "_公表用農地情報.csv"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            .Filter = "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"

            If .ShowDialog = DialogResult.OK Then
                sPath = .FileName
            End If
        End With

        Dim CSVText As New System.IO.StreamWriter(sPath, False, System.Text.Encoding.GetEncoding(932))
        CSVText.Write(sCSV.Body.ToString)
        CSVText.Dispose()

        MsgBox("CSVの出力が完了しました。")
        SysAD.ShowFolder(System.IO.Directory.GetParent(sPath).ToString)
    End Sub
End Class


