'調査中のコードに気を付ける
'対象列の内容を随時確認
'問題列があった場合フラグを立て
'レコードを丸々別テーブルに保存
'


Public Class CF2データ出力
    Inherits C各フェーズCSV出力

    Dim pOutPutType As Integer = 0
    Public Sub New(ByVal pValue As Integer)
        pOutPutType = pValue

        Me.Start(True, True)
    End Sub

    Public Overrides Sub Execute()
        'Try
        都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
        市町村CD = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)
        市町村名 = SysAD.DB(sLRDB).DBProperty("市町村名")

        Message = "個人情報読み込み中..."
        TBL個人 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].* AS D, * FROM [D:個人Info] WHERE ((Not ([D:個人Info].世帯ID)=0) AND (([D:個人Info].氏名) Is Not Null Or ([D:個人Info].氏名)<>''));")
        TBL個人.PrimaryKey = {TBL個人.Columns("ID")}

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

        Message = "耕作者情報読み込み中..."
        TBL耕作者 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_公開用個人")

        Message = "世帯情報読み込み中..."
        TBL世帯 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:世帯Info] WHERE Not ([D:世帯Info].世帯主ID)=0;")
        TBL世帯.PrimaryKey = {TBL世帯.Columns("ID")}

        Message = "農地情報読み込み中..."
        TBL農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:農地Info].* AS D, V_大字.名称 AS 大字, V_小字.名称 AS 小字, [V_大字].[名称] & IIf(IsNull([V_小字].[名称]),'',IIf([V_小字].[名称]='-','',[V_小字].[名称])) & [D:農地Info].[地番] AS 土地所在, V_地目.名称 AS 登記簿地目名, V_現況地目.名称 AS 現況地目名 " & _
                                                      "FROM ((([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D:農地Info].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D:農地Info].現況地目 = V_現況地目.ID " & _
                                                      "WHERE (((V_大字.名称) Is Not Null Or (V_大字.名称)<>'') AND (([D:農地Info].大字ID)>0) AND (([D:農地Info].地番) Is Not Null Or ([D:農地Info].地番)<>''));") '20161114追加(テスト)

        'Message = "転用済み農地情報読み込み中..."
        'TBL転用農地 = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D_転用農地].* AS D, V_大字.名称 AS 大字, V_小字.名称 AS 小字, [V_大字].[名称] & IIf(IsNull([V_小字].[名称]),'',IIf([V_小字].[名称]='-','',[V_小字].[名称])) & [D_転用農地].[地番] AS 土地所在, V_地目.名称 AS 登記簿地目名, V_現況地目.名称 AS 現況地目名 " & _
        '                                                  "FROM ((([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID) LEFT JOIN V_地目 ON [D_転用農地].登記簿地目 = V_地目.ID) LEFT JOIN V_現況地目 ON [D_転用農地].現況地目 = V_現況地目.ID " & _
        '                                                  "WHERE (((V_大字.名称) Is Not Null Or (V_大字.名称)<>'') AND (([D_転用農地].大字ID)>0) AND (([D_転用農地].地番) Is Not Null Or ([D_転用農地].地番)<>''));") '20161114追加(テスト)
        'TBL農地.Merge(TBL転用農地)    '20161114追加(テスト)

        Message = "市町村コード情報読み込み中..."
        Set市町村コード()

        '/***************20161114追加(テスト)***************/
        ColumnCheck(TBL農地, "本番区分", GetType(String))
        ColumnCheck(TBL農地, "本番", GetType(Integer))
        ColumnCheck(TBL農地, "枝番区分", GetType(String))
        ColumnCheck(TBL農地, "枝番", GetType(Integer))
        ColumnCheck(TBL農地, "孫番区分", GetType(String))
        ColumnCheck(TBL農地, "孫番", GetType(Integer))

        For Each pRow As DataRow In TBL農地.Rows
            Conv地番(pRow)
        Next
        '/**************************************************/

        Select Case pOutPutType
            Case EnumOutPutType.全件
                Message = "農地データ出力中..."
                Set取込用農地(pOutPutType)
                Message = "個人データ出力中..."
                Set取込用個人(pOutPutType)
                Message = "世帯・法人データ出力中..."
                Set取込用世帯・法人(pOutPutType)
            Case EnumOutPutType.農地
                Message = "農地データ出力中..."
                Set取込用農地(pOutPutType)
            Case EnumOutPutType.個人
                Message = "個人データ出力中..."
                Set取込用個人(pOutPutType)
            Case EnumOutPutType.世帯
                Message = "世帯・法人データ出力中..."
                Set取込用世帯・法人(pOutPutType)
        End Select

        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub

    Private Sub Set取込用農地(ByVal pValue As Integer)
        Dim pLineHeader As New StringBEx("市町村コード", EnumCnv.設定無)
        Dim pLineHeader論理 As New StringBEx("連番", EnumCnv.設定無)
        Dim pLineHeaderレイアウト As New StringBEx("連番", EnumCnv.設定無)
        Dim h取込用農地 As String() = GetHeader農地()
        Dim h取込用論理 As String() = GetHeader論理()
        Dim h取込用レイアウト As String() = GetHeaderレイアウト()
        Dim pView As DataView = New DataView(TBL農地, "", "[市町村ID],[大字ID],[小字ID],[本番区分],[本番],[枝番区分],[枝番],[孫番区分],[孫番],[一部現況]", DataViewRowState.CurrentRows)   '20161114追加(テスト)

        Dim TBL登記地目 As New 登記簿地目変換()
        TBL登記地目.Init()
        Dim TBL現況地目 As New 現況地目変換()
        TBL現況地目.Init()
        Dim TBL公開用個人 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_公開用個人]")
        TBL公開用個人.PrimaryKey = New DataColumn() {TBL公開用個人.Columns("PID")}

        Initialization()

        For n As Integer = 0 To UBound(h取込用農地)
            pLineHeader.mvarBody.Append("," & h取込用農地(n))
        Next
        sCSV.Body.AppendLine(pLineHeader.Body.ToString)

        For n = 0 To UBound(h取込用論理)
            pLineHeader論理.mvarBody.Append("," & h取込用論理(n))
        Next
        sCSV論理.Body.AppendLine(pLineHeader論理.Body.ToString)

        For n = 0 To UBound(h取込用レイアウト)
            pLineHeaderレイアウト.mvarBody.Append("," & h取込用レイアウト(n))
        Next
        sCSVレイアウト.Body.AppendLine(pLineHeaderレイアウト.Body.ToString)

        Me.Maximum = pView.Count
        Me.Value = 0

        For Each pRow As DataRowView In pView
            Me.Value += 1
            Message = "取込用農地データ出力中(" & Me.Value & "/" & pView.Count & ")..."

            Dim Conv市町村CD As String = 市町村コード(pRow.Item("土地所在").ToString)
            Dim pLineRow As New StringBEx(IIf(Len(pRow.Item("市町村ID").ToString) = 6, pRow.Item("市町村ID").ToString, IIf(Conv市町村CD = "", Find市町村コード(pRow.Item("土地所在").ToString), Conv市町村CD)), EnumCnv.半角, "1:市町村コード:6", True) ' 市町村コード
            With pLineRow
                .CnvData(pRow.Item("大字ID"), EnumCnv.半角, , "2:大字コード:8", True)
                .CnvData(pRow.Item("大字"), EnumCnv.全角, , "3:大字名:30", True)
                .CnvData(pRow.Item("小字ID"), EnumCnv.半角, , "4:小字コード:7")
                .CnvData(pRow.Item("小字"), EnumCnv.全角, , "5:小字名:20")
                .CnvData(pRow.Item("本番区分"), EnumCnv.全角, , "6:本番区分:2")
                .CnvData(pRow.Item("本番"), EnumCnv.半角, , "7:本番:6", True)
                .CnvData(pRow.Item("枝番区分"), EnumCnv.全角, , "8:枝番区分:2")
                .CnvData(pRow.Item("枝番"), EnumCnv.半角, , "9:枝番:6")
                .CnvData(pRow.Item("孫番区分"), EnumCnv.全角, , "10:枝番区分:2")
                .CnvData(pRow.Item("孫番"), EnumCnv.半角, , "11:孫番:6")
                .mvarBody.Append(",") '曾孫番区分
                .mvarBody.Append(",") '曾孫番
                .mvarBody.Append(",") '玄孫番区分
                .mvarBody.Append(",") '玄孫番
                .CnvData(IIf(Val(pRow.Item("一部現況").ToString) > 0, Val(pRow.Item("一部現況").ToString), ""), EnumCnv.全角, , "16:区分:4")
                .CnvData(Cnv農地ID(pRow.Item("ID")), EnumCnv.半角, , "17:耕地番号:8")
                .CnvData(pRow.Item("耕地番号作成日"), EnumCnv.日付, , "18:耕地番号の作成年月日（提供元の情報時点）:10")
                .Cnv地目(TBL登記地目, pRow.Item("登記簿地目名"), "19:登記簿地目:2")
                .Cnv地目(TBL現況地目, pRow.Item("現況地目名"), "20:現況地目:3")
                .CnvData(pRow.Item("登記簿面積"), EnumCnv.面積, , "21:登記簿面積:13", True)
                .CnvData(.Cnv部分面積(pRow), EnumCnv.面積, , "22:登記簿面積の内訳:13", True)
                .CnvData(pRow.Item("実面積"), EnumCnv.面積, , "23:現況面積:13", True)
                .CnvData(pRow.Item("本地面積"), EnumCnv.面積, , "24:本地面積:13")
                .CnvData(pRow.Item("本地面積作成日"), EnumCnv.日付, , "25:本地面積の作成年月日（提供元の情報時点）10")
                .CnvData(IIf(Val(pRow.Item("農振法区分").ToString) = 4, 5, IIf(Val(pRow.Item("農振法区分").ToString) = 5, 9, pRow.Item("農振法区分"))), EnumCnv.選択, , "26:農振法区分:1")
                .CnvData(IIf(SysAD.市町村.市町村名 = "日置市", 0, pRow.Item("都市計画法区分")), EnumCnv.選択, 6, "27:都市計画法区分:1")
                .CnvData(IIf(pRow.Item("生産緑地法") = True, 2, 1), EnumCnv.選択, , "28:生産緑地法に基づく指定:1")
                .CnvData(pRow.Item("生産緑地法指定日"), EnumCnv.日付, , "29:生産緑地法に基づく指定（指定年月日）10")
                .CnvData(pRow.Item("生産緑地法指定面積"), EnumCnv.面積, , "30:生産緑地法に基づく指定（指定面積）:13")
                .CnvData(IIf(Val(pRow.Item("生産緑地法種別").ToString) = 3, 1, IIf(Val(pRow.Item("生産緑地法種別").ToString) = 1, 2, 0)), EnumCnv.半角, , "31:生産緑地法に基づく指定（種別）:1")
                .CnvData(pRow.Item("農地種別"), EnumCnv.半角, , "32:農地種別:1")
                .CnvData(CnvID(pRow.Item("所有者ID")), EnumCnv.半角, , "33:所有者世帯員番号:13", True)
                .CnvData(pRow.Item("所有者農地意向"), EnumCnv.半角, , "34:所有者の農地に関する意向:1")
                .CnvData(pRow.Item("所有者農地意向その他"), EnumCnv.全角, , "35:所有者の農地に関する意向「その他」の内訳:100")
                .CnvData(IIf(Val(pRow.Item("農地法第52公表同意").ToString) = 0, 1, IIf(Val(pRow.Item("農地法第52公表同意").ToString) = 1, 2, 0)), EnumCnv.半角, , "36:農法第52条の3第1項による公表への同意:1")
                .CnvData(CnvID(.Fnc耕作者整理番号(pRow, TBL公開用個人, 1)), EnumCnv.半角, , "37:耕作者世帯員番号:13", True)
                .CnvData(CnvID(.Fnc耕作者整理番号(pRow, TBL公開用個人, 2)), EnumCnv.半角, , "38:耕作者整理番号:18")
                .CnvData(pRow.Item("耕作状況"), EnumCnv.選択, , "39:耕作状況:6")
                .CnvData(CnvID(pRow.Item("特定作業者ID")), EnumCnv.半角, , "40:作業者世帯員番号:13")
                .CnvData(pRow.Item("特定作業作目種別"), EnumCnv.全角, , "41:作物:20")
                .CnvData(pRow.Item("特定作業内容"), EnumCnv.全角, , "42:作業内容:40")
                .mvarBody.Append(",") '許可年月日
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, .Fnc適用法令(pRow.Item("小作地適用法")), ""), EnumCnv.選択, , "44:適用法:2")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, .Fnc小作形態(pRow.Item("小作形態")), ""), EnumCnv.選択, , "45:権利の種類:2")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("小作開始年月日"), ""), EnumCnv.日付, , "46:始期年月日:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("小作終了年月日"), ""), EnumCnv.日付, , "47:終期年月日:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("再設定終了年月日"), ""), EnumCnv.日付, , "48:再設定した場合の終期年月日:10")
                .mvarBody.Append(",") '1年間の借賃額
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("10a賃借料"), ""), EnumCnv.半角, , "50:10a当たりの借賃額:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("物納"), ""), EnumCnv.半角, , "51:物納:50")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("物納単位"), ""), EnumCnv.半角, , "52:物納単位:1")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("利用集積計画番号"), ""), EnumCnv.半角, , "53:利用集積計画番号:8")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("利用集積公告日"), ""), EnumCnv.日付, , "54:利用集積公告年月日:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("利用目的"), 0), EnumCnv.選択, 3, "55:利用目的:1")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("利用目的備考"), ""), EnumCnv.全角, , "56:利用目的備考:50")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("利用権設定区分"), ""), EnumCnv.選択, , "57:利用権設定区分:1")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("交付金判定"), ""), EnumCnv.選択, , "58:交付金判定:1")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("交付金対象額"), ""), EnumCnv.半角, , "59:交付金対象額:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("借受人ID"), ""), EnumCnv.半角, , "60:交付金対象額:13")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, .Fnc適用法令(pRow.Item("転貸適用法")), ""), EnumCnv.選択, , "61:転貸適用法:2")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, .Fnc小作形態(pRow.Item("転貸形態")), ""), EnumCnv.選択, , "62:転貸権利の種類:2")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸始期年月日"), ""), EnumCnv.日付, , "63:転貸始期年月日:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸終期年月日"), ""), EnumCnv.日付, , "64:転貸終期年月日:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸料"), ""), EnumCnv.半角, , "65:転貸1年間の借賃額:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸10a転貸料"), ""), EnumCnv.半角, , "66:転貸10a当たりの借賃額:10")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸物納"), ""), EnumCnv.半角, , "67:物納:50")
                .CnvData(IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("転貸料単位"), ""), EnumCnv.半角, , "68:物納単位:1")
                .CnvData(pRow.Item("中間管理権取得日"), EnumCnv.日付, , "69:機構が農地中間管理権を取得した年月日:10")
                .CnvData(pRow.Item("意見回答日"), EnumCnv.日付, , "70:利用配分計画案への意見回答年月日:10")
                .CnvData(pRow.Item("知事公告日"), EnumCnv.日付, , "71:利用配分計画の知事公告年月日:10")
                .CnvData(pRow.Item("認可通知日"), EnumCnv.日付, , "72:計画の認可通知年月日:10")
                .CnvData(pRow.Item("貸借契約解除年月日"), EnumCnv.日付, , "73:農地中間管理事業法20条に基づく貸借解約年月日:10")
                .CnvData(pRow.Item("納税猶予対象農地"), EnumCnv.選択, 3, "74:納税猶予:1")
                .CnvData(pRow.Item("納税猶予種別"), EnumCnv.選択, 3, "75:種別:1")
                .CnvData(pRow.Item("納税猶予相続日"), EnumCnv.日付, , "76:相続日・贈与日:10")
                .CnvData(pRow.Item("納税猶予適用日"), EnumCnv.日付, , "77:適用年月日:10")
                .CnvData(pRow.Item("納税猶予継続日"), EnumCnv.日付, , "78:継続年月日:10")
                .CnvData(pRow.Item("納税猶予確認日"), EnumCnv.日付, , "79:確認年月日:10")
                .CnvData(pRow.Item("租税処置法"), EnumCnv.選択, 4, "80:特定貸付け根拠条項（租税特別措置法第70条の4の2第1項又は第70条の6の2第1項）:1")
                .CnvData(pRow.Item("営農困難"), EnumCnv.選択, 2, "81:営農困難時貸付け:1")
                .CnvData(pRow.Item("利用状況調査日"), EnumCnv.日付, , "82:調査実施年月日:10")
                .CnvData(pRow.Item("利用状況調査農地法"), EnumCnv.選択, 6, "83:農地法第32条第1項第1号:1")
                .CnvData(pRow.Item("利用状況調査荒廃"), EnumCnv.選択, 3, "84:荒廃農地調査分類:1")
                .CnvData(pRow.Item("利用状況"), EnumCnv.全角, , "85:利用状況:80")
                .CnvData(.Find農家情報(Val(pRow.Item("利用状況調査委員ID").ToString), Enum農家.氏名), EnumCnv.全角, , "86:調査委員名:80")
                .CnvData(pRow.Item("利用状況耕作放棄地通し番号"), EnumCnv.半角, , "87:耕作放棄地通し番号:8")
                .CnvData(IIf(Val(pRow.Item("利用状況調査転用").ToString) = 1, pRow.Item("利用状況一時転用区分"), 0), EnumCnv.選択, 4, "88:一時転用:1")
                .CnvData(IIf(Val(pRow.Item("利用状況調査転用").ToString) = 2, 2, 1), EnumCnv.選択, , "89:無断転用:1")
                .CnvData(IIf(Val(pRow.Item("利用状況調査転用").ToString) = 3, 2, 1), EnumCnv.選択, , "90:違反転用:1")
                .CnvData(pRow.Item("利用状況調査不可判断日"), EnumCnv.日付, , "91:調査不可と判断した年月日:10")
                .CnvData(pRow.Item("利用状況調査不可判断理由"), EnumCnv.選択, 3, "92:理由:1")
                .CnvData(pRow.Item("利用状況調査不可判断その他理由"), EnumCnv.全角, , "93:理由「その他」の内訳:100")
                .CnvData(pRow.Item("利用意向調査日"), EnumCnv.日付, , "94:調査実施年月日:10")
                .CnvData(pRow.Item("利用意向根拠条項"), EnumCnv.選択, 4, "95:根拠条項:1")
                .CnvData(pRow.Item("利用意向意思表明日"), EnumCnv.日付, , "96:所有者意思表明年月日:10")
                .CnvData(pRow.Item("利用意向意向内容区分"), EnumCnv.選択, 6, "97:調査結果:1")
                .CnvData(pRow.Item("利用意向調査結果その他理由"), EnumCnv.全角, , "98:調査結果のその他任意文字:100")
                .CnvData(pRow.Item("利用意向措置実施状況"), EnumCnv.全角, , "99:措置の実施状況:100")
                .CnvData(IIf(Val(pRow.Item("利用意向権利関係調査区分").ToString) = 3, 0, Val(pRow.Item("利用意向権利関係調査区分").ToString)), EnumCnv.選択, , "100:権利関係調査:1")
                .CnvData(pRow.Item("利用意向調査不可年月日"), EnumCnv.日付, , "101:調査結果年月日:10")
                .CnvData(pRow.Item("利用意向調査不可結果"), EnumCnv.選択, , "102:調査結果:1")
                .CnvData(pRow.Item("利用意向権利関係調査記録"), EnumCnv.全角, , "103:調査結果のその他任意文字:100")
                .CnvData(pRow.Item("利用意向公示年月日"), EnumCnv.日付, , "104:農地法第32条第3項に基づく公示年月日:10")
                .CnvData(pRow.Item("利用意向通知年月日"), EnumCnv.日付, , "105:農地法第43条第1項に基づく農地中間管理機構への通知発出年月日:10")
                .CnvData(pRow.Item("農地法35の1通知日"), EnumCnv.日付, , "106:農地法第35条第1項に基づく通知発出年月日:10")
                .CnvData(pRow.Item("農地法35の2通知日"), EnumCnv.日付, , "107:農地法第35条第2項に基づく協議を行わない通知発出年月日:10")
                .CnvData(pRow.Item("農地法35の2申入日"), EnumCnv.日付, , "108:農地法第35条第2項に基づく協議を機構が所有者に申し入れた年月日:10")
                .CnvData(pRow.Item("農地法35の3通知日"), EnumCnv.日付, , "109:農地法第35条第3項に基づく通知発出年月日:10")
                .CnvData(pRow.Item("勧告年月日"), EnumCnv.日付, , "110:勧告年月日:10")
                .CnvData(pRow.Item("勧告内容"), EnumCnv.選択, 6, "111:勧告内容:1")
                .CnvData(pRow.Item("中間管理勧告日"), EnumCnv.日付, , "112:農地中間管理機構等への通知発出年月日:10")
                .CnvData(pRow.Item("再生利用困難農地"), EnumCnv.選択, 7, "113:再生利用困難な農地:1")
                .CnvData(pRow.Item("農地法40裁定公告日"), EnumCnv.日付, , "114:農地法第40条に基づく裁定公告日:10")
                .CnvData(pRow.Item("農地法43裁定公告日"), EnumCnv.日付, , "115:農地法第43条に基づく裁定公告日:10")
                .CnvData(pRow.Item("農地法44の1裁定公告日"), EnumCnv.日付, , "116:農地法第44条第1項に基づく命令年月日:10")
                .CnvData(pRow.Item("農地法44の3裁定公告日"), EnumCnv.日付, , "117:農地法第44条第3項に基づく公告年月日:10")
                .CnvData(IIf(pRow.Item("利用状況報告対象") = True, 2, 1), EnumCnv.選択, , "118:利用状況報告の対象:1")
                .CnvData(pRow.Item("利用状況報告年月日"), EnumCnv.日付, , "119:利用状況報告年月日:10")
                .CnvData(pRow.Item("是正勧告日"), EnumCnv.日付, , "120:勧告年月日:10")
                .CnvData(pRow.Item("是正勧告内容"), EnumCnv.全角, , "121:内容:100")
                .CnvData(pRow.Item("是正期限"), EnumCnv.日付, , "122:期限:10")
                .Cnv利用状況根拠条項(pRow, "根拠", "123:根拠条項:1")
                .CnvData(pRow.Item("是正確認"), EnumCnv.日付, , "124:確認年月日:10")
                .CnvData(pRow.Item("是正状況"), EnumCnv.全角, , "125:是正状況:100")
                .CnvData(pRow.Item("取消年月日"), EnumCnv.日付, , "126:取消年月日:10")
                .CnvData(pRow.Item("取消事由"), EnumCnv.全角, , "127:取消事由:100")
                .Cnv利用状況根拠条項(pRow, "取消", "128:根拠条項:1")
                .CnvData(pRow.Item("届出年月日"), EnumCnv.日付, , "129:届出年月日:10")
                .CnvData(pRow.Item("届出事由"), EnumCnv.選択, , "130:届出事由:1")
                .CnvData(CnvID(pRow.Item("相続届出者ID")), EnumCnv.半角, , "131:権利取得者世帯員番号:13")
                .CnvData(.Fnc判定(pRow.Item("相続登記の有無")), EnumCnv.選択, , "132:相続登記の有無:1") '
                .CnvData(pRow.Item("仮登記日"), EnumCnv.日付, , "133:設定年月日:10")
                .CnvData(CnvID(pRow.Item("仮登記者ID")), EnumCnv.半角, , "134:仮登記権者世帯員番号:13") '
                .CnvData(IIf(Val(pRow.Item("環境保全交付金").ToString) = 1, 2, 1), EnumCnv.選択, , "135:環境保全型農業直接支払交付金:1")
                .CnvData(pRow.Item("環境保全交付基準日"), EnumCnv.日付, , "136:交付金対象の基準年月日（提供元の情報時点）:10")
                .CnvData(IIf(Val(pRow.Item("農地維持交付金").ToString) = 1, 2, 1), EnumCnv.選択, , "137:農地維持支払交付金:1")
                .CnvData(pRow.Item("農地維持交付基準日"), EnumCnv.日付, , "138:交付金対象の基準年月日（提供元の情報時点）:10")
                .CnvData(IIf(Val(pRow.Item("資源向上交付金").ToString) = 1, 2, 1), EnumCnv.選択, , "139:資源向上支払い交付金:1")
                .CnvData(pRow.Item("資源向上交付基準日"), EnumCnv.日付, , "140:交付金対象の基準年月日（提供元の情報時点）:10")
                .CnvData(IIf(Val(pRow.Item("中山間直接支払").ToString) = 1, 2, 1), EnumCnv.選択, , "141:中山間地域等直接支払:1")
                .CnvData(pRow.Item("中山間直接支払基準日"), EnumCnv.日付, , "142:交付金対象の基準年月日（提供元の情報時点）:10")
                .CnvData(IIf(Val(pRow.Item("特定処分対象農地等").ToString) = 1, 2, 1), EnumCnv.選択, , "143:特定処分対象農地:1")
                .CnvData(pRow.Item("農業者年金処分対象農地"), EnumCnv.選択, , "144:農業者年金処分対象農地:1")
                .CnvData(pRow.Item("農業者年金処分適用日"), EnumCnv.日付, , "145:農業者年金処分適用年月日:10")
                .CnvData(pRow.Item("転用適用法"), EnumCnv.選択, , "146:転用適用法:2")
                .CnvData(pRow.Item("転用形態"), EnumCnv.選択, 4, "147:転用形態:1")
                .CnvData(pRow.Item("転用用途"), EnumCnv.選択, , "148:転用用途:1")
                .CnvData(pRow.Item("転用換地有無"), EnumCnv.選択, , "149:転用換地有無:1")
                .CnvData(pRow.Item("転用始期年月日"), EnumCnv.日付, , "150:始期年月日:10")
                .CnvData(pRow.Item("転用終期年月日"), EnumCnv.日付, , "151:終期年月日:10")
                .CnvData(pRow.Item("土地改良法"), EnumCnv.選択, 4, "152:圃場整備:1")
                .CnvData(pRow.Item("区画整理"), EnumCnv.選択, 4, "153:区画整理:1")
                .CnvData(.Fnc判定(pRow.Item("前払いの有無")), EnumCnv.選択, , "154:前払いの有無:1")
                .CnvData(.Fnc判定(pRow.Item("指定の有無")), EnumCnv.選択, , "155:指定の有無:1")
                .CnvData(IIf(IsDBNull(pRow.Item("推測耕作者ID")), IIf(Val(pRow.Item("自小作別").ToString) > 0, pRow.Item("借受人ID"), pRow.Item("所有者ID")), pRow.Item("推測耕作者ID")), EnumCnv.半角, , "156:耕作しているであろう人の世帯員番号:13")
                .CnvData(pRow.Item("備考"), EnumCnv.全角, , "157:備考:100")

                sCSV.Body.AppendLine(.Body.ToString)


                '/***論理チェック***/
                'Check個人未登録(sCSV論理, pRow.Item("所有者ID"), "農地", 1) 'エラーコード１
                'Check個人未登録(sCSV論理, .Fnc耕作者整理番号(pRow, TBL公開用個人, 1), "農地", 2) 'エラーコード２
                'Check耕作者番号(sCSV論理, .Fnc耕作者整理番号(pRow, TBL公開用個人, 1), 1) 'エラーコード５
                'Check耕作者番号(sCSV論理, .Fnc耕作者整理番号(pRow, TBL公開用個人, 2), 2) 'エラーコード５
            End With

            RowCount += 1
        Next

        Select Case pValue
            Case EnumOutPutType.全件
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用農地", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")), True)
            Case EnumOutPutType.農地
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用農地", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")), True, True)
        End Select

        '名前を付けて保存(sCSV論理, String.Format("{0}_{1}_取込用農地論理チェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
        '名前を付けて保存(sCSVレイアウト, String.Format("{0}_{1}_取込用農地レイアウトチェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
    End Sub
    Private Sub Set取込用個人(ByVal pValue As Integer)
        Dim pLineHeader As New StringBEx("世帯員番号", EnumCnv.設定無)
        Dim pLineHeader論理 As New StringBEx("連番", EnumCnv.設定無)
        Dim pLineHeaderレイアウト As New StringBEx("連番", EnumCnv.設定無)
        Dim h取込用個人 As String() = GetHeader個人()
        Dim h取込用論理 As String() = GetHeader論理()
        Dim h取込用レイアウト As String() = GetHeaderレイアウト()
        Dim pView As DataView = New DataView(TBL個人, "", "[ID]", DataViewRowState.CurrentRows)   '20161114追加(テスト)

        Dim TBL続柄 As DataTable = App農地基本台帳.DataMaster.GetClassTable("続柄")
        TBL続柄.PrimaryKey = New DataColumn() {TBL続柄.Columns("ID")}

        Dim TBL住民区分 As DataTable = App農地基本台帳.DataMaster.GetClassTable("住民区分")
        TBL住民区分.PrimaryKey = New DataColumn() {TBL住民区分.Columns("ID")}

        Initialization()

        For n As Integer = 0 To UBound(h取込用個人)
            pLineHeader.mvarBody.Append("," & h取込用個人(n))
        Next
        sCSV.Body.AppendLine(pLineHeader.Body.ToString)

        For n = 0 To UBound(h取込用論理)
            pLineHeader論理.mvarBody.Append("," & h取込用論理(n))
        Next
        sCSV論理.Body.AppendLine(pLineHeader論理.Body.ToString)

        For n = 0 To UBound(h取込用レイアウト)
            pLineHeaderレイアウト.mvarBody.Append("," & h取込用レイアウト(n))
        Next
        sCSVレイアウト.Body.AppendLine(pLineHeaderレイアウト.Body.ToString)

        Me.Maximum = pView.Count
        Me.Value = 0

        For Each pRow As DataRowView In pView
            Me.Value += 1
            Message = "取込用個人データ出力中(" & Me.Value & "/" & pView.Count & ")..."

            Dim pLineRow As New StringBEx(CnvID(pRow.Item("ID")), EnumCnv.半角, "1:世帯員番号:13")
            With pLineRow
                .CnvData(CnvID(pRow.Item("世帯ID")), EnumCnv.半角, , "2:世帯コード:13")
                .CnvData(CnvID(pRow.Item("世帯ID")), EnumCnv.半角, , "3:住基の世帯コード:13")
                .CnvData(pRow.Item("氏名"), EnumCnv.全角, , "4:氏名又は名称:100")
                .CnvData(IIf(IsDBNull(pRow.Item("フリガナ")), "フメイ", pRow.Item("フリガナ").ToString), EnumCnv.全角, , "5:フリガナ:100")
                .CnvData(.Cnv性別(pRow.Item("性別")), EnumCnv.選択, , "6:性別コード:1")
                .CnvData(.Cnv続柄(TBL続柄, pRow.Item("続柄1")), EnumCnv.半角, , "7:続柄1:2")
                .CnvData(.Cnv続柄(TBL続柄, pRow.Item("続柄2")), EnumCnv.半角, , "8:続柄2:2")
                .CnvData(.Cnv続柄(TBL続柄, pRow.Item("続柄3")), EnumCnv.半角, , "9:続柄3:2")
                .CnvData(.Cnv続柄(TBL続柄, pRow.Item("続柄4")), EnumCnv.半角, , "10:続柄4:2")
                .CnvData(pRow.Item("郵便番号"), EnumCnv.半角, , "11:郵便番号:8")
                Dim Conv市町村CD As String = 市町村コード(pRow.Item("住所").ToString)
                .CnvData(IIf(Len(pRow.Item("市町村ID").ToString) = 6, pRow.Item("市町村ID").ToString, IIf(Conv市町村CD = "", Find市町村コード(pRow.Item("住所").ToString), Conv市町村CD)), EnumCnv.半角, , "12:市町村コード:6")
                .mvarBody.Append(",") '大字コード
                .mvarBody.Append(",") '大字名
                .mvarBody.Append(",") '小字コード
                .mvarBody.Append(",") '小字名
                .CnvData(pRow.Item("住所"), EnumCnv.全角, , "17:住所:60")
                .CnvData(pRow.Item("電話番号"), EnumCnv.半角, , "18:電話:13")
                .CnvData(pRow.Item("FAX番号"), EnumCnv.半角, , "19:FAX:13")
                .CnvData(pRow.Item("メールアドレス"), EnumCnv.半角, , "20:EMAIL:80")
                .CnvData(pRow.Item("生年月日"), EnumCnv.日付, , "21:生年月日:10")
                .CnvData(.Cnv住民区分(TBL住民区分, pRow.Item("住民区分")), EnumCnv.半角, , "22:住民区分:1")
                .CnvData(pRow.Item("異動区分"), EnumCnv.選択, , "23:異動区分:1")
                .CnvData(pRow.Item("住記異動日"), EnumCnv.日付, , "24:異動年月日:10")
                .CnvData(Val(pRow.Item("注意区分").ToString), EnumCnv.選択, , "25:注意区分:1")
                .CnvData(IIf(pRow.Item("世帯責任者") = True, 1, 0), EnumCnv.選択, , "26:世帯責任者:1")
                .CnvData(IIf(pRow.Item("農業経営者") = True, 1, 0), EnumCnv.選択, , "27:農業経営主:1")
                .CnvData(IIf(pRow.Item("農業跡継ぎ") = True, 1, 0), EnumCnv.選択, , "28:農業あとつぎ:1")
                .CnvData(pRow.Item("担い手等の区分"), EnumCnv.選択, , "29:担い手等の区分:2")
                .CnvData(pRow.Item("認定日"), EnumCnv.日付, , "30:認定農業者における認定年月日:10")
                .CnvData(pRow.Item("新規就農者認定日"), EnumCnv.日付, , "31:認定新規就農者における認定年月日:10")
                .CnvData(Val(pRow.Item("あっせん候補者区分").ToString), EnumCnv.選択, 2, "32:農地移動適正化あっせん事業候補者:1")
                .CnvData(pRow.Item("あっせん登録日"), EnumCnv.日付, , "33:あっせん登録年月日:10")
                .CnvData(pRow.Item("あっせん登録番号"), EnumCnv.半角, , "34:あっせん登録番号:20")
                .CnvData(pRow.Item("農業従事日数"), EnumCnv.日付, , "35:年間農業従事日数:3")
                .CnvData(Val(pRow.Item("自家農業従事程度").ToString), EnumCnv.選択, 5, "36:自家農業従事程度:1")
                .CnvData(Val(pRow.Item("兼業形態").ToString), EnumCnv.選択, 6, "37:兼業の形態:1")
                .CnvData(pRow.Item("職業"), EnumCnv.全角, , "38:就労または就学先:30")
                .CnvData(Val(pRow.Item("在留資格").ToString), EnumCnv.選択, 6) '在留資格

                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1, 1, 0), EnumCnv.選択) '旧制度加入者
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 2, 1, 0), EnumCnv.選択) '旧制度受給者
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("国民年金加入種別"), ""), EnumCnv.選択, 4) '国民年金加入種別
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("農年加入種別"), ""), EnumCnv.選択, 4) '農業者年金加入種別
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("被保険者番号"), ""), EnumCnv.半角) '農業者年金被保険者番号
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("農年受給者番号"), ""), EnumCnv.半角) '農業者年金受給者番号
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("資格取得年月日"), ""), EnumCnv.日付) '取得年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("資格喪失年月日"), ""), EnumCnv.日付) '喪失年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("農年受給日"), ""), EnumCnv.日付) '農年受給日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("経営移譲種別"), ""), EnumCnv.選択, 5) '経営移譲種別
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("経営移譲終了日"), ""), EnumCnv.日付) '移譲終了年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("経営移譲裁定日"), ""), EnumCnv.日付) '移譲裁定年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("老齢裁定日"), ""), EnumCnv.日付) '老年裁定年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("老齢受給の有無"), ""), EnumCnv.選択) '老年加算の有無
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("一時給付金の有無"), ""), EnumCnv.選択) '一時給付金の有無
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 1 Or Val(pRow.Item("農年加入受給種別").ToString) = 2, pRow.Item("その他年金種別"), ""), EnumCnv.選択, 3) 'その他年金種別

                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 3, 1, 0), EnumCnv.選択) '新制度加入者
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) = 4, 1, 0), EnumCnv.選択) '新制度受給者
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度年金種別"), ""), EnumCnv.選択, 3) '年金の種類
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度変更前種別"), ""), EnumCnv.選択, 3) '変更前の種類
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度変更日"), ""), EnumCnv.日付) '変更年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度政策支援加入区分"), ""), EnumCnv.選択, 7) '政策支援加入区分
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度変更前政策支援加入区分"), ""), EnumCnv.選択, 7) '変更前政策支援加入区分
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度政策支援認定日"), ""), EnumCnv.日付) '政策支援認定年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度被保険者記号番号"), ""), EnumCnv.半角) '新制度被保険者記号番号
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("経営移譲種別"), ""), EnumCnv.選択, 5) '継承種別
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("経営移譲終了日"), ""), EnumCnv.日付) '継承終了年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("経営移譲裁定日"), ""), EnumCnv.日付) '継承裁定年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("老齢裁定日"), ""), EnumCnv.日付) '老年裁定年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("資格取得年月日"), ""), EnumCnv.日付) '資格取得年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度資格停止日"), ""), EnumCnv.日付) '資格停止年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("資格喪失年月日"), ""), EnumCnv.日付) '資格喪失年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("農年受給日"), ""), EnumCnv.日付) '受給年月日
                .CnvData(IIf(Val(pRow.Item("農年加入受給種別").ToString) > 2, pRow.Item("新制度死亡一時金の有無"), ""), EnumCnv.選択) '死亡一時金の有無
                '↓↓↓必要？
                .CnvData(IIf(pRow.Item("選挙権の有無").ToString = True, 1, 0), EnumCnv.選択) '選挙権有無
                .mvarBody.Append(",") '登録年月日
                .mvarBody.Append(",") '抹消年月日
                .mvarBody.Append(",") '選挙区コード
                .mvarBody.Append(",") '選挙区名称
                .mvarBody.Append(",") '投票区コード
                .mvarBody.Append(",") '投票区名称

                .CnvData(pRow.Item("備考"), EnumCnv.全角) '備考

                sCSV.Body.AppendLine(.Body.ToString)

                '/***論理チェック***/
                'Check個人未登録(sCSV論理, pRow.Item("ID"), "個人", 1) 'エラーコード１
                'Check個人未登録(sCSV論理, pRow.Item("ID"), "個人", 2) 'エラーコード２
                'Check個人未登録(sCSV論理, pRow.Item("ID"), "個人", 3) 'エラーコード３
                'Check世帯未登録(sCSV論理, pRow.Item("世帯ID"), "個人") 'エラーコード４
            End With

            RowCount += 1
        Next

        Select Case pValue
            Case EnumOutPutType.全件
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用個人", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
            Case EnumOutPutType.個人
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用個人", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")), True, True)
        End Select

        '名前を付けて保存(sCSV論理, String.Format("{0}_{1}_取込用個人論理チェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
        '名前を付けて保存(sCSVレイアウト, String.Format("{0}_{1}_取込用個人レイアウトチェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
    End Sub
    Private Sub Set取込用世帯・法人(ByVal pValue As Integer)
        Dim pLineHeader As New StringBEx("世帯コード", EnumCnv.設定無)
        Dim pLineHeader論理 As New StringBEx("連番", EnumCnv.設定無)
        Dim h取込用世帯・法人 As String() = GetHeader世帯・法人()
        Dim h取込用論理 As String() = GetHeader論理()
        Dim pView As DataView = New DataView(TBL世帯, "", "[ID]", DataViewRowState.CurrentRows)   '20161114追加(テスト)

        Initialization()

        For n As Integer = 0 To UBound(h取込用世帯・法人)
            pLineHeader.mvarBody.Append("," & h取込用世帯・法人(n))
        Next
        sCSV.Body.AppendLine(pLineHeader.Body.ToString)

        For n = 0 To UBound(h取込用論理)
            pLineHeader論理.mvarBody.Append("," & h取込用論理(n))
        Next
        sCSV論理.Body.AppendLine(pLineHeader論理.Body.ToString)

        Me.Maximum = pView.Count
        Me.Value = 0

        For Each pRow As DataRowView In pView
            Me.Value += 1
            Message = "取込用世帯・法人データ出力中(" & Me.Value & "/" & pView.Count & ")..."

            'Debug.Print(pRow.Item("ID"))
            'If pRow.Item("ID") = 3675 Then
            '    Stop
            'End If

            Dim pLineRow As New StringBEx(CnvID(pRow.Item("ID")), EnumCnv.半角, "01:世帯コード:13")
            With pLineRow
                If IsDBNull(pRow.Item("農地所有区分")) Then : .Cnv農地所有区分(Val(pRow.Item("世帯主ID").ToString)) '農地所有区分コード
                Else : .CnvData(pRow.Item("農地所有区分"), EnumCnv.選択) '農地所有区分コード
                End If
                .CnvData(CnvID(pRow.Item("世帯主ID")), EnumCnv.半角) '経営者世帯員番号
                .CnvData(.Find農家情報(Val(pRow.Item("世帯主ID").ToString), Enum農家.郵便番号), EnumCnv.半角) '郵便番号
                Dim Conv市町村CD As String = 市町村コード(pRow.Item("住所").ToString)
                .CnvData(IIf(Len(pRow.Item("市町村ID").ToString) = 6, pRow.Item("市町村ID").ToString, IIf(Conv市町村CD = "", Find市町村コード(pRow.Item("住所").ToString), Conv市町村CD)), EnumCnv.半角) '市町村コード
                .mvarBody.Append(",") '大字コード
                .mvarBody.Append(",") '大字名
                .mvarBody.Append(",") '小字コード
                .mvarBody.Append(",") '小字名
                .CnvData(.Find農家情報(Val(pRow.Item("世帯主ID").ToString), Enum農家.住所), EnumCnv.全角) '住所
                .CnvData(pRow.Item("支店等住所"), EnumCnv.全角) '支店等住所
                .Find世帯情報(pRow.Item("世帯主ID"), "世帯") '電話・FAX・EMAIL  '20161114追加(テスト)
                .CnvData(pRow.Item("就業状況"), EnumCnv.選択) '就業状況
                .CnvData(pRow.Item("農事組合ID"), EnumCnv.半角) '農事組合コード
                .CnvData(.Find農家情報(Val(pRow.Item("農事組合ID").ToString), Enum農家.氏名), EnumCnv.全角) '農事組合名称   '20161114追加(テスト)
                .CnvData(pRow.Item("所属農協ID"), EnumCnv.半角) '所属農協コード
                .CnvData(.Find農家情報(Val(pRow.Item("所属農協ID").ToString), Enum農家.氏名), EnumCnv.全角) '所属農協名称   '20161114追加(テスト)
                .Find世帯情報(pRow.Item("世帯主ID"), "その他") '担い手等の区分・認定農業者等認定年月日・認定新規就農者における認定年月日 '20161114追加(テスト)
                .CnvData(pRow.Item("認定時面積"), EnumCnv.面積) '認定時面積
                .CnvData(pRow.Item("人農地プラン中心経営体区分"), EnumCnv.選択, 3) '人・農地プランにおける中心経営体かどうか
                .CnvData(pRow.Item("農地移動適正化あっせん事業"), EnumCnv.選択, 2) '農地移動適正化あっせん事業
                .CnvData(pRow.Item("あっせん登録日"), EnumCnv.日付) 'あっせん登録年月日
                .CnvData(pRow.Item("あっせん登録時面積"), EnumCnv.面積) 'あっせん登録時面積
                .CnvData(pRow.Item("トラクター台数"), EnumCnv.半角) 'トラクター
                .CnvData(pRow.Item("耕運機台数"), EnumCnv.半角) '耕運機
                .CnvData(pRow.Item("田植機台数"), EnumCnv.半角) '田植機
                .CnvData(pRow.Item("コンバイン台数"), EnumCnv.半角) 'コンバイン
                .CnvData(pRow.Item("乾燥機台数"), EnumCnv.半角) '乾燥機
                .CnvData(pRow.Item("噴霧器台数"), EnumCnv.半角) '噴霧器
                .CnvData(pRow.Item("その他機具台数"), EnumCnv.半角) 'その他機具
                .CnvData(pRow.Item("畜舎規模"), EnumCnv.面積) '畜舎規模
                .CnvData(pRow.Item("畜舎数"), EnumCnv.半角) '畜舎数
                .CnvData(pRow.Item("温室規模"), EnumCnv.面積) '温室規模
                .CnvData(pRow.Item("温室数"), EnumCnv.半角) '温室数
                .CnvData(pRow.Item("その他施設規模"), EnumCnv.面積) 'その他施設規模
                .CnvData(pRow.Item("その他施設数"), EnumCnv.半角) 'その他施設数
                .CnvData(pRow.Item("販売収入1位"), EnumCnv.選択) '1位
                .CnvData(pRow.Item("販売収入2位"), EnumCnv.選択) '2位
                .CnvData(pRow.Item("販売収入3位"), EnumCnv.選択) '3位
                .CnvData(pRow.Item("主要作物1"), EnumCnv.選択) '作目１
                .CnvData(pRow.Item("主要作物規模1"), EnumCnv.面積) '作目規模１
                .CnvData(pRow.Item("主要作物2"), EnumCnv.選択) '作目２
                .CnvData(pRow.Item("主要作物規模2"), EnumCnv.面積) '作目規模２
                .CnvData(pRow.Item("主要作物3"), EnumCnv.選択) '作目３
                .CnvData(pRow.Item("主要作物規模3"), EnumCnv.面積) '作目規模３
                .CnvData(pRow.Item("主要作物4"), EnumCnv.選択) '作目４
                .CnvData(pRow.Item("主要作物規模4"), EnumCnv.面積) '作目規模４
                .CnvData(pRow.Item("主要作物5"), EnumCnv.選択) '作目５
                .CnvData(pRow.Item("主要作物規模5"), EnumCnv.面積) '作目規模５
                .CnvData(pRow.Item("肉用牛頭数"), EnumCnv.半角) '肉用牛
                .CnvData(pRow.Item("乳牛頭数"), EnumCnv.半角) '乳牛
                .CnvData(pRow.Item("豚頭数"), EnumCnv.半角) '豚
                .CnvData(pRow.Item("採卵用鶏羽数"), EnumCnv.半角) '採卵用鶏
                .CnvData(pRow.Item("ブロイラー羽数"), EnumCnv.半角) 'ブロイラー
                .CnvData(pRow.Item("その他家畜頭数"), EnumCnv.半角) 'その他家畜
                .CnvData(pRow.Item("青色申告"), EnumCnv.選択, 3) '申告納税方式
                .CnvData(pRow.Item("制度資金種別1"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦1"), EnumCnv.半角) '年次（西暦）
                .CnvData(pRow.Item("制度資金種別2"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦2"), EnumCnv.半角) '年次（西暦）
                .CnvData(pRow.Item("制度資金種別3"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦3"), EnumCnv.半角) '年次（西暦）
                .CnvData(pRow.Item("制度資金種別4"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦4"), EnumCnv.半角) '年次（西暦）
                .CnvData(pRow.Item("制度資金種別5"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦5"), EnumCnv.半角) '年次（西暦）
                .CnvData(pRow.Item("制度資金種別6"), EnumCnv.全角) '種類
                .CnvData(pRow.Item("制度資金西暦6"), EnumCnv.半角) '年次（西暦）
                .Set経営規模(pRow.Item("ID")) '経営規模
                .CnvData(pRow.Item("農家分類専業形態"), EnumCnv.選択) '専兼形態
                .CnvData(pRow.Item("農家分類あとつぎ"), EnumCnv.選択) 'あとつぎ
                .CnvData(pRow.Item("農家分類規模拡大志向"), EnumCnv.選択) '規模拡大志向
                .CnvData(pRow.Item("団地数"), EnumCnv.半角) '筆数
                .CnvData(pRow.Item("基盤整備実施済筆数"), EnumCnv.半角) '筆数
                .CnvData(pRow.Item("基盤整備実施済面積"), EnumCnv.面積) '面積
                .CnvData(pRow.Item("転作筆数"), EnumCnv.半角) '筆数
                .CnvData(pRow.Item("転作面積"), EnumCnv.面積) '面積
                .CnvData(pRow.Item("裏作利用筆数"), EnumCnv.半角) '筆数
                .CnvData(pRow.Item("裏作利用面積"), EnumCnv.面積) '面積
                .CnvData(pRow.Item("経営意向等調査日"), EnumCnv.日付) '経営意向等調査年月日
                .CnvData(pRow.Item("経営意向等農業志向"), EnumCnv.選択, 6) '農業志向
                .CnvData(pRow.Item("経営意向等経営計画"), EnumCnv.選択, 5) '経営計画
                .CnvData(pRow.Item("経営部門1"), EnumCnv.選択) '部門１
                .CnvData(pRow.Item("経営部門1拡大縮小"), EnumCnv.選択) '部門１拡大・縮小
                .CnvData(pRow.Item("経営部門1拡大縮小方法"), EnumCnv.選択) '部門１拡大・縮小方法
                .CnvData(pRow.Item("経営部門1拡大縮小面積"), EnumCnv.面積) '拡大・縮小面積１
                .CnvData(pRow.Item("経営部門2"), EnumCnv.選択) '部門２
                .CnvData(pRow.Item("経営部門2拡大縮小"), EnumCnv.選択) '部門２拡大・縮小
                .CnvData(pRow.Item("経営部門2拡大縮小方法"), EnumCnv.選択) '部門２拡大・縮小方法
                .CnvData(pRow.Item("経営部門2拡大縮小面積"), EnumCnv.面積) '拡大・縮小面積２
                .CnvData(pRow.Item("経営部門3"), EnumCnv.選択) '部門３
                .CnvData(pRow.Item("経営部門3拡大縮小"), EnumCnv.選択) '部門３拡大・縮小
                .CnvData(pRow.Item("経営部門3拡大縮小方法"), EnumCnv.選択) '部門３拡大・縮小方法
                .CnvData(pRow.Item("経営部門3拡大縮小面積"), EnumCnv.面積) '拡大・縮小面積３
                .CnvData(IIf(IsDBNull(pRow.Item("農用地改善団体参加")), 0, IIf(Val(pRow.Item("農用地改善団体参加").ToString) = 1, 2, 1)), EnumCnv.選択) '農用地利用改善団体
                .CnvData(IIf(IsDBNull(pRow.Item("地域農業集団構成員")), 0, IIf(Val(pRow.Item("地域農業集団構成員").ToString) = 1, 1, 2)), EnumCnv.選択) '地域農業集団
                .CnvData(pRow.Item("法人格"), EnumCnv.選択) '法人格
                .CnvData(pRow.Item("法人格設立日"), EnumCnv.日付) '設立年月日
                .CnvData(pRow.Item("法人格初回許可日"), EnumCnv.日付) '最初の許可年月日
                .CnvData(pRow.Item("備考"), EnumCnv.全角) '備考

                sCSV.Body.AppendLine(.Body.ToString)

                '/***論理チェック***/
                'Check個人未登録(sCSV論理, pRow.Item("世帯主ID"), "世帯・法人", 3) 'エラーコード３
                'Check世帯未登録(sCSV論理, pRow.Item("世帯主ID"), "世帯・法人") 'エラーコード４
            End With


            RowCount += 1
        Next

        Select Case pValue
            Case EnumOutPutType.全件
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用世帯・法人", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")), , True)
            Case EnumOutPutType.世帯
                名前を付けて保存(sCSV, String.Format("{0}_{1}_取込用世帯・法人", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")), True, True)
        End Select


        '名前を付けて保存(sCSV論理, String.Format("{0}_{1}_取込用世帯・法人論理チェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
        '名前を付けて保存(sCSVレイアウト, String.Format("{0}_{1}_取込用世帯・法人レイアウトチェック", 都道府県ID & 市町村CD, Format(Now, "yyyyMMdd")))
    End Sub
    Public Sub Set市町村コード()
        If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
            Dim cReader As New System.IO.StreamReader(SysAD.CustomReportFolder("共通様式") & "\code_list.csv", System.Text.Encoding.Default)
            Dim LoopCount As Integer = 0

            While (cReader.Peek() >= 0)
                Dim stBuffer As String = cReader.ReadLine() ' ファイルを 1 行ずつ読み込む
                Dim cAr As Object = Split(stBuffer, ",")

                If LoopCount = 0 Then
                    With TBL市町村コード
                        .Columns.Add(cAr(0))
                        .Columns.Add(cAr(1))
                        .Columns.Add(cAr(2))
                        .Columns.Add(cAr(3))
                        .Columns.Add(cAr(4))
                    End With

                    LoopCount += 1
                Else
                    Dim pRow As DataRow = TBL市町村コード.NewRow
                    pRow.Item("団体コード") = cAr(0)
                    pRow.Item("都道府県名（漢字）") = cAr(1)
                    pRow.Item("市区町村名（漢字）") = cAr(2)
                    pRow.Item("都道府県名（カナ）") = cAr(3)
                    pRow.Item("市区町村名（カナ）") = cAr(4)

                    TBL市町村コード.Rows.Add(pRow)
                End If
            End While
        End If
    End Sub

    Private Sub Initialization()
        sCSV = New StringBEx("", EnumCnv.設定無)
        sCSV論理 = New StringBEx("", EnumCnv.設定無)
        sCSVレイアウト = New StringBEx("", EnumCnv.設定無)
        論理Flg = False
        論理連番 = 1
        レイアウトFlg = False
        レイアウト連番 = 1
        RowCount = 1
    End Sub

    Private Function GetHeader農地()
        Dim sResult As String() = {"大字コード", "大字名", "小字コード", "小字名", "本番区分", "本番", "枝番区分", "枝番", "孫番区分", "孫番", "曾孫番区分", "曾孫番", _
                                        "玄孫番区分", "玄孫番", "区分", "耕地番号", "耕地番号の作成年月日（提供元の情報時点）", "登記簿地目", "現況地目", "登記簿面積", "登記簿面積の内訳", _
                                        "現況面積", "本地面積", "本地面積の作成年月日（提供元の情報時点）", "農振法区分", "都市計画法区分", "生産緑地法に基づく指定", "生産緑地法に基づく指定（指定年月日）", "生産緑地法に基づく指定（指定面積）", "生産緑地法に基づく指定（種別）", _
                                        "農地種別", "所有者世帯員番号", "所有者の農地に関する意向", "所有者の農地に関する意向「その他」の内訳", "農法第52条の3第1項による公表への同意", "耕作者世帯員番号", _
                                        "耕作者整理番号", "耕作状況", "作業者世帯員番号", "作物", "作業内容", "許可年月日", "適用法", "権利の種類", "始期年月日", "終期年月日", "再設定した場合の終期年月日", _
                                        "1年間の借賃額", "10a当たりの借賃額", "物納", "物納単位", "利用集積計画番号", "利用集積公告年月日", "利用目的", "利用目的備考", "利用権設定区分", "交付金判定", "交付金対象額", _
                                        "借受人世帯員番号", "転貸適用法", "転貸権利の種類", "転貸始期年月日", "転貸終期年月日", "転貸1年間の借賃額", "転貸10a当たりの借賃額", "物納", "物納単位", _
                                        "機構が農地中間管理権を取得した年月日", "利用配分計画案への意見回答年月日", "利用配分計画の知事公告年月日", "計画の認可通知年月日", "農地中間管理事業法20条に基づく貸借解約年月日", _
                                        "納税猶予", "種別", "相続日・贈与日", "適用年月日", "継続年月日", "確認年月日", "特定貸付け根拠条項（租税特別措置法第70条の4の2第1項又は第70条の6の2第1項）", _
                                        "営農困難時貸付け", "調査実施年月日", "農地法第32条第1項第1号", "荒廃農地調査分類", "利用状況", "調査委員名", "耕作放棄地通し番号", "一時転用", "無断転用", "違反転用", _
                                        "調査不可と判断した年月日", "理由", "理由「その他」の内訳", "調査実施年月日", "根拠条項", "所有者意思表明年月日", "調査結果", "調査結果のその他任意文字", _
                                        "措置の実施状況", "権利関係調査", "調査結果年月日", "調査結果", "調査結果のその他任意文字", "農地法第32条第3項に基づく公示年月日", "農地法第43条第1項に基づく農地中間管理機構への通知発出年月日", _
                                        "農地法第35条第1項に基づく通知発出年月日", "農地法第35条第2項に基づく協議を行わない通知発出年月日", "農地法第35条第2項に基づく協議を機構が所有者に申し入れた年月日", "農地法第35条第3項に基づく通知発出年月日", _
                                        "勧告年月日", "勧告内容", "農地中間管理機構等への通知発出年月日", "再生利用困難な農地", "農地法第40条に基づく裁定公告日", "農地法第43条に基づく裁定公告日", _
                                        "農地法第44条第1項に基づく命令年月日", "農地法第44条第3項に基づく公告年月日", "利用状況報告の対象", "利用状況報告年月日", "勧告年月日", "内容", "期限", _
                                        "根拠条項", "確認年月日", "是正状況", "取消年月日", "取消事由", "根拠条項", "届出年月日", "届出事由", "権利取得者世帯員番号", "相続登記の有無", "設定年月日", _
                                        "仮登記権者世帯員番号", "環境保全型農業直接支払交付金", "交付金対象の基準年月日（提供元の情報時点）", "農地維持支払交付金", "交付金対象の基準年月日（提供元の情報時点）", _
                                        "資源向上支払い交付金", "交付金対象の基準年月日（提供元の情報時点）", "中山間地域等直接支払", "交付金対象の基準年月日（提供元の情報時点）", "特定処分対象農地", _
                                        "農業者年金処分対象農地", "農業者年金処分適用年月日", "転用適用法", "転用形態", "転用用途", "転用換地有無", "始期年月日", "終期年月日", "圃場整備", "区画整理", _
                                        "前払いの有無", "指定の有無", "耕作しているであろう人の世帯員番号", "備考"}
        Return sResult
    End Function
    Private Function GetHeader個人()
        Dim sResult As String() = {"世帯コード", "住基の世帯コード", "氏名又は名称", "フリガナ", "性別コード", "続柄1", "続柄2", "続柄3", "続柄4", "郵便番号", "市町村コード", "大字コード", "大字名", _
                                   "小字コード", "小字名", "住所", "電話", "FAX", "EMAIL", "生年月日", "住民区分", "異動区分", "異動年月日", "注意区分", "世帯責任者", "農業経営主", "農業あとつぎ", _
                                   "担い手等の区分", "認定農業者における認定年月日", "認定新規就農者における認定年月日", "農地移動適正化あっせん事業候補者", "あっせん登録年月日", "あっせん登録番号", _
                                   "年間農業従事日数", "自家農業従事程度", "兼業の形態", "就労または就学先", "在留資格", "旧制度加入者", "旧制度受給者", "国民年金加入種別", "農業者年金加入種別", _
                                   "農業者年金被保険者番号", "農業者年金受給者番号", "取得年月日", "喪失年月日", "受給年月日", "経営移譲種別", "移譲終了年月日", "移譲裁定年月日", "老年裁定年月日", _
                                   "老年加算の有無", "一時給付金の有無", "その他年金種別", "新制度加入者", "新制度受給者", "年金の種類", "変更前の種類", "変更年月日", "政策支援加入区分", _
                                   "変更前政策支援加入区分", "政策支援認定年月日", "被保険者記号番号", "継承種別", "継承終了年月日", "継承裁定年月日", "老年裁定年月日", "資格取得年月日", _
                                   "資格停止年月日", "資格喪失年月日", "受給年月日", "死亡一時金の有無", "選挙権有無", "登録年月日", "抹消年月日", "選挙区コード", "選挙区名称", "投票区コード", "投票区名称", "備考"}
        Return sResult
    End Function
    Private Function GetHeader世帯・法人()
        Dim sResult As String() = {"農地所有者区分コード", "経営者世帯員番号", "郵便番号", "市町村コード", "大字コード", "大字名", "小字コード", "小字名", _
                                            "住所", "支店等住所", "電話", "FAX", "EMAIL", "就業状況", "農事組合コード", "農事組合名称", "所属農協コード", "所属農協名称", "担い手等の区分", _
                                            "認定農業者等認定年月日", "認定新規就農者における認定年月日", "認定時面積", "人・農地プランにおける中心経営体かどうか", "農地移動適正化あっせん事業", _
                                            "あっせん登録年月日", "あっせん登録時面積", "トラクター", "耕運機", "田植機", "コンバイン", "乾燥機", "噴霧器", "その他機具", "畜舎規模", "畜舎数", _
                                            "温室規模", "温室数", "その他施設規模", "その他施設数", "1位", "2位", "3位", "作目1", "作目規模1", "作物2", "作目規模2", "作目3", "作目規模3", _
                                            "作目4", "作目規模4", "作目5", "作目規模5", "肉用牛", "乳牛", "豚", "採卵用鶏", "ブロイラー", "その他家畜", "申告納税方式", "種類", "年次（西暦）", _
                                            "種類", "年次（西暦）", "種類", "年次（西暦）", "種類", "年次（西暦）", "種類", "年次（西暦）", "種類", "年次（西暦）", "経営規模", "兼業形態", _
                                            "あとつぎ", "規模拡大志向", "筆数", "筆数", "面積", "筆数", "面積", "筆数", "面積", "経営意向等調査年月日", "農業志向", "経営計画", _
                                            "部門1", "部門1拡大・縮小", "部門1拡大・縮小方法", "拡大・縮小面積1", "部門2", "部門2拡大・縮小", "部門2拡大・縮小方法", "拡大・縮小面積2", _
                                            "部門3", "部門3拡大・縮小", "部門3拡大・縮小方法", "拡大・縮小面積3", "農用地利用改善団体", "地域農業集団", "法人格", "設立年月日", _
                                            "最初の許可年月日", "備考"}
        Return sResult
    End Function
    Private Function GetHeader論理()
        Dim sResult As String() = {"データ種類", "行番号", "エラー内容", "エラーコード", "値"}
        Return sResult
    End Function
    Private Function GetHeaderレイアウト()
        Dim sResult As String() = {"行番号", "列番号", "項目名", "項目内容", "エラー内容", "エラーコード"}
        Return sResult
    End Function

    '/***論理チェック用***/
    Private Sub Check個人未登録(ByRef sCSV論理 As StringBEx, ByVal pID As Object, ByVal DataType As String, ByVal pECode As Integer)
        Dim 論理Flg As Boolean = False
        If pID Is Nothing Or IsDBNull(pID) Then
            pID = ""
            論理Flg = True
        Else
            Dim pRow As DataRow = TBL個人.Rows.Find(pID)
            If pID = 0 Or pRow Is Nothing Then
                論理Flg = True
            End If
        End If

        If 論理Flg = True Then
            Dim pLineRow As New StringBEx(論理連番, EnumCnv.設定無) ' 連番
            pLineRow.Body.Append("," & DataType) ' データ種類
            pLineRow.Body.Append("," & Me.Value) ' 行番号
            Select Case pECode ' エラー内容
                Case 1
                    pLineRow.Body.Append(",個人未登録(所有者)")
                Case 2
                    pLineRow.Body.Append(",個人未登録(耕作者)")
                Case 3
                    pLineRow.Body.Append(",個人未登録(世帯・法人)")
            End Select
            pLineRow.Body.Append("," & pECode) ' エラーコード
            pLineRow.Body.Append("," & pID) ' 値
            sCSV論理.Body.AppendLine(pLineRow.Body.ToString)

            論理Flg = False
            論理連番 += 1
        End If
    End Sub
    Private Sub Check世帯未登録(ByRef sCSV論理 As StringBEx, ByVal pID As Object, ByVal DataType As String)
        Dim 論理Flg As Boolean = False
        If pID Is Nothing Or IsDBNull(pID) Then
            pID = ""
            論理Flg = True
        Else
            Dim pRow As DataRow = TBL世帯.Rows.Find(pID)
            If pID = 0 Or pRow Is Nothing Then
                論理Flg = True
            End If
        End If

        If 論理Flg = True Then
            Dim pLineRow As New StringBEx(論理連番, EnumCnv.設定無) ' 連番
            pLineRow.Body.Append("," & DataType) ' データ種類
            pLineRow.Body.Append("," & Me.Value) ' 行番号
            pLineRow.Body.Append(",世帯未登録") ' エラー内容
            pLineRow.Body.Append(",4") ' エラーコード
            pLineRow.Body.Append("," & pID) ' 値
            sCSV論理.Body.AppendLine(pLineRow.Body.ToString)

            論理Flg = False
            論理連番 += 1
        End If
    End Sub
    Private Sub Check耕作者番号(ByRef sCSV論理 As StringBEx, ByVal pID As Decimal, ByVal pType As Integer)
        Dim pTBL As DataTable = Nothing

        Select Case pType
            Case 1
                pTBL = New DataView(TBL耕作者, "[PID] = " & pID, "", DataViewRowState.CurrentRows).ToTable
            Case 2
                pTBL = New DataView(TBL耕作者, "[AutoID] = " & pID, "", DataViewRowState.CurrentRows).ToTable
        End Select

        If pTBL.Rows.Count > 1 Then
            Dim pLineRow As New StringBEx(論理連番, EnumCnv.設定無) ' 連番
            pLineRow.Body.Append(",農地") ' データ種類
            pLineRow.Body.Append("," & Me.Value) ' 行番号
            pLineRow.Body.Append(",耕作者世帯員番号・整理番号不整合") ' エラー内容
            pLineRow.Body.Append(",5") ' エラーコード
            pLineRow.Body.Append("," & pID) ' 値
            sCSV論理.Body.AppendLine(pLineRow.Body.ToString)
            論理Flg = True

            論理連番 += 1
        End If
    End Sub
End Class

Public Class CF2ErrorCheck
    Inherits HIMTools.Controls.CTabPageWithToolStrip

    Public WithEvents mvarTabCtrl As HIMTools.Controls.TabControlBase
    Private mvarTab農地 As CTabPage農地
    Private mvarTab個人 As CTabPage個人
    Private mvarTab世帯 As CTabPage世帯

    Public Sub New()
        MyBase.New(True, False, "フェーズ２エラーチェック", "フェーズ２エラーチェック")
        mvarTabCtrl = New HIMTools.Controls.TabControlBase()
        mvarTabCtrl.Dock = DockStyle.Fill
        Me.ControlPanel.Add(mvarTabCtrl)

        Me.SuspendLayout()
        mvarTab農地 = New CTabPage農地("農地", "")
        mvarTabCtrl.TabPages.Add(mvarTab農地)

        mvarTab個人 = New CTabPage個人("個人", "")
        mvarTabCtrl.AddPage(mvarTab個人)

        mvarTab世帯 = New CTabPage世帯("世帯", "")
        mvarTabCtrl.AddPage(mvarTab世帯)
        Me.ResumeLayout()
    End Sub


    Private Class CTabPage共通
        Inherits HIMTools.Controls.CTabPageWithToolStrip

        Public 都道府県ID As Integer = 0
        Public 市町村CD As Integer = 0
        Public TBL市町村コード As New DataTable

        Public pTBLBK As DataTable
        Public pTBL As DataTable
        Public mvarGrid As HIMTools.Controls.DataGridViewWithDataView
        Public Sub New(ByVal pKey As String, ByVal sTitle As String)
            MyBase.New(False, True, pKey, IIf(sTitle = "", pKey, sTitle))

            都道府県ID = Val(SysAD.DB(sLRDB).DBProperty("都道府県ID").ToString)
            市町村CD = Val(SysAD.DB(sLRDB).DBProperty("市町村ID").ToString)

            Set市町村コード()

            mvarGrid = New HIMTools.Controls.DataGridViewWithDataView
            With mvarGrid
                .Dock = DockStyle.Fill
                .AllowUserToAddRows = False
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                Me.Controls.Add(mvarGrid)
            End With
        End Sub

        Public Sub Set市町村コード()
            If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
                If TBL市町村コード.Rows.Count = 0 Then
                    Dim cReader As New System.IO.StreamReader(SysAD.CustomReportFolder("共通様式") & "\code_list.csv", System.Text.Encoding.Default)
                    Dim LoopCount As Integer = 0

                    While (cReader.Peek() >= 0)
                        Dim stBuffer As String = cReader.ReadLine() ' ファイルを 1 行ずつ読み込む
                        Dim cAr As Object = Split(stBuffer, ",")

                        If LoopCount = 0 Then
                            With TBL市町村コード
                                .Columns.Add(cAr(0))
                                .Columns.Add(cAr(1))
                                .Columns.Add(cAr(2))
                                .Columns.Add(cAr(3))
                                .Columns.Add(cAr(4))
                            End With

                            LoopCount += 1
                        Else
                            Dim pRow As DataRow = TBL市町村コード.NewRow
                            pRow.Item("団体コード") = cAr(0)
                            pRow.Item("都道府県名（漢字）") = cAr(1)
                            pRow.Item("市区町村名（漢字）") = cAr(2)
                            pRow.Item("都道府県名（カナ）") = cAr(3)
                            pRow.Item("市区町村名（カナ）") = cAr(4)

                            TBL市町村コード.Rows.Add(pRow)
                        End If
                    End While
                End If
            End If
        End Sub

        Public ReadOnly Property 市町村コード(ByVal sValue As String)
            Get
                Dim CodePath As String = ""
                If System.IO.File.Exists(SysAD.CustomReportFolder("共通様式") & "\code_list.csv") Then
                    CodePath = SysAD.CustomReportFolder("共通様式") & "\code_list.csv"
                End If

                Dim cityCodeModel As CitiesCode.Interface.ICityCodeModel = New CitiesCode.Factory.CityCodeFactory().CreateCityCodeModel(CodePath)

                Dim jusyoModel As CitiesCode.Interface.IJusyoModel = cityCodeModel.GetCityCode(sValue)  ' 文字列より市町村コード取得
                If jusyoModel.MatchState = CitiesCode.Types.MatchType.Match Then
                    Return jusyoModel.CityCode()
                End If

                Return ""
            End Get
        End Property

        Public Function Find市町村コード(ByVal sValue As String)
            Dim Find大分類 As String = ""
            Dim Find小分類 As String = ""

            For Each pRow As DataRow In TBL市町村コード.Rows
                If InStr(sValue, pRow.Item("都道府県名（漢字）")) > 0 Then
                    Find大分類 = pRow.Item("都道府県名（漢字）")
                    Exit For
                End If
            Next

            For Each pRow As DataRow In TBL市町村コード.Rows
                If InStr(sValue, pRow.Item("市区町村名（漢字）")) > 0 Then
                    Find小分類 = pRow.Item("市区町村名（漢字）")
                    Exit For
                End If
            Next

            Dim pTBL As DataTable
            If Len(Find大分類) > 0 AndAlso Len(Find小分類) > 0 Then
                pTBL = New DataView(TBL市町村コード, String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "", DataViewRowState.CurrentRows).ToTable

                If pTBL.Rows.Count = 1 Then
                    Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[都道府県名（漢字）] = '{0}' And [市区町村名（漢字）] = '{1}'", Find大分類, Find小分類), "")
                    Return FindDataRow(0).Item("団体コード")
                Else
                    Return "(" & 都道府県ID & 市町村CD & ")"
                End If
            ElseIf Len(Find小分類) > 0 Then
                pTBL = New DataView(TBL市町村コード, String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "", DataViewRowState.CurrentRows).ToTable

                If pTBL.Rows.Count = 1 Then
                    Dim FindDataRow() As DataRow = pTBL.Select(String.Format("[市区町村名（漢字）] = '{0}'", Find小分類), "")
                    Return FindDataRow(0).Item("団体コード")
                Else
                    Return "(" & 都道府県ID & 市町村CD & ")"
                End If
            Else
                Return "(" & 都道府県ID & 市町村CD & ")"
            End If

        End Function

        Public Function Set市町村ID(ByRef pRow As DataRow, ByVal pValue As String) As String
            Dim Conv市町村CD As String = 市町村コード(pValue)
            If Len(pRow.Item("市町村ID").ToString) = 6 Then
                Return pRow.Item("市町村ID").ToString
            Else
                If Conv市町村CD = "" Then
                    Return Find市町村コード(pValue)
                Else
                    Return Conv市町村CD
                End If
            End If
        End Function

        Public Function CheckDBNull(ByRef pRow As DataRow, ByVal sItems() As String)
            For n As Integer = 0 To UBound(sItems)
                If IsDBNull(pRow.Item(sItems(n))) Then
                    Return True
                End If
            Next

            Return False
        End Function

        Public Sub Set備考(ByRef pRow As DataRow, ByRef pAddRow As DataRow, ByVal pColName As String)
            With pAddRow
                Select Case pColName
                    Case "市町村ID"
                        If IsDBNull(.Item("備考")) Then
                            .Item("備考") = "市町村IDもしくは住所を正しく入力してください。"
                        Else
                            .Item("備考") = .Item("備考") & vbCrLf & "市町村IDもしくは住所を正しく入力してください。"
                        End If
                    Case "大字ID", "大字", "本番", "登記簿面積", "実面積", "所有者ID", "耕作者ID"
                        If IsDBNull(pRow.Item(pColName)) Then
                            If IsDBNull(.Item("備考")) Then
                                .Item("備考") = pColName & "が未入力です"
                            Else
                                .Item("備考") = .Item("備考") & vbCrLf & pColName & "が未入力です"
                            End If
                        End If
                    Case "部分面積"
                        If .Item("一部現況") > 0 Then
                            If .Item("部分面積") = 0 Then
                                If IsDBNull(.Item("備考")) Then
                                    .Item("備考") = "部分面積が未入力です"
                                Else
                                    .Item("備考") = .Item("備考") & vbCrLf & "部分面積が未入力です"
                                End If
                            End If
                        End If
                    Case "世帯ID"
                        If .Item("世帯ID") > 0 Then

                        End If
                End Select
            End With
        End Sub
    End Class

    Private Class CTabPage農地
        Inherits CTabPage共通

        Public Sub New(ByVal pKey As String, ByVal sTitle As String)
            MyBase.New(pKey, sTitle)

            pTBLBK = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT 農地.*, [D:個人Info].氏名 AS 所有者名, [D:個人Info_1].氏名 AS 耕作者名 " & _
                                                         "FROM ((SELECT [D:農地Info].ID, [D:農地Info].市町村ID, [D:農地Info].大字ID, V_大字.名称 AS 大字, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])) AS 本番, IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0) AS 枝番, [D:農地Info].一部現況, [D:農地Info].登記簿面積, [D:農地Info].部分面積, [D:農地Info].実面積, [V_大字].[名称] & IIf(IsNull([V_小字].[名称]),'',IIf([V_小字].[名称]='-','',[V_小字].[名称])) & [D:農地Info].[地番] AS 土地所在, [D:農地Info].所有者ID, IIf([自小作別]<>0,[借受人ID],IIf([管理者ID]<>0,[管理者ID],[所有者ID])) AS 耕作者ID " & _
                                                         "FROM ([D:農地Info] LEFT JOIN V_大字 ON [D:農地Info].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D:農地Info].小字ID = V_小字.ID " & _
                                                         "WHERE ((([D:農地Info].大字ID)>0) AND ((V_大字.名称) Is Not Null Or (V_大字.名称)<>'') AND (([D:農地Info].地番) Is Not Null Or ([D:農地Info].地番)<>'')))  AS 農地 LEFT JOIN [D:個人Info] ON 農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON 農地.耕作者ID = [D:個人Info_1].ID " & _
                                                         "ORDER BY 農地.大字ID, 農地.本番, 農地.枝番 DESC;")

            Dim pTBLBK2 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT 農地.*, [D:個人Info].氏名 AS 所有者名, [D:個人Info_1].氏名 AS 耕作者名 " & _
                                                         "FROM ((SELECT [D_転用農地].ID, [D_転用農地].市町村ID, [D_転用農地].大字ID, V_大字.名称 AS 大字, IIf(InStr([地番],'-')>0,Val(Left([地番],InStr([地番],'-')-1)),Val([地番])) AS 本番, IIf(InStr([地番],'-')>0,Val(Mid([地番],InStr([地番],'-')+1)),0) AS 枝番, [D_転用農地].一部現況, [D_転用農地].登記簿面積, [D_転用農地].部分面積, [D_転用農地].実面積, [V_大字].[名称] & IIf(IsNull([V_小字].[名称]),'',IIf([V_小字].[名称]='-','',[V_小字].[名称])) & [D_転用農地].[地番] AS 土地所在, [D_転用農地].所有者ID, IIf([自小作別]<>0,[借受人ID],IIf([管理者ID]<>0,[管理者ID],[所有者ID])) AS 耕作者ID " & _
                                                         "FROM ([D_転用農地] LEFT JOIN V_大字 ON [D_転用農地].大字ID = V_大字.ID) LEFT JOIN V_小字 ON [D_転用農地].小字ID = V_小字.ID " & _
                                                         "WHERE ((([D_転用農地].大字ID)>0) AND ((V_大字.名称) Is Not Null Or (V_大字.名称)<>'') AND (([D_転用農地].地番) Is Not Null Or ([D_転用農地].地番)<>''))) AS 農地 LEFT JOIN [D:個人Info] ON 農地.所有者ID = [D:個人Info].ID) LEFT JOIN [D:個人Info] AS [D:個人Info_1] ON 農地.耕作者ID = [D:個人Info_1].ID " & _
                                                         "ORDER BY 農地.大字ID, 農地.本番, 農地.枝番 DESC;")
            pTBLBK.Merge(pTBLBK2)


            pTBL = New DataTable
            SetDataBaseColumns(pTBL)
            For Each pRow As DataRow In pTBLBK.Rows
                Dim 市町村ID As String = Set市町村ID(pRow, pRow.Item("土地所在").ToString)
                Dim Flg市町村 As Boolean = IIf(Len(市町村ID) = 6, True, False)
                If Flg市町村 = False AndAlso CheckDBNull(pRow, {"大字ID", "大字", "本番", "登記簿面積", "部分面積", "実面積", "所有者ID", "耕作者ID"}) Then
                    Dim pAddRow As DataRow = pTBL.NewRow
                    With pAddRow
                        .Item("ID") = pRow.Item("ID")
                        .Item("市町村ID") = 市町村ID
                        .Item("大字ID") = Val(pRow.Item("大字ID").ToString)
                        .Item("大字") = pRow.Item("大字").ToString
                        .Item("本番") = Val(pRow.Item("本番").ToString)
                        .Item("枝番") = Val(pRow.Item("枝番").ToString)
                        .Item("一部現況") = Val(pRow.Item("一部現況").ToString)
                        .Item("登記簿面積") = Val(pRow.Item("登記簿面積").ToString)
                        .Item("部分面積") = Val(pRow.Item("部分面積").ToString)
                        .Item("実面積") = Val(pRow.Item("実面積").ToString)
                        .Item("所有者ID") = Val(pRow.Item("所有者ID").ToString)
                        .Item("所有者名") = pRow.Item("所有者名").ToString
                        .Item("耕作者ID") = Val(pRow.Item("耕作者ID").ToString)
                        .Item("耕作者名") = pRow.Item("耕作者名").ToString

                        Set備考(pRow, pAddRow, "市町村ID")
                        Set備考(pRow, pAddRow, "大字ID")
                        Set備考(pRow, pAddRow, "大字")
                        Set備考(pRow, pAddRow, "本番")
                        Set備考(pRow, pAddRow, "登記簿面積")
                        Set備考(pRow, pAddRow, "部分面積")
                        Set備考(pRow, pAddRow, "実面積")
                        Set備考(pRow, pAddRow, "所有者ID")
                        Set備考(pRow, pAddRow, "耕作者ID")

                        pTBL.Rows.Add(pAddRow)
                    End With
                End If
            Next

            mvarGrid.SetDataView(pTBL, "", "")
        End Sub

        Private Function SetDataBaseColumns(ByRef pTBL As DataTable)
            With pTBL
                .Columns.Add("ID", GetType(Integer))
                .Columns.Add("市町村ID", GetType(String)) '必須
                .Columns.Add("大字ID", GetType(Integer)) '必須
                .Columns.Add("大字", GetType(String)) '必須
                .Columns.Add("本番", GetType(Integer)) '必須
                .Columns.Add("枝番", GetType(Integer))
                .Columns.Add("一部現況", GetType(Integer))
                .Columns.Add("登記簿面積", GetType(Decimal)) '必須
                .Columns.Add("部分面積", GetType(Decimal)) '必須
                .Columns.Add("実面積", GetType(Decimal)) '必須
                .Columns.Add("所有者ID", GetType(Decimal)) '必須
                .Columns.Add("所有者名", GetType(String))
                .Columns.Add("耕作者ID", GetType(Decimal)) '必須
                .Columns.Add("耕作者名", GetType(String))
                .Columns.Add("備考", GetType(String))
            End With

            Return pTBL
        End Function
    End Class

    Private Class CTabPage個人
        Inherits CTabPage共通

        Public Sub New(ByVal pKey As String, ByVal sTitle As String)
            MyBase.New(pKey, sTitle)

            pTBLBK = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:個人Info].ID, [D:個人Info].世帯ID, [D:個人Info].氏名, [D:個人Info].[フリガナ], [D:個人Info].市町村ID, [D:個人Info].住所 " & _
                                                                        "FROM [D:個人Info] " & _
                                                                        "WHERE ((([D:個人Info].氏名) Is Not Null Or ([D:個人Info].氏名)<>''));")

            pTBL = New DataTable
            SetDataBaseColumns(pTBL)
            For Each pRow As DataRow In pTBLBK.Rows
                Dim 市町村ID As String = Set市町村ID(pRow, pRow.Item("住所").ToString)
                Dim Flg市町村 As Boolean = IIf(Len(市町村ID) = 6, True, False)
                If Flg市町村 = False AndAlso CheckDBNull(pRow, {"世帯ID", "氏名", "フリガナ", "市町村ID"}) Then
                    Dim pAddRow As DataRow = pTBL.NewRow
                    With pAddRow
                        .Item("ID") = pRow.Item("ID")
                        .Item("世帯ID") = pRow.Item("世帯ID")
                        .Item("氏名") = pRow.Item("氏名").ToString
                        .Item("フリガナ") = pRow.Item("フリガナ").ToString
                        .Item("市町村ID") = 市町村ID
                        .Item("住所") = pRow.Item("住所").ToString
                    End With
                End If
            Next

            mvarGrid.SetDataView(pTBL, "", "")
        End Sub

        Private Function SetDataBaseColumns(ByRef pTBL As DataTable)
            With pTBL
                .Columns.Add("ID", GetType(Decimal))
                .Columns.Add("世帯ID", GetType(Decimal)) '必須
                .Columns.Add("氏名", GetType(String)) '必須
                .Columns.Add("フリガナ", GetType(String)) '必須
                .Columns.Add("市町村ID", GetType(String)) '必須
                .Columns.Add("住所", GetType(Integer))
                .Columns.Add("備考", GetType(String))
            End With

            Return pTBL
        End Function
    End Class

    Private Class CTabPage世帯
        Inherits CTabPage共通

        Public Sub New(ByVal pKey As String, ByVal sTitle As String)
            MyBase.New(pKey, sTitle)

            pTBLBK = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT [D:世帯Info].ID, [D:世帯Info].農地所有区分, [D:世帯Info].世帯主ID, [D:個人Info].氏名, [D:個人Info].性別, [D:世帯Info].市町村ID, [D:個人Info].住所 " & _
                                                                        "FROM [D:世帯Info] INNER JOIN [D:個人Info] ON [D:世帯Info].世帯主ID = [D:個人Info].ID;")

            mvarGrid.SetDataView(pTBL, "", "")
        End Sub
    End Class
End Class




