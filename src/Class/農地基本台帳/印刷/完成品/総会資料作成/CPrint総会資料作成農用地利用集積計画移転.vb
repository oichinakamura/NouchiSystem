'20160406霧島


Public Class CPrint総会資料作成農用地利用集積計画移転
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)
        調査委員(pSheet, pRow)

        pSheet.ValueReplace("{申請者Ｂ氏名}", pRow.Item("氏名B").ToString)
        pSheet.ValueReplace("{申請者Ｂ住所}", pRow.Item("住所B").ToString)
        pSheet.ValueReplace("{申請者Ｂ職業}", pRow.Item("職業B").ToString)
        pSheet.ValueReplace("{申請者Ｂ年齢}", IIf(Val(pRow.Item("年齢B").ToString) = 0, "-", pRow.Item("年齢B").ToString))
        pSheet.ValueReplace("{申請者Ｂ経営面積}", pRow.Item("経営面積B").ToString)
        pSheet.ValueReplace("{申請者Ｂ申請理由}", pRow.Item("申請理由B").ToString)
        pSheet.ValueReplace("{申請者Ｂ集落名}", pRow.Item("集落B").ToString)
        pSheet.ValueReplace("{申請者Ｂ労力}", "")

        pSheet.ValueReplace("{申請者Ｃ氏名}", pRow.Item("氏名C").ToString)
        pSheet.ValueReplace("{申請者Ｃ住所}", pRow.Item("住所C").ToString)
        pSheet.ValueReplace("{申請者Ｃ集落名}", pRow.Item("集落C").ToString)
        pSheet.ValueReplace("{利用権内容}", pRow.Item("利用権内容").ToString)
        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)

        貸借共通(pSheet, pRow)
        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{同意書}", "")
        pSheet.ValueReplace("{図頁}", "")

    End Sub
End Class

