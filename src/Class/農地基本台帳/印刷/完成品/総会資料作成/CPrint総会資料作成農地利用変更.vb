
Public Class CPrint総会資料作成農地利用変更
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)

        複数土地設定(pSheet, pRow, Nothing)
        区分設定(pSheet, pRow)
        転用共通(pSheet, pRow)
        pSheet.ValueReplace("{発行年月日}", 和暦Format(Now))
        pSheet.ValueReplace("{周囲の状況}", pRow.Item("申請地目安").ToString)
        pSheet.ValueReplace("{変更後の使用目的}", pRow.Item("用途").ToString)
        pSheet.ValueReplace("{申請理由}", pRow.Item("申請理由A").ToString)
        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
    End Sub
End Class