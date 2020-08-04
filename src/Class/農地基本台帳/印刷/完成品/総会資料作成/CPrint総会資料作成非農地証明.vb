'20160406霧島

Public Class CPrint総会資料作成非農地証明
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)
        調査委員(pSheet, pRow)

        If Not IsDBNull(pRow.Item("代理人A")) AndAlso Not pRow.Item("代理人A") = 0 AndAlso pRow.Item("代理人A") <> pRow.Item("申請者A") Then
            pSheet.ValueReplace("{申請者X氏名}", pRow.Item("代理人名").ToString)
            pSheet.ValueReplace("{申請者X住所}", pRow.Item("代理人住所").ToString)
            pSheet.ValueReplace("{申請者X集落名}", "")
        Else
            pSheet.ValueReplace("{申請者X氏名}", pRow.Item("氏名A").ToString)
            pSheet.ValueReplace("{申請者X住所}", pRow.Item("住所A").ToString)
            pSheet.ValueReplace("{申請者X集落名}", pRow.Item("集落A").ToString)
        End If

        pSheet.ValueReplace("{変更年月日}", pRow.Item("変更年月日txt").ToString)
        pSheet.ValueReplace("{申請理由}", pRow.Item("申請理由A").ToString)
        pSheet.ValueReplace("{意見}", pRow.Item("意見").ToString)
        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)

        複数土地設定(pSheet, pRow, Nothing)
        区分設定(pSheet, pRow)
    End Sub
End Class
