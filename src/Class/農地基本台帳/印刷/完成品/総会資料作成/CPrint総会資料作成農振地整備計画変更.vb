
Public Class CPrint総会資料作成農振地整備計画変更
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional ByVal pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)

        複数土地設定(pSheet, pRow, Nothing)
        区分設定(pSheet, pRow)

        If IsDBNull(pRow.Item("区分")) Then
            pSheet.ValueReplace("{区分}", "")
        Else
            Dim 変更区分名 As String = App農地基本台帳.DataMaster.Body.Rows.Find({pRow.Item("区分"), "農振整備計画変更区分"}).Item("名称").ToString
            pSheet.ValueReplace("{区分}", 変更区分名)
        End If

        If IsDBNull(pRow.Item("用途")) Then
            pSheet.ValueReplace("{転用目的}", "")
        Else
            Dim 変更目的名 As String = App農地基本台帳.DataMaster.Body.Rows.Find({pRow.Item("用途"), "農振用途区分"}).Item("名称").ToString
            pSheet.ValueReplace("{転用目的}", 変更目的名)
        End If

        If Val(pRow.Item("用途").ToString) > 0 Then
            pSheet.ValueReplace("{建築面積}", pRow.Item("建築面積").ToString)
            pSheet.ValueReplace("{棟数}", " (" & pRow.Item("数量").ToString & ")")
        Else
            pSheet.ValueReplace("{建築面積}", "")
            pSheet.ValueReplace("{棟数}", "")
        End If

        pSheet.ValueReplace("{区分}", pRow.Item("区分").ToString)
        pSheet.ValueReplace("{申請理由}", pRow.Item("申請理由A").ToString)

        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
    End Sub
End Class
