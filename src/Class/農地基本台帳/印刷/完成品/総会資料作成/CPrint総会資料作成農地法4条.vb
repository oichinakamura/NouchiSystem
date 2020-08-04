'20160406霧島

Public Class CPrint総会資料作成農地法4条
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        受付情報(pSheet, pRow)

        申請者A(pSheet, pRow)
        調査委員(pSheet, pRow)

        pSheet.ValueReplace("{市町村名}", SysAD.市町村.市町村名)
        pSheet.ValueReplace("{転用目的}", pRow.Item("申請理由A"))
        pSheet.ValueReplace("{転用内容}", pRow.Item("申請理由A"))
        pSheet.ValueReplace("{申請事由}", pRow.Item("申請理由B"))

        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
        pSheet.ValueReplace("{契約期間}", "")
        pSheet.ValueReplace("{権利種別}", "所有権")

        Dim p集計結果 As C筆明細と集計作成 = 複数土地設定(pSheet, pRow, Nothing)

        Dim s地目 As New System.Text.StringBuilder
        Dim s面積 As New System.Text.StringBuilder

        If p集計結果.田面計 > 0 Then
            s地目.Append("田")
            s面積.Append(String.Format("{0:0,0}㎡", p集計結果.田面計))
        End If
        If p集計結果.畑面計 > 0 Then
            s地目.Append(IIf(s地目.Length > 0, "&#10;&#10;", "") & "畑")
            s面積.Append(IIf(s面積.Length > 0, "&#10;&#10;", "") & String.Format("{0:0,0}㎡", p集計結果.畑面計))
        End If
        If p集計結果.田面計 > 0 AndAlso p集計結果.畑面計 > 0 Then
            s地目.Append("&#10;&#10;計")
            s面積.Append(String.Format("&#10;&#10;{0:0,0}", p集計結果.総面積))
        End If
        pSheet.ValueReplace("{諮問地目}", s地目.ToString())
        pSheet.ValueReplace("{諮問面積}", s面積.ToString())

        転用共通(pSheet, pRow)

        pSheet.ValueReplace("{図頁}", "")
    End Sub
End Class

