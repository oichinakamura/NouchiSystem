

Public Class CPrint総会資料作成18条解約
    Inherits CPrint総会資料作成
    '  Private Type dt総括表
    '    s公告日 As String
    '    s始期 As String
    '    s終期 As String
    '    n期間 As Long
    '    s解約日 As String
    '    n面積(5) As Double
    '    n奨励金額(5) As Long
    '    s奨励金種類(5) As String
    'End Type
    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)
        申請者B(pSheet, pRow)

        pSheet.ValueReplace("{受付番号}", pRow.Item("受付番号").ToString)
        pSheet.ValueReplace("{受付年月日}", 和暦Format(CDate(pRow.Item("受付年月日"))))

        pSheet.ValueReplace("{承認年月日}", pRow.Item("解約年月日").ToString)

        If IsDate(pRow.Item("進達年月日").ToString) Then
            pSheet.ValueReplace("{成立年月日}", 和暦Format(pRow.Item("進達年月日")))
        Else
            pSheet.ValueReplace("{成立年月日}", "未入力")
        End If
        If IsDate(pRow.Item("解約年月日").ToString) Then
            pSheet.ValueReplace("{解約年月日}", 和暦Format(pRow.Item("解約年月日")))
        Else
            pSheet.ValueReplace("{解約年月日}", "未入力")
        End If
        If IsDate(pRow.Item("完了報告年月日").ToString) Then
            pSheet.ValueReplace("{引渡年月日}", 和暦Format(pRow.Item("完了報告年月日")))
        Else
            pSheet.ValueReplace("{引渡年月日}", "未入力")
        End If

        pSheet.ValueReplace("{農地区分}", pRow.Item("農地区分名称").ToString)
        pSheet.ValueReplace("{解約の理由}", Replace(pRow.Item("申請理由A").ToString, vbCrLf, "&#10;"))
        pSheet.ValueReplace("{解約の条件}", Replace(pRow.Item("申請理由B").ToString, vbCrLf, "&#10;"))
        pSheet.ValueReplace("{備考}", Replace(pRow.Item("備考").ToString, vbCrLf, "&#10;"))

        複数土地設定(pSheet, pRow, Nothing)
    End Sub
End Class
