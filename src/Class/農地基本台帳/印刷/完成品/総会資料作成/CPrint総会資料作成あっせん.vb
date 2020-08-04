

Public Class CPrint総会資料作成あっせん出し手
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        申請者A(pSheet, pRow)

        Select Case Val(pRow.Item("区分").ToString)
            Case 1 : pSheet.ValueReplace("{内容}", "売渡")
            Case 2 : pSheet.ValueReplace("{内容}", "貸付")
            Case Else : pSheet.ValueReplace("{内容}", "その他")
        End Select
        pSheet.ValueReplace("{期間}", pRow.Item("条件B").ToString)
        pSheet.ValueReplace("{条件}", Replace(pRow.Item("条件A").ToString, vbCrLf, "&#10;"))
        pSheet.ValueReplace("{備考}", Replace(pRow.Item("備考").ToString, vbCrLf, "&#10;"))

        複数土地設定(pSheet, pRow, Nothing)
    End Sub
End Class

Public Class CPrint総会資料作成あっせん受け手
    Inherits CPrint総会資料作成

    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        '複数申請人A

        SetNO(pSheet, True)
        申請者A(pSheet, pRow)
        If IsDate(pRow.Item("許可年月日").ToString) Then
            pSheet.ValueReplace("{年月日}", 和暦Format(pRow.Item("許可年月日"), "gyy.M.d"))
        End If

        pSheet.ValueReplace("{希望地目}", pRow.Item("用途").ToString)
        pSheet.ValueReplace("{希望面積}", Val(pRow.Item("数量").ToString).ToString("#,##0"))

        Select Case Val(pRow.Item("区分").ToString)
            Case 1 : pSheet.ValueReplace("{内容}", "買受")
            Case 2 : pSheet.ValueReplace("{内容}", "借受")
            Case Else : pSheet.ValueReplace("{内容}", "その他")
        End Select
        pSheet.ValueReplace("{条件}", Replace(pRow.Item("条件A").ToString, vbCrLf, "&#10;"))
        pSheet.ValueReplace("{備考}", Replace(pRow.Item("備考").ToString, vbCrLf, "&#10;"))

        Dim pView As New DataView(App農地基本台帳.DataMaster.Body, "[Class]='営農類型' AND [ID]=" & Val(pRow.Item("営農類型").ToString), "", DataViewRowState.CurrentRows)
        If pView.Count = 0 Then
            pSheet.ValueReplace("{経営形態}", "")
        Else
            pSheet.ValueReplace("{経営形態}", pView(0).Item("名称"))
        End If
        pSheet.ValueReplace("{作付面積㌃}", (Val(pRow.Item("経営面積A").ToString) / 100).ToString("#,##0"))
        pSheet.ValueReplace("{経営面積㌃}", (Val(pRow.Item("経営面積A").ToString) / 100).ToString("#,##0"))


        pSheet.ValueReplace("{男子専業者}", pRow.Item("働手男数B").ToString)
        pSheet.ValueReplace("{女子専業者}", pRow.Item("働手女数B").ToString)
        pSheet.ValueReplace("{専業}", "")
        pSheet.ValueReplace("{非専業}", "")
        pSheet.ValueReplace("{あっせん}", "")
    End Sub
End Class
