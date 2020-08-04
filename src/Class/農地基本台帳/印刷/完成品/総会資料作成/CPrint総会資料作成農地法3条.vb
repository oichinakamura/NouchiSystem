'20160406霧島

Public Class CPrint総会資料作成農地法3条
    Inherits CPrint総会資料作成


    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)
        受付情報(pSheet, pRow)

        申請者A(pSheet, pRow)
        申請者B(pSheet, pRow)
        調査委員(pSheet, pRow)

        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
        pSheet.ValueReplace("{契約期間}", "")

        複数土地設定(pSheet, pRow, Nothing)

        pSheet.ValueReplace("{図頁}", "")

        Select Case pRow.Item("法令")
            Case 30, 32
                pSheet.ValueReplace("{契約内容}", "所有権移転")

                Select Case Val(pRow.Item("所有権移転の種類").ToString)
                    Case 1 : pSheet.ValueReplace("{種類}", "売買")
                    Case 2 : pSheet.ValueReplace("{種類}", "贈与")
                    Case 3 : pSheet.ValueReplace("{種類}", "交換")
                    Case Else
                        pSheet.ValueReplace("{種類}", "")
                End Select
            Case 31, 33
                貸借共通(pSheet, pRow)
                pSheet.ValueReplace("{種類}", "")

                If Val(pRow.Item("区分地上権").ToString) > 0 Then
                    pSheet.ValueReplace("{区分地上権}", "区分地上権設定")
                    pSheet.ValueReplace("{区分地上権内容}", pRow.Item("区分地上権内容").ToString)
                End If

                Select Case Val(pRow.Item("権利種類").ToString)
                    Case 1 : pSheet.ValueReplace("{契約内容}", "賃借権")
                    Case 2, 104 : pSheet.ValueReplace("{契約内容}", "使用貸借権")
                    Case 4, 101 : pSheet.ValueReplace("{契約内容}", "地上権")
                    Case 5, 102 : pSheet.ValueReplace("{契約内容}", "永小作権")
                    Case 6, 103 : pSheet.ValueReplace("{契約内容}", "質権")
                    Case 7 : pSheet.ValueReplace("{契約内容}", "期間借地")
                    Case 8 : pSheet.ValueReplace("{契約内容}", "残存小作地")
                    Case 9 : pSheet.ValueReplace("{契約内容}", "使用賃借")
                    Case 105 : pSheet.ValueReplace("{契約内容}", "賃貸借権")
                    Case 106 : pSheet.ValueReplace("{契約内容}", "使用貸借権（期間借地）")
                    Case 107 : pSheet.ValueReplace("{契約内容}", "賃貸借権（期間借地）")
                    Case 108 : pSheet.ValueReplace("{契約内容}", "使用貸借権（円滑化）")
                    Case 109 : pSheet.ValueReplace("{契約内容}", "賃貸借権（円滑化）")
                    Case 199 : pSheet.ValueReplace("{契約内容}", "調査中")
                    Case Else
                        pSheet.ValueReplace("{契約内容}", "その他")
                End Select



        End Select
    End Sub
End Class
