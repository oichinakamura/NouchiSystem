Public Class CPrint総会資料作成転用事業計画変更
    Inherits CPrint総会資料作成

    Private pTBL As New DataTable
    Public Overrides Sub SetDataRow(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As System.Data.DataRowView, Optional pObj As Object = Nothing)
        SetNO(pSheet, True)

        申請者C(pSheet, pRow)
        申請者A(pSheet, pRow)

        AddColumns("土地所在", GetType(String))
        AddColumns("登記地目", GetType(String))
        AddColumns("現況地目", GetType(String))
        AddColumns("面積", GetType(Decimal))

        Dim BeforeRow As DataRow
        Dim Ar As Object = Split(pRow.Item("予備1"), vbCrLf)
        For n As Integer = 0 To UBound(Ar)
            Dim ArColumn As Object = Split(Ar(n), ";")
            BeforeRow = pTBL.NewRow()

            BeforeRow("土地所在") = ArColumn(0)
            BeforeRow("登記地目") = ArColumn(1)
            BeforeRow("現況地目") = ArColumn(2)
            BeforeRow("面積") = ArColumn(3)

            pTBL.Rows.Add(BeforeRow)
        Next

        For Each bRow As DataRow In pTBL.Rows
            変更前複数土地設定(pSheet, bRow)
        Next

        pSheet.ValueReplace("{変更前土地の所在}", p土地所在)
        pSheet.ValueReplace("{変更前登記地目}", p登記地目)
        pSheet.ValueReplace("{変更前現況地目}", p現況地目)
        pSheet.ValueReplace("{変更前面積}", p面積)
        pSheet.ValueReplace("{変更前面積計}", p面積計)
        pSheet.ValueReplace("{変更前田筆数計}", p田筆数計)
        pSheet.ValueReplace("{変更前畑筆数計}", p畑筆数計)
        pSheet.ValueReplace("{変更前筆数計}", p筆数計)

        InitializationVariable()

        pSheet.ValueReplace("{変更前転用目的}", pRow.Item("予備2"))
        pSheet.ValueReplace("{変更後転用目的}", pRow.Item("申請理由A"))
        pSheet.ValueReplace("{理由}", pRow.Item("申請理由B"))
        pSheet.ValueReplace("{承認区分}", "")

        '/***変更後土地所在***/
        複数土地設定(pSheet, pRow, Nothing)
        転用共通(pSheet, pRow)
        貸借共通(pSheet, pRow)

        pSheet.ValueReplace("{備考}", pRow.Item("備考").ToString)
    End Sub

    Private Sub 申請者C(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRowView)
        pSheet.ValueReplace("{申請者Ｃ氏名}", pRow.Item("氏名C").ToString)
        pSheet.ValueReplace("{申請者Ｃ住所}", pRow.Item("住所C").ToString)
        pSheet.ValueReplace("{申請者Ｃ職業}", pRow.Item("職業C").ToString)
        pSheet.ValueReplace("{申請者Ｃ年齢}", IIf(Val(pRow.Item("年齢C").ToString) = 0, "-", pRow.Item("年齢C").ToString))
        pSheet.ValueReplace("{申請者Ｃ申請理由}", pRow.Item("予備2").ToString)
        pSheet.ValueReplace("{申請者Ｃ集落名}", pRow.Item("集落C").ToString)
    End Sub

    Private Sub AddColumns(ByVal pColumnName As String, ByVal pColumnType As Object)
        If Not pTBL.Columns.Contains(pColumnName) Then
            pTBL.Columns.Add(New DataColumn(pColumnName, pColumnType))
        End If
    End Sub

    Private Sub InitializationVariable()
        pTBL.Clear()
        p土地所在 = ""
        p登記地目 = ""
        p現況地目 = ""
        p面積 = ""
        p面積計 = 0
        p田筆数計 = 0
        p畑筆数計 = 0
        p筆数計 = 0
    End Sub

    Private p土地所在 As String = ""
    Private p登記地目 As String = ""
    Private p現況地目 As String = ""
    Private p面積 As String = ""
    Private p面積計 As Decimal = 0
    Private p田筆数計 As Integer = 0
    Private p畑筆数計 As Integer = 0
    Private p筆数計 As Integer = 0
    Private Sub 変更前複数土地設定(ByRef pSheet As HimTools2012.Excel.XMLSS2003.XMLSSWorkSheet, ByRef pRow As DataRow)
        If p土地所在 = "" Then : p土地所在 = pRow.Item("土地所在")
        Else : p土地所在 = p土地所在 & vbCrLf & pRow.Item("土地所在")
        End If

        If p登記地目 = "" Then : p登記地目 = pRow.Item("登記地目")
        Else : p登記地目 = p登記地目 & vbCrLf & pRow.Item("登記地目")
        End If

        If p現況地目 = "" Then : p現況地目 = pRow.Item("現況地目")
        Else : p現況地目 = p現況地目 & vbCrLf & pRow.Item("現況地目")
        End If

        If p面積 = "" Then : p面積 = Format(pRow.Item("面積"), "#,#")
        Else : p面積 = p面積 & vbCrLf & Format(pRow.Item("面積"), "#,#")
        End If
        p面積計 += pRow.Item("面積")

        If pRow.Item("登記地目").ToString.IndexOf("田") > -1 AndAlso pRow.Item("登記地目").ToString <> "塩田" Then
            p田筆数計 += 1
        ElseIf pRow.Item("登記地目").ToString.IndexOf("畑") > -1 Then
            p畑筆数計 += 1
        End If
        p筆数計 += 1
    End Sub
End Class
