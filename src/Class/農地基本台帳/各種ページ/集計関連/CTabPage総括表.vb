
Imports HimTools2012.controls.DataGridViewWithDataView

Public Class CTabPage総括表
    Inherits HimTools2012.TabPages.CTabPageWithDataGridView

    Private pTable As New DataTable("総括表")

    Public Sub New(pTarget As HimTools2012.TargetSystem.CTargetObjWithView)
        MyBase.New(True, "総括表" & pTarget.Key.KeyValue, "総括表 " & pTarget.ToString, ObjectMan)

        pTable.Columns.Add("No", GetType(Integer)).DefaultValue = 0
        pTable.Columns.Add("地目", GetType(String))
        pTable.Columns.Add("経営自作面積", GetType(Decimal)).DefaultValue = 0
        pTable.Columns.Add("経営借入面積", GetType(Decimal)).DefaultValue = 0
        pTable.Columns.Add("経営総面積", GetType(Decimal), "[経営自作面積]+[経営借入面積]")
        pTable.Columns.Add("経営筆数", GetType(Integer)).DefaultValue = 0
        pTable.Columns.Add("貸付面積", GetType(Decimal)).DefaultValue = 0
        pTable.PrimaryKey = {pTable.Columns("地目")}

        Select Case TypeName(pTarget)
            Case "CObj農家" : 集計(pTarget.ID, pTarget, "管理世帯ID", "所有世帯ID", "借受世帯ID")
            Case "CObj個人" : 集計(pTarget.ID, pTarget, "管理者ID", "所有者ID", "借受人ID")
            Case Else
                Stop
        End Select
        mvarGrid.SetDataView(pTable, "", "No", AutoGenerateColumnsMode.AutoGenerateEnable)
        mvarGrid.Columns("No").Visible = False
    End Sub

    Public Sub 集計(nID As Long, pTarget As HimTools2012.TargetSystem.CTargetObjWithView, ByVal s管理 As String, ByVal s所有 As String, ByVal s借受 As String)
        Dim sWhere As String = String.Format("([" & s管理 & "]={0} Or [" & s所有 & "]={0}) Or ([自小作別]<>0 AND [" & s借受 & "]={0})", pTarget.ID)
        App農地基本台帳.TBL農地.MergePlus("SELECT * FROM [D:農地Info] WHERE {0}", sWhere)

        Dim p合計 As DataRow = pTable.NewRow
        p合計.Item("No") = 10
        p合計.Item("地目") = "合計"

        GetRow("田", 1)
        GetRow("畑", 2)
        GetRow("樹園地", 3)

        Dim pView As New DataView(App農地基本台帳.TBL農地.Body, sWhere, "", DataViewRowState.CurrentRows)
        For Each pRow As DataRowView In pView

            Dim n田面積 As Decimal = Val(pRow.Item("田面積").ToString) * Math.Abs(CInt(Val(pRow.Item("農地状況").ToString) < 20))
            Dim n畑面積 As Decimal = Val(pRow.Item("畑面積").ToString) * Math.Abs(CInt(Val(pRow.Item("農地状況").ToString) < 20))
            Dim n樹園地 As Decimal = Val(pRow.Item("樹園地").ToString) * Math.Abs(CInt(Val(pRow.Item("農地状況").ToString) < 20))

            Select Case Val(pRow.Item("自小作別").ToString)
                Case 0
                    If n田面積 > 0 Then 経営追加(p合計, GetRow("田", 1), "経営自作面積", n田面積)
                    If n畑面積 > 0 Then 経営追加(p合計, GetRow("畑", 2), "経営自作面積", n畑面積)
                    If n樹園地 > 0 Then 経営追加(p合計, GetRow("樹園地", 3), "経営自作面積", n樹園地)
                Case Else
                    If (Val(pRow.Item(s所有).ToString) = nID OrElse Val(pRow.Item(s管理).ToString) = nID) AndAlso Not pRow.Item(s借受) = nID Then
                        If n田面積 > 0 Then 貸付追加(p合計, GetRow("田", 1), "貸付面積", n田面積) '20170128 経営借入面積→貸付面積
                        If n畑面積 > 0 Then 貸付追加(p合計, GetRow("畑", 2), "貸付面積", n畑面積) '20170128 経営借入面積→貸付面積
                        If n樹園地 > 0 Then 貸付追加(p合計, GetRow("樹園地", 3), "貸付面積", n樹園地) '20170128 経営借入面積→貸付面積
                    ElseIf Not (Val(pRow.Item(s所有).ToString) = nID OrElse Val(pRow.Item(s管理).ToString) = nID) AndAlso pRow.Item(s借受) = nID Then
                        If n田面積 > 0 Then 経営追加(p合計, GetRow("田", 1), "経営借入面積", n田面積)
                        If n畑面積 > 0 Then 経営追加(p合計, GetRow("畑", 2), "経営借入面積", n畑面積)
                        If n樹園地 > 0 Then 経営追加(p合計, GetRow("樹園地", 3), "経営借入面積", n樹園地)
                    Else
                        If n田面積 > 0 Then 経営追加(p合計, GetRow("田", 1), "経営自作面積", n田面積)
                        If n畑面積 > 0 Then 経営追加(p合計, GetRow("畑", 2), "経営自作面積", n畑面積)
                        If n樹園地 > 0 Then 経営追加(p合計, GetRow("樹園地", 3), "経営自作面積", n樹園地)
                    End If
            End Select
        Next
        pTable.Rows.Add(p合計)
    End Sub

    Private Sub 経営追加(ByRef p合計 As DataRow, ByRef pRowT As DataRow, ByVal sField As String, nArea As Decimal)
        pRowT.Item(sField) += nArea
        pRowT.Item("経営筆数") += 1
        p合計.Item(sField) += nArea
        p合計.Item("経営筆数") += 1
    End Sub
    Private Sub 貸付追加(ByRef p合計 As DataRow, ByRef pRowT As DataRow, ByVal sField As String, nArea As Decimal)
        pRowT.Item(sField) += nArea
        p合計.Item(sField) += nArea
    End Sub

    Private Function GetRow(ByVal s地目 As String, ByVal No As Integer) As DataRow
        Dim pRowT As DataRow = pTable.Rows.Find(s地目)
        If pRowT Is Nothing Then
            pRowT = pTable.NewRow
            pRowT.Item("No") = No
            pRowT.Item("地目") = s地目
            pTable.Rows.Add(pRowT)
        End If

        Return pRowT
    End Function


    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

End Class
