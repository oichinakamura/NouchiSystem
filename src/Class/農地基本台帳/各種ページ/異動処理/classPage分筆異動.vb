
Public Class classPage分筆異動
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarPanel As New FlowLayoutPanel
    Private WithEvents mvartext地番1 As TextBox
    Private WithEvents mvartext地番2 As TextBox
    Private WithEvents mvarNum面積1 As System.Windows.Forms.NumericUpDown
    Private WithEvents mvarNum面積2 As System.Windows.Forms.NumericUpDown
    Private WithEvents mvarCHK面積調整 As System.Windows.Forms.CheckBox

    Private WithEvents mvarOK As Button
    Private n元ID As Integer = 0
    Private n元地番 As String = ""

    Private n元面積 As Decimal = 0
    Private nWidth As Decimal = 0
    Public Sub New(p農地 As CObj農地)
        MyBase.New(True, True, "分筆処理" & p農地.ID, "分筆処理[" & p農地.土地所在 & "]")
        mvarPanel.BackColor = Color.LightYellow
        mvarPanel.Dock = DockStyle.Fill
        n元ID = p農地.ID
        n元面積 = p農地.登記簿面積

        AddGroup("分筆前")
        AddLabel("元　番", False, 80)
        AddLabel(p農地.地番, True, 160)
        n元地番 = p農地.地番
        AddLabel("面　積", False, 80)
        AddLabel(n元面積 & "㎡", True, 160)

        AddGroup("分筆後")
        AddLabel("新番①", False, 80)
        mvartext地番1 = AddTextBox(p農地.地番, True, 160)
        AddLabel("面　積", False, 80)
        mvarNum面積1 = AddTUpdown(p農地.登記簿面積, p農地.登記簿面積, True, 160)

        AddLabel("新番②", False, 80)
        mvartext地番2 = AddTextBox(p農地.地番, True, 160)
        AddLabel("面　積", False, 80)
        mvarNum面積2 = AddTUpdown(0, p農地.登記簿面積, True, 160)

        AddGroup("")
        mvarCHK面積調整 = New System.Windows.Forms.CheckBox
        mvarCHK面積調整.Checked = True
        mvarCHK面積調整.Text = "面積の自動調整"
        mvarCHK面積調整.Width = 160
        mvarPanel.Controls.Add(mvarCHK面積調整)

        mvarOK = New Button
        mvarOK.Text = "実行"
        mvarOK.Enabled = False
        mvarPanel.Controls.Add(mvarOK)
        Me.ControlPanel.Add(mvarPanel)
    End Sub

    Private Sub AddGroup(sText As String)
        Dim GBox As New GroupBoxp
        GBox.Text = sText
        GBox.Height = 12
        mvarPanel.Controls.Add(GBox)
        mvarPanel.SetFlowBreak(GBox, True)
    End Sub

    Private Sub AddLabel(sText As String, Optional bBreak As Boolean = False, Optional nWidth As Integer = 20)
        Dim pL1 As New Label
        pL1.Text = sText
        pL1.TextAlign = ContentAlignment.MiddleCenter
        pL1.Width = nWidth
        pL1.Height = 24
        pL1.BorderStyle = Windows.Forms.BorderStyle.FixedSingle
        mvarPanel.Controls.Add(pL1)
        If bBreak Then
            mvarPanel.SetFlowBreak(pL1, bBreak)
        End If
    End Sub

    Private Function AddTextBox(sText As String, Optional bBreak As Boolean = False, Optional nWidth As Integer = 20) As TextBox
        Dim pL1 As New TextBox
        pL1.Text = sText
        pL1.Width = nWidth
        pL1.Height = 24
        pL1.BorderStyle = Windows.Forms.BorderStyle.FixedSingle

        mvarPanel.Controls.Add(pL1)
        If bBreak Then
            mvarPanel.SetFlowBreak(pL1, bBreak)
        End If
        Return pL1
    End Function
    Private Function AddTUpdown(nValue As Decimal, nMax As Decimal, Optional bBreak As Boolean = False, Optional nWidth As Integer = 20) As NumericUpDown
        Dim pL1 As New NumericUpDown
        pL1.Maximum = nMax
        pL1.Value = nValue
        pL1.Width = nWidth
        pL1.Height = 24
        pL1.BorderStyle = Windows.Forms.BorderStyle.FixedSingle
        pL1.DecimalPlaces = 2

        mvarPanel.Controls.Add(pL1)
        If bBreak Then
            mvarPanel.SetFlowBreak(pL1, bBreak)
        End If
        Return pL1
    End Function

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property



    Private Sub mvarNum面積1_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarNum面積1.ValueChanged
        If mvarCHK面積調整.Checked Then
            mvarNum面積2.Value = n元面積 - mvarNum面積1.Value
        End If

        CheckValue()
    End Sub

    Private Sub mvarNum面積2_ValueChanged(sender As Object, e As System.EventArgs) Handles mvarNum面積2.ValueChanged
        If mvarCHK面積調整.Checked Then
            mvarNum面積1.Value = n元面積 - mvarNum面積2.Value
        End If

        CheckValue()
    End Sub

    Private Sub CheckValue()
        If n元面積 >= mvarNum面積1.Value + mvarNum面積2.Value AndAlso mvarNum面積1.Value > 0 AndAlso mvarNum面積2.Value > 0 Then
            mvarOK.Enabled = True
        End If
    End Sub

    Private Sub mvarOK_Click(sender As Object, e As System.EventArgs) Handles mvarOK.Click
        If mvartext地番1.Text = mvartext地番2.Text Then
            MsgBox("分筆後の地番が等しいので分筆できません")
        ElseIf n元面積 >= mvarNum面積1.Value + mvarNum面積2.Value AndAlso mvarNum面積1.Value > 0 AndAlso mvarNum面積2.Value > 0 AndAlso mvartext地番1.Text.Length > 0 AndAlso mvartext地番2.Text.Length > 0 Then
            Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D:農地Info];")
            Dim p転用 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D_転用農地];")
            pTBL.Merge(p転用)
            Dim NewID As Integer = 0
            For Each pRow As DataRow In pTBL.Rows
                If pRow.Item("MinID") < NewID Then
                    NewID = pRow.Item("MinID") - 1
                End If
            Next

            Dim Rs1 As DataRow = App農地基本台帳.TBL農地.FindRowByID(n元ID)
            Dim pUpdateRecord As New RecordSQL(Rs1)
            pUpdateRecord.NewValue("地番") = mvartext地番1.Text
            pUpdateRecord.NewValue("登記簿面積") = mvarNum面積1.Value
            pUpdateRecord.NewValue("実面積") = mvarNum面積1.Value
            pUpdateRecord.NewValue("田面積") = IIf(Val(Rs1.Item("田面積").ToString) > 0, mvarNum面積1.Value, 0)
            pUpdateRecord.NewValue("畑面積") = IIf(Val(Rs1.Item("畑面積").ToString) > 0, mvarNum面積1.Value, 0)
            pUpdateRecord.NewValue("樹園地") = IIf(Val(Rs1.Item("樹園地").ToString) > 0, mvarNum面積1.Value, 0)

            If n元地番 = mvartext地番1.Text Then
                Make農地履歴(n元ID, Now, Now, 99984, enum法令.分筆登記, String.Format("[{0}]へ分筆", mvartext地番2.Text))
            Else
                Make農地履歴(n元ID, Now, Now, 99984, enum法令.分筆登記, String.Format("[{0}]より[{1}][{2}]に分筆", n元地番, mvartext地番1.Text, mvartext地番2.Text))
            End If

            Dim sSQL1 As String = pUpdateRecord.UpdateSQL("D:農地Info")

            Dim Rs2 As DataRow = App農地基本台帳.TBL農地.NewRow
            For Each pCol As DataColumn In App農地基本台帳.TBL農地.Columns
                If pCol.ColumnName = "ID" Then
                    Rs2.Item(pCol.ColumnName) = NewID
                Else
                    Rs2.Item(pCol.ColumnName) = Rs1.Item(pCol.ColumnName)
                End If
            Next



            Dim pTBLExt As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID]=0")
            Dim pCopyRecord As New RecordSQL(Rs2)
            pCopyRecord.NewValue("地番") = mvartext地番2.Text
            pCopyRecord.NewValue("登記簿面積") = mvarNum面積2.Value
            pCopyRecord.NewValue("実面積") = mvarNum面積2.Value
            pCopyRecord.NewValue("田面積") = IIf(Val(Rs1.Item("田面積").ToString) > 0, mvarNum面積2.Value, 0)
            pCopyRecord.NewValue("畑面積") = IIf(Val(Rs1.Item("畑面積").ToString) > 0, mvarNum面積2.Value, 0)
            pCopyRecord.NewValue("樹園地") = IIf(Val(Rs1.Item("樹園地").ToString) > 0, mvarNum面積2.Value, 0)

            Dim sSQL2 As String = pCopyRecord.InsertSQL("D:農地Info", pTBLExt)




            If n元地番 = mvartext地番1.Text Then
                Make農地履歴(NewID, Now, Now, 99984, enum法令.分筆登記, String.Format("[{0}]より分筆", n元地番))
            Else
                Make農地履歴(NewID, Now, Now, 99984, enum法令.分筆登記, String.Format("[{0}]より[{1}][{2}]に分筆", n元地番, mvartext地番1.Text, mvartext地番2.Text))
            End If

            SysAD.DB(sLRDB).ExecuteSQL(sSQL1)
            SysAD.DB(sLRDB).ExecuteSQL(sSQL2)
            App農地基本台帳.TBL農地.Rows.Add(Rs2)


            MsgBox("終了しました")
            Me.DoClose()
        Else


        End If
    End Sub
End Class

Public Class RecordSQL
    Private mvarRow As DataRow
    Private mvarPrimaryKey As DataColumn()
    Private mvarUpdateList As New List(Of String)
    Public Sub New(ByRef pRow As DataRow)
        mvarRow = pRow
        mvarPrimaryKey = mvarRow.Table.PrimaryKey

    End Sub

    Public Property PrimaryKey() As DataColumn()
        Get
            Return mvarPrimaryKey
        End Get
        Set(value As DataColumn())
            mvarPrimaryKey = value
        End Set
    End Property

    Public Function UpdateSQL(ByVal sTableName As String) As String
        Dim sSQL As New System.Text.StringBuilder
        Dim sValues As New List(Of String)

        For Each sField As String In mvarUpdateList
            If mvarRow.Table.Columns.Contains(sField) Then
                Dim pCol1 As DataColumn = mvarRow.Table.Columns(sField)
                Select Case pCol1.DataType.Name
                    Case "Int32", "Double", "Single", "Decimal"
                        sValues.Add("[" & pCol1.ColumnName & "]=" & Val(mvarRow.Item(pCol1.ColumnName).ToString))
                    Case "String"
                        If mvarRow.Item(pCol1.ColumnName).ToString = "" Then
                            sValues.Add("[" & pCol1.ColumnName & "]=Null")
                        Else
                            sValues.Add("[" & pCol1.ColumnName & "]='" & mvarRow.Item(pCol1.ColumnName).ToString & "'")
                        End If

                    Case Else
                        CasePrint(pCol1.DataType.Name)
                End Select

            End If
        Next
        Dim sVal As String = Join(sValues.ToArray, ",")

        sSQL.Append("UPDATE [" & sTableName & "] SET " & sVal & " WHERE [ID]=" & mvarRow.Item("ID"))
        Return sSQL.ToString
    End Function

    Public WriteOnly Property NewValue(sField As String) As Object
        Set(value As Object)
            mvarRow.Item(sField) = value
            mvarUpdateList.Add(sField)
        End Set
    End Property

    Public Function InsertSQL(ByVal sTableName As String, DistTable As DataTable) As String
        Dim sSQL As New System.Text.StringBuilder
        Dim sField As New List(Of String)
        Dim sValue As New List(Of String)

        For Each pCol1 As DataColumn In DistTable.Columns
            If mvarRow.Table.Columns.Contains(pCol1.ColumnName) Then
                sField.Add("[" & pCol1.ColumnName & "]")
                Select Case pCol1.DataType.Name
                    Case "Int32", "Int16", "Double", "Single", "Decimal"
                        sValue.Add(Val(mvarRow.Item(pCol1.ColumnName).ToString))
                    Case "String"
                        If mvarRow.Item(pCol1.ColumnName).ToString = "" Then
                            sValue.Add("Null")
                        Else
                            sValue.Add("'" & mvarRow.Item(pCol1.ColumnName).ToString & "'")
                        End If
                    Case "DateTime"
                        If mvarRow.Item(pCol1.ColumnName).ToString = "" Then
                            sValue.Add("Null")
                        Else
                            With CDate(mvarRow.Item(pCol1.ColumnName).ToString)
                                sValue.Add(String.Format("#{0}/{1}/{2}#", .Month, .Day, .Year))
                            End With
                        End If
                    Case "Boolean"
                        If mvarRow.Item(pCol1.ColumnName).ToString = "" Then
                            sValue.Add("False")
                        ElseIf CBool(mvarRow.Item(pCol1.ColumnName).ToString) = True Then
                            sValue.Add("True")
                        Else
                            sValue.Add("False")
                        End If
                    Case Else
                        CasePrint(pCol1.DataType.Name)
                End Select
            End If
            If sField.Count > 30 Then
                Exit For
            End If
        Next

        sSQL.Append("INSERT INTO [" & sTableName & "](" & Join(sField.ToArray, ",") & ")")
        sSQL.Append(" VALUES(" & Join(sValue.ToArray, ",") & ")")

        Return sSQL.ToString
    End Function

    Public Function GetNewRow(ByVal nID As Integer, ByVal ParamArray SkipField() As String) As DataRow
        Return Nothing
    End Function
End Class

Public Class GroupBoxp
    Inherits GroupBox
    WithEvents ParentC As Panel

    Public Sub New()

    End Sub

    Private Sub GroupBoxp_ParentChanged(sender As Object, e As System.EventArgs) Handles Me.ParentChanged
        ParentC = Me.Parent
    End Sub

    Private Sub ParentC_Resize(sender As Object, e As System.EventArgs) Handles ParentC.Resize
        Me.Width = ParentC.Width
    End Sub
End Class