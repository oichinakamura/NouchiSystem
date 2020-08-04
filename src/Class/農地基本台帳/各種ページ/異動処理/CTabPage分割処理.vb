
Public Class CTabPage分割処理
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvar農地 As CObj農地
    Private Fl As FlowLayoutPanel
    Private G0 As 元農地
    Private GX As New Dictionary(Of Integer, GBox)

    Private WithEvents Numeric As NumericUpDown
    Public WithEvents 決定 As ToolStripButton
    Public WithEvents Btn決定 As Button

    Public Sub New(p農地 As CObj農地)
        MyBase.New(True, True, "分割処理." & p農地.ID, "一部現況分割[" & p農地.ToString & "]")
        mvar農地 = p農地

        決定 = New ToolStripButton
        決定.Text = "決定"
        Me.ToolStrip.Items.Add(決定)

        Btn決定 = New Button()
        Btn決定.Text = "決定"


        Fl = New FlowLayoutPanel
        Fl.Dock = DockStyle.Fill
        Fl.BackColor = Color.LightYellow
        Fl.AutoScroll = True
        Me.ControlPanel.Add(Fl)
        Me.AutoScroll = True

        G0 = New 元農地(p農地)
        Numeric = G0.分割件数
        Fl.Controls.Add(G0)
        Fl.SetFlowBreak(G0, True)
        Fl.BorderStyle = Windows.Forms.BorderStyle.Fixed3D

        For n = 1 To 2
            Dim G1 As New 分割先("分割先[" & HimTools2012.StringF.Right("000" & n, 3) & "]", n, Numeric.Value, mvar農地)
            GX.Add(n, G1)
            AddHandler G1.現況面積.ValueChanged, AddressOf 現況面積変更
            Fl.Controls.Add(G1)
            Btn決定.Height = G1.Height
        Next
        現況面積変更(Nothing, Nothing)

        Fl.Controls.Add(Btn決定)

        Me.Active()
    End Sub

    Public Sub 現況面積変更(sender As Object, e As System.EventArgs)
        Dim Area As Decimal = 0

        For Each G1 As 分割先 In GX.Values
            If G1.Visible Then
                Area += Val(G1.現況面積.Value)
            End If
        Next

        G0.分筆後現況.Text = String.Format("{0:#,#.00} ㎡", Area)
        G0.分筆後現況合計 = Area

        For Each G1 As 分割先 In GX.Values
            If G1.Visible Then
                G1.地目別面積設定()
            End If
        Next
        If mvar農地.実面積 > 0 AndAlso (Area > mvar農地.実面積 OrElse (mvar農地.実面積 - Area) / mvar農地.実面積 > 0.1) Then
            G0.分筆後現況.ForeColor = Color.Red
        Else
            G0.分筆後現況.ForeColor = Color.Black
        End If
    End Sub


    Private Sub Numeric_ValueChanged(sender As Object, e As System.EventArgs) Handles Numeric.ValueChanged
        For Each pN As Integer In GX.Keys
            If pN > Numeric.Value Then
                GX.Item(pN).Visible = False
            Else
                GX.Item(pN).Visible = True
            End If
        Next
        For n As Integer = 1 To Numeric.Value
            If n > GX.Count Then
                Dim G1 As New 分割先("分割先[" & HimTools2012.StringF.Right("000" & n, 3) & "]", n, Numeric.Value, mvar農地)

                AddHandler G1.現況面積.ValueChanged, AddressOf 現況面積変更
                GX.Add(n, G1)
                Fl.Controls.Add(G1)

            End If
        Next
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

    Public Class 元農地
        Inherits GBox
        Public 分割件数 As NumericUpDown
        Public 分筆後現況 As Label
        Private mvar合計 As Decimal = 0
        Public 異動日 As New DateTimePicker
        Public Property 分筆後現況合計 As Decimal
            Get
                Return mvar合計
            End Get
            Set(value As Decimal)
                mvar合計 = value
            End Set
        End Property

        Public Sub New(ByRef p農地 As CObj農地)
            MyBase.New("分割元[" & p農地.土地所在 & "]")

            AddLabel("登記地目:" & p農地.GetItem("登記簿地目名"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("登記面積(㎡):" & p農地.GetItem("登記簿面積"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)

            AddLabel("現況地目:" & p農地.GetItem("現況地目名"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("現況面積(㎡):" & p農地.GetItem("実面積"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("⇒分筆後:", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            分筆後現況 = AddLabel("2.00 ㎡", ContentAlignment.MiddleRight, FlowDirection.TopDown)

            AddLabel("田面積:" & p農地.田面積, ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("畑面積:" & p農地.畑面積, ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("樹園地:" & p農地.樹園地, ContentAlignment.MiddleLeft, FlowDirection.TopDown)

            AddLabel("分筆数：", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            分割件数 = AddNumeric(2, 100, FlowDirection.TopDown, 0)
            分割件数.Minimum = 2
            分割件数.Maximum = 100
            分割件数.Value = 2

            異動日.Location = New Point(分割件数.Right + 20, 分割件数.Top)
            Me.Controls.Add(異動日)
            異動日.Value = Now.Date
        End Sub
    End Class

    Public Class 分割先
        Inherits GBox
        Public 番号 As Integer = 0
        Public WithEvents 現況面積 As NumericUpDown
        Public 田面積 As NumericUpDown
        Public 畑面積 As NumericUpDown

        Public 地目振り分け As Integer
        Public NewID As Long = 0
        Public SQL As New System.Text.StringBuilder

        Public Sub New(ByVal sText As String, n番号 As Integer, nMax As Integer, ByRef p農地 As CObj農地)
            MyBase.New(sText)
            番号 = n番号

            AddLabel("新現況面積", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            現況面積 = AddNumeric(1, p農地.実面積, FlowDirection.TopDown, 2)

            AddLabel("新田面積", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            田面積 = AddNumeric(0, p農地.田面積, FlowDirection.TopDown, 2)
            田面積.Enabled = False
            AddLabel("新畑面積", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            畑面積 = AddNumeric(0, p農地.畑面積, FlowDirection.TopDown, 2)
            畑面積.Enabled = False

            If p農地.田面積 > 0 Then
                地目振り分け = 1
            ElseIf p農地.畑面積 > 0 Then
                地目振り分け = 2
            Else
                地目振り分け = 0
            End If
        End Sub

        Public Sub 地目別面積設定()
            Select Case 地目振り分け
                Case 1
                    田面積.Value = Me.現況面積.Value
                    畑面積.Value = 0
                Case 2
                    田面積.Value = 0
                    畑面積.Value = Me.現況面積.Value
            End Select
        End Sub
    End Class

    Public Class GBox
        Inherits GroupBox
        Public StartX As Integer = 10
        Public StartY As Integer = 0

        Public Sub New(ByVal sText As String)
            Me.Text = sText
            Me.Dock = DockStyle.Fill
            Me.AutoSize = True
            StartY = 15
        End Sub

        Public Function AddLabel(ByVal sText As String, ByVal pTextAlignment As System.Drawing.ContentAlignment, p As System.Windows.Forms.FlowDirection)
            Dim LB01 As New Label()
            LB01.Text = sText
            LB01.AutoSize = False
            LB01.TextAlign = pTextAlignment
            Me.Controls.Add(LB01)
            LB01.Location = New Point(StartX, StartY)
            Select Case p
                Case FlowDirection.LeftToRight : StartX = LB01.Right + LB01.Margin.Right
                Case FlowDirection.TopDown : StartY = LB01.Bottom + LB01.Margin.Bottom : StartX = 10
            End Select
            Return LB01
        End Function

        Public Function AddLabel(ByVal sText As String, ByVal nLeft As Integer, ByVal nTop As Integer, Optional ByVal p As System.Windows.Forms.FlowDirection = FlowDirection.LeftToRight) As Control
            Dim pCtrl As New Label()
            pCtrl.Text = sText
            Me.Controls.Add(pCtrl)
            pCtrl.Location = New Point(nLeft, nTop)
            Select Case p
                Case FlowDirection.LeftToRight : StartX = pCtrl.Right + pCtrl.Margin.Right
                Case FlowDirection.TopDown : StartY = pCtrl.Bottom + pCtrl.Margin.Bottom : StartX = 10
            End Select
            Return pCtrl
        End Function

        Public Function AddTextBox(ByVal sText As String, ByVal p As System.Windows.Forms.FlowDirection) As TextBox
            Dim pCtrl As New TextBox
            Me.Controls.Add(pCtrl)
            pCtrl.Location = New Point(StartX, StartY)
            pCtrl.Text = sText
            Select Case p
                Case FlowDirection.LeftToRight : StartX = pCtrl.Right + pCtrl.Margin.Right
                Case FlowDirection.TopDown : StartY = pCtrl.Bottom + pCtrl.Margin.Bottom : StartX = 10
            End Select

            Return pCtrl
        End Function

        Public Function AddNumeric(ByVal nValue As Decimal, nMax As Decimal, ByVal p As System.Windows.Forms.FlowDirection, Optional pDecimalPlaces As Integer = 0) As NumericUpDown
            Dim pCtrl As New NumericUpDown
            pCtrl.DecimalPlaces = pDecimalPlaces
            pCtrl.Maximum = nMax
            Me.Controls.Add(pCtrl)
            pCtrl.Location = New Point(StartX, StartY)
            pCtrl.Value = nValue
            pCtrl.TextAlign = HorizontalAlignment.Right
            Select Case p
                Case FlowDirection.LeftToRight : StartX = pCtrl.Right + pCtrl.Margin.Right
                Case FlowDirection.TopDown : StartY = pCtrl.Bottom + pCtrl.Margin.Bottom : StartX = 10
            End Select

            Return pCtrl
        End Function

    End Class

    Private Sub 決定_Click(sender As Object, e As System.EventArgs) Handles 決定.Click, Btn決定.Click
        If MsgBox("分筆を実行しますか（履歴作成の為、処理に時間がかかる場合があります。）", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try

                Dim s地番 As New System.Text.StringBuilder("元地[" & mvar農地.土地所在 & "]を")
                For Each pG As 分割先 In GX.Values
                    If pG.Visible Then
                        If Not pG.現況面積.Value > 0 Then
                            MsgBox(pG.Text & "の値が不正です", MsgBoxStyle.Critical)
                            Return
                        Else
                            s地番.Append(mvar農地.地番 & "(" & pG.番号 & ")、")

                        End If
                    End If
                Next

                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D_削除農地] WHERE [ID]=" & mvar農地.ID)
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_削除農地] SELECT * FROM [D:農地Info] WHERE [ID] = " & mvar農地.ID)
                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D:農地Info] WHERE [ID]=" & mvar農地.ID)
                Dim s地番元 As String = HimTools2012.StringF.Left(s地番.ToString, s地番.Length - 1)

                Make農地履歴(mvar農地.ID, G0.異動日.Value, Now, 土地異動事由.一部現況分割, enum法令.その他分割処理, s地番元 & "に分割")
                Dim s地番後 As String = Replace(s地番元, "を", "より")

                For Each pG As 分割先 In GX.Values
                    If pG.Visible Then
                        pG.NewID = AddRecord(mvar農地.ID, mvar農地.土地所在)
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [一部現況]={0}, [実面積]={1},[田面積]={2},[畑面積]={3},[樹園地]={4},[部分面積]={5} WHERE [ID]={6}", pG.番号, pG.現況面積.Value, pG.田面積.Value, pG.畑面積.Value, 0, pG.現況面積.Value, pG.NewID)
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [元ID]='{0}' WHERE [ID]={1}", mvar農地.ID, pG.NewID)
                        Make農地履歴(pG.NewID, G0.異動日.Value, Now, 土地異動事由.一部現況分割, enum法令.その他分割処理, s地番後 & "へ分割")
                    End If
                Next
                App農地基本台帳.TBL農地.Rows.Remove(mvar農地.Row.Body)

                MsgBox("部分分割は正常に終了しました。", MsgBoxStyle.Information)
                CType(Me.Parent, TabControl).TabPages.Remove(Me)
                Me.Dispose()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Function AddRecord(ByVal n元ID As Long, ByVal s元所在 As String) As Long
        Dim pNewRow As DataRow = App農地基本台帳.TBL農地.NewRow
        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D:農地Info];")
        Dim p転用 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D_転用農地];")
        Dim p削除 As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT Min([ID]) AS MinID FROM [D_削除農地];")

        If pTBL.Rows.Count > 0 Then
            Dim nID As Integer = Val(pTBL.Rows(0).Item("MinID").ToString) - 1
            If nID >= 0 Then
                nID = -1
            End If
            If p転用.Rows.Count > 0 Then
                Dim TID As Integer = Val(p転用.Rows(0).Item("MinID").ToString) - 1
                If nID > TID Then
                    nID = TID
                End If
            End If
            If p削除.Rows.Count > 0 Then
                Dim DID As Integer = Val(p削除.Rows(0).Item("MinID").ToString) - 1
                If nID > DID Then
                    nID = DID
                End If
            End If

            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D:農地Info] SELECT * FROM [D_削除農地] WHERE [ID] = " & mvar農地.ID)
            Do Until Replace(SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [ID]=" & nID & " WHERE [ID]=" & mvar農地.ID), "OK", "") = ""
                nID = nID - 1
            Loop
            SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地系図]([自ID],[元ID],[元土地所在]) VALUES({0},{1},'{2}')", nID, n元ID, s元所在)
            Return nID
        End If
        Return 0
    End Function
End Class

Public Class CTabPage部分結合
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvar農地 As CObj農地
    Private S1 As SplitContainer
    Private Fl As FlowLayoutPanel
    Private WithEvents mvarGrid As New HimTools2012.controls.DataGridViewWithDataView

    Public WithEvents 決定 As ToolStripButton
    Public WithEvents Btn決定 As Button
    Public 異動日 As New DateTimePicker

    Public Sub New(p農地 As CObj農地)
        MyBase.New(True, True, "部分結合." & p農地.ID, "一部現況⇒結合処理[" & p農地.ToString & "]")
        決定 = New ToolStripButton
        決定.Text = "決定"
        Me.ToolStrip.Items.Add(決定)
        mvar農地 = p農地
        Btn決定 = New Button()
        Btn決定.Text = "決定"

        S1 = New SplitContainer
        S1.Dock = DockStyle.Fill
        S1.BackColor = Color.LightYellow
        S1.Orientation = Orientation.Horizontal


        Me.ControlPanel.Add(S1)
        Me.AutoScroll = True

        Dim sWhere As String = ""
        If p農地.所在.Length > 0 Then
            sWhere = String.Format("[所在]='{0}' AND [地番]='{1}'", p農地.所在, p農地.地番)
        Else
            sWhere = String.Format("[大字ID]={0} AND [地番]='{1}'", p農地.大字ID, p農地.地番)
        End If

        mvarGrid = New HimTools2012.controls.DataGridViewWithDataView()
        mvarGrid.AutoGenerateColumns = False
        mvarGrid.Dock = DockStyle.Fill

        mvarGrid.AddColumnText("基準筆", "基準筆", "基準筆", HimTools2012.enumReadOnly.bReadOnly)
        For Each sColumn As String In {"ID", "土地所在", "一部現況", "登記簿面積", "実面積", "登記簿地目", "所有者ID", "所有者氏名"}
            mvarGrid.AddColumnText(sColumn, sColumn, sColumn, HimTools2012.enumReadOnly.bReadOnly)
        Next

        Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE " & sWhere)
        App農地基本台帳.TBL農地.MergePlus(pTBL)


        S1.Panel1.Controls.Add(mvarGrid)

        Fl = New FlowLayoutPanel
        Fl.Dock = DockStyle.Fill
        Fl.BackColor = Color.LightYellow

        S1.Panel2.Controls.Add(Fl)
        Fl.Controls.Add(異動日)
        Fl.Controls.Add(Btn決定)

        mvarGrid.SetDataView(App農地基本台帳.TBL農地.Body, sWhere, "")
    End Sub

    Private Sub mvarGrid_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles mvarGrid.Paint
        For Each pRow As DataGridViewRow In mvarGrid.Rows
            If CType(pRow.DataBoundItem, DataRowView).Item("ID") = mvar農地.ID Then
                pRow.Cells("基準筆").Value = "○"
            End If
        Next
    End Sub

    Private Sub 決定_Click(sender As Object, e As System.EventArgs) Handles 決定.Click, Btn決定.Click
        If MsgBox("部分分割を結合しますか（履歴作成の為、処理に時間がかかる場合があります。）", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
                Dim s地番 As New List(Of String)
                Dim sID As New List(Of String)
                For Each pRowG As DataGridViewRow In mvarGrid.Rows
                    If pRowG.Cells("基準筆").Value = "○" Then
                    Else
                        Dim s元 As String = String.Format("{0}({1})", pRowG.Cells("土地所在").Value, pRowG.Cells("一部現況").Value)
                        s地番.Add(s元)
                        sID.Add(pRowG.Cells("ID").Value)
                        Make農地履歴(pRowG.Cells("ID").Value, 異動日.Value, Now, 土地異動事由.一部現況結合, enum法令.その他分割統合, mvar農地.土地所在 & "(" & mvar農地.Row.Item("一部現況") & ")へ結合")
                        AddRecord(pRowG.Cells("ID").Value, mvar農地.ID, s元)
                    End If
                Next

                For Each sX As String In sID
                    Dim pRow As DataRow = App農地基本台帳.TBL農地.Body.Rows.Find(Val(sX))
                    If pRow IsNot Nothing Then
                        App農地基本台帳.TBL農地.Body.Rows.Remove(pRow)
                    End If
                Next

                Make農地履歴(mvar農地.ID, 異動日.Value, Now, 土地異動事由.一部現況結合, enum法令.その他分割統合, "[" & Join(s地番.ToArray(), "、") & "]を結合")


                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D_削除農地] WHERE [ID] IN ({0})", Join(sID.ToArray(), ","))
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_削除農地] SELECT * FROM [D:農地Info] WHERE [ID] IN ({0})", Join(sID.ToArray(), ","))
                '/*20161027 条件フィールド【ID】の追加*/
                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D:農地Info] WHERE [ID] IN ({0})", Join(sID.ToArray(), ","))

                Dim n登記簿面積 As Decimal = mvar農地.GetDecimalValue("登記簿面積")
                Dim UpdateSQL As String = "[一部現況]=0,[実面積]=" & n登記簿面積

                mvar農地.SetIntegerValue("一部現況", 0)
                mvar農地.SetDecimalValue("実面積", n登記簿面積)
                If Val(mvar農地.Row.Item("田面積")) > 0 Then
                    UpdateSQL &= ",[田面積]=" & n登記簿面積 & ",[畑面積]=0"
                    mvar農地.SetDecimalValue("田面積", n登記簿面積)
                    mvar農地.SetDecimalValue("畑面積", 0)
                ElseIf Val(mvar農地.Row.Item("畑面積")) > 0 Then
                    UpdateSQL &= ",[田面積]=0, [畑面積]=" & n登記簿面積
                    mvar農地.SetDecimalValue("田面積", 0)
                    mvar農地.SetDecimalValue("畑面積", n登記簿面積)
                Else
                    UpdateSQL &= ",[田面積]=0, [畑面積]=0"
                    mvar農地.SetDecimalValue("田面積", 0)
                    mvar農地.SetDecimalValue("畑面積", n登記簿面積)
                End If

                SysAD.DB(sLRDB).ExecuteSQL("Update [D:農地Info] SET {0} WHERE [ID]={1}", UpdateSQL, mvar農地.ID)


                MsgBox("結合処理は正常に終了しました。", MsgBoxStyle.Information)
                CType(Me.Parent, TabControl).TabPages.Remove(Me)
                Me.Dispose()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Function AddRecord(ByVal n元ID As Long, ByVal n先ID As Long, ByVal s元所在 As String) As Long



        SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地系図]([自ID],[元ID],[元土地所在]) VALUES({0},{1},'{2}')", n先ID, n元ID, s元所在)

        Return 0
    End Function

End Class


