
Public Class CTabPage分筆処理
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private mvar農地 As CObj農地
    Private Fl As FlowLayoutPanel
    Private G0 As 元農地
    Private GX As New Dictionary(Of Integer, GBox)

    Private WithEvents Numeric As NumericUpDown
    Public 本番 As String
    Public WithEvents 決定 As ToolStripButton

    Public Sub New(p農地 As CObj農地)
        MyBase.New(True, True, "分筆処理." & p農地.ID, "分筆[" & p農地.ToString & "]")
        mvar農地 = p農地
        本番 = p農地.地番
        If InStr(本番, "-") > 0 Then
            本番 = HimTools2012.StringF.Left(本番, InStr(本番, "-") - 1)
        End If

        決定 = New ToolStripButton
        決定.Text = "決定"
        Me.ToolStrip.Items.Add(決定)

        Fl = New FlowLayoutPanel
        Fl.Dock = DockStyle.Fill
        Fl.BackColor = Color.LightYellow
        Fl.AutoScroll = True
        Me.ControlPanel.Add(Fl)
        Me.AutoScroll = True

        G0 = New 元農地(p農地)
        Numeric = G0.分筆件数
        Fl.Controls.Add(G0)
        Fl.SetFlowBreak(G0, True)
        Fl.BorderStyle = Windows.Forms.BorderStyle.Fixed3D

        For n = 1 To 2
            Dim G1 As New 分筆先("分筆先" & n, IIf(n = 1, p農地.地番, 本番), p農地)
            GX.Add(n, G1)
            AddHandler G1.登記面積.ValueChanged, AddressOf 登記面積変更
            AddHandler G1.現況面積.ValueChanged, AddressOf 現況面積変更
            Fl.Controls.Add(G1)
        Next
        登記面積変更(Nothing, Nothing)
        現況面積変更(Nothing, Nothing)
        Me.Active()
    End Sub

    Public Sub 登記面積変更(sender As Object, e As System.EventArgs)
        Dim Area As Decimal = 0

        For Each G1 As 分筆先 In GX.Values
            If G1.Visible Then
                Area += Val(G1.登記面積.Value)
            End If
        Next

        G0.分筆後登記.Text = String.Format("{0:#,#.00} ㎡", Area)
        If mvar農地.登記簿面積 > 0 AndAlso (Area > mvar農地.登記簿面積 OrElse (mvar農地.登記簿面積 - Area) / mvar農地.登記簿面積 > 0.1) Then
            G0.分筆後登記.ForeColor = Color.Red
        Else
            G0.分筆後登記.ForeColor = Color.Black
        End If

    End Sub
    Public Sub 現況面積変更(sender As Object, e As System.EventArgs)
        Dim Area As Decimal = 0

        For Each G1 As 分筆先 In GX.Values
            If G1.Visible Then
                Area += Val(G1.現況面積.Value)
            End If
        Next

        G0.分筆後現況.Text = String.Format("{0:#,#.00} ㎡", Area)
        G0.分筆後現況合計 = Area

        For Each G1 As 分筆先 In GX.Values
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
                Dim G1 As New 分筆先("分筆先" & n, 本番, mvar農地)
                AddHandler G1.登記面積.ValueChanged, AddressOf 登記面積変更
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
        Public 分筆件数 As NumericUpDown
        Public 分筆後登記 As Label
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
            MyBase.New("分筆元[" & p農地.土地所在 & "]")

            AddLabel("登記地目:" & p農地.GetItem("登記簿地目名"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("登記面積(㎡):" & p農地.GetItem("登記簿面積"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("⇒分筆後:", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            分筆後登記 = AddLabel("0.00 ㎡", ContentAlignment.MiddleRight, FlowDirection.TopDown)

            AddLabel("現況地目:" & p農地.GetItem("現況地目名"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("現況面積(㎡):" & p農地.GetItem("実面積"), ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("⇒分筆後:", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            分筆後現況 = AddLabel("2.00 ㎡", ContentAlignment.MiddleRight, FlowDirection.TopDown)

            AddLabel("田面積:" & p農地.田面積, ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("畑面積:" & p農地.畑面積, ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            AddLabel("樹園地:" & p農地.樹園地, ContentAlignment.MiddleLeft, FlowDirection.TopDown)



            AddLabel("分筆数：", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            分筆件数 = AddNumeric(2, 100, FlowDirection.TopDown)
            分筆件数.Minimum = 2
            分筆件数.Maximum = 100
            分筆件数.Value = 2

            異動日.Location = New Point(分筆件数.Right + 20, 分筆件数.Top)
            Me.Controls.Add(異動日)
            異動日.Value = Now.Date
        End Sub
    End Class

    Public Class 分筆先
        Inherits GBox
        Public 地番 As TextBox
        Public WithEvents 登記面積 As NumericUpDown
        Public WithEvents 現況面積 As NumericUpDown
        Public 田面積 As NumericUpDown
        Public 畑面積 As NumericUpDown
        Public Chk同期 As CheckBox
        Public 地目振り分け As Integer
        Public NewID As Long = 0
        Public SQL As New System.Text.StringBuilder

        Public Sub New(ByVal sText As String, s本番 As String, ByRef p農地 As CObj農地)
            MyBase.New(sText)

            Dim pTop As Integer = 0

            AddLabel("新地番", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            地番 = AddTextBox(s本番 & IIf(InStr(s本番, "-") > 0, "", "-"), FlowDirection.TopDown)
            AddLabel("新登記面積", ContentAlignment.MiddleLeft, FlowDirection.LeftToRight)
            登記面積 = AddNumeric(1, p農地.登記簿面積, FlowDirection.TopDown, 2)
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

            Chk同期 = New CheckBox
            Me.Controls.Add(Chk同期)
            Chk同期.AutoSize = True
            Chk同期.Text = "↓同期"
            Chk同期.Location = New Point(登記面積.Right, 登記面積.Top)
            Chk同期.Checked = True
        End Sub

        Private Sub 登記面積_ValueChanged(sender As Object, e As System.EventArgs) Handles 登記面積.ValueChanged
            If Chk同期.Checked Then
                現況面積.Value = 登記面積.Value
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

    Private Sub 決定_Click(sender As Object, e As System.EventArgs) Handles 決定.Click
        If MsgBox("分筆を実行しますか（履歴作成の為、処理に時間がかかる場合があります。）", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            Try
                Dim s地番 As New System.Text.StringBuilder("元地[" & mvar農地.土地所在 & "]を")
                For Each pG As 分筆先 In GX.Values
                    If pG.Visible Then
                        If Not pG.地番.Text.StartsWith(本番) OrElse pG.地番.Text.EndsWith("-") OrElse Val(pG.登記面積.Text) = 0 OrElse Val(pG.現況面積.Text) = 0 Then
                            MsgBox(pG.Text & "の値が不正です", MsgBoxStyle.Critical)
                            Return
                        Else
                            s地番.Append(pG.地番.Text & "、")
                            If mvar農地.地番 = pG.地番.Text Then

                            Else
                                Dim pTBL As DataRow() = App農地基本台帳.TBL農地.FindRowBySQL(String.Format("[大字ID]={0} AND [地番]='{1}'", mvar農地.大字ID, pG.地番.Text))
                                If pTBL IsNot Nothing AndAlso pTBL.Count > 0 Then
                                    MsgBox(mvar農地.大字 & pG.地番.Text & "は既に存在します", MsgBoxStyle.Critical)
                                    Return
                                End If
                            End If
                        End If

                    End If
                Next

                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D_削除農地] WHERE [ID]=" & mvar農地.ID)
                SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_削除農地] SELECT * FROM [D:農地Info] WHERE [ID] = " & mvar農地.ID)
                SysAD.DB(sLRDB).ExecuteSQL("DELETE * FROM [D:農地Info] WHERE [ID]=" & mvar農地.ID)
                Dim s地番元 As String = HimTools2012.StringF.Left(s地番.ToString, s地番.Length - 1)
                Make農地履歴(mvar農地.ID, G0.異動日.Value, Now, 土地異動事由.分筆登記, enum法令.分筆登記, s地番元 & "に分筆")
                Dim s地番後 As String = Replace(s地番元, "を", "より")

                For Each pG As 分筆先 In GX.Values
                    If pG.Visible Then
                        pG.NewID = AddRecord(mvar農地.ID, mvar農地.土地所在)
                        SysAD.DB(sLRDB).ExecuteSQL("UPDATE [D:農地Info] SET [地番]='{0}', [登記簿面積]={1}, [実面積]={2},[田面積]={3},[畑面積]={4},[樹園地]={5} WHERE [ID]={6}", pG.地番.Text, pG.登記面積.Value, pG.現況面積.Value, pG.田面積.Value, pG.畑面積.Value, 0, pG.NewID)
                        Make農地履歴(pG.NewID, G0.異動日.Value, Now, 土地異動事由.分筆登記, enum法令.分筆登記, s地番後 & "へ分筆")
                    End If
                Next
                App農地基本台帳.TBL農地.Rows.Remove(mvar農地.Row.Body)
                MsgBox("分筆は正常に終了しました。", MsgBoxStyle.Information)
                CType(Me.Parent, TabControl).TabPages.Remove(Me)
                Me.Dispose()
            Catch ex As Exception
                MsgBox("分筆に失敗しました。")
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


