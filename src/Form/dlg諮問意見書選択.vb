Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing.Design

Public Class dlg諮問意見書選択
    Implements HimTools2012.controls.XMLLayoutContainer
    Implements HimTools2012.controls.XMLCtrlParent



    Public WithEvents mvarTab As TabControl

    Public DT開始年月日 As HimTools2012.controls.ToolStripDateTimePicker
    Public DT終了年月日 As HimTools2012.controls.ToolStripDateTimePicker

    Private WithEvents mvarXMLLayout As HimTools2012.controls.XMLLayout

    Public Sub New()
        InitializeComponent()

        With Me
            .SuspendLayout()
            .Text = "諮問意見書選択"
            .Width = 9990 / 15
            .Height = 7695 / 15

            mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
            With mvarXMLLayout
                .StartLayout(My.Resources.Resource1.基本画面, "諮問意見書選択")

                DT開始年月日 = .Controls("DT開始年月日")
                DT終了年月日 = .Controls("DT終了年月日")


                mvarTab = .Controls("mvarTab")
                mvarTab.DrawMode = TabDrawMode.OwnerDrawFixed
                mvarTab.Multiline = True
                mvarTab.Appearance = TabAppearance.Normal
            End With
            .ResumeLayout()
        End With
    End Sub

    Public Sub Load諮問意見書資料()
        Data初期設定()
        mvarXMLLayout.Controls("txt対象年月").Text = Strings.Right("0000" & Year(DT終了年月日.Value), 4) & Strings.Right("00" & Month(DT終了年月日.Value), 2)
        SetTabData()
    End Sub

    Private Sub SetTabData()
        mvarTab.TabPages.Clear()
        ' 受付中・審査中・許可済み
        'Dim sWhere As String = String.Format("([状態]=0 OR [状態]=1 OR [状態]=2) AND [受付年月日]>=#{1}/{2}/{0}# AND [受付年月日]<=#{4}/{5}/{3}#",
        '        DT開始年月日.Value.Year, DT開始年月日.Value.Month, DT開始年月日.Value.Day,
        '        DT終了年月日.Value.Year, DT終了年月日.Value.Month, DT終了年月日.Value.Day)
        Dim sWhere As String = String.Format("[状態]=0 AND [受付年月日]>=#{1}/{2}/{0}# AND [受付年月日]<=#{4}/{5}/{3}#",
                DT開始年月日.Value.Year, DT開始年月日.Value.Month, DT開始年月日.Value.Day,
                DT終了年月日.Value.Year, DT終了年月日.Value.Month, DT終了年月日.Value.Day)

        Dim pTableA As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請 WHERE {0}", sWhere)
        App農地基本台帳.TBL申請.MergePlus(pTableA)

        Dim pTable As DataTable = App農地基本台帳.TBL申請.ToDataView(sWhere, "法令,受付番号").ToTable
        pTable.Columns.Add(New DataColumn("選択", GetType(Boolean)))
        pTable.Columns.Add(New DataColumn("諮問番号", GetType(Integer)))
        pTable.Columns.Add(New DataColumn("総面積", GetType(Decimal)))

        For Each pRow As DataRow In pTable.Rows
            pRow.Item("選択") = True
        Next

        For Each pRow As DataRow In pTable.Rows
            Dim nH As enum法令 = pRow.Item("法令")
            Select Case nH
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    nH = 50
            End Select

            Dim pTab As 諮問意見書Page
            If Not mvarTab.TabPages.ContainsKey("n." & nH) Then
                Dim St As String = ""
                Dim sTag As String = ""
                Select Case nH
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用 : St = "4条" : sTag = "[法令]=40 Or [法令]=42"
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : St = "5条" : sTag = "[法令]=50 Or [法令]=51 Or [法令]=52"
                        'Case enum法令.非農地証明願, 600 : St = "非農地証明願"
                    Case Else
                End Select
                If Len(St) Then
                    pTab = New 諮問意見書Page
                    pTab.ImageKey = "OK"

                    pTab.Name = "n." & nH
                    pTab.Text = "□" & St
                    If Len(sTag) > 0 Then
                        pTab.Tag = sTag
                    Else
                        pTab.Tag = "[法令]=" & nH
                    End If
                    mvarTab.TabPages.Add(pTab)
                    pTab.List.DataSource = New DataView(pTable, pTab.Tag, "受付番号", DataViewRowState.CurrentRows)
                End If
            Else
                pTab = mvarTab.TabPages.Item("n." & nH)
            End If
        Next
    End Sub

#Region "Legucy"


    Private Sub Data初期設定()

        With New 議案書作成パラメータ
            DT終了年月日.Value = .終了年月日
            DT開始年月日.Value = .開始年月日
        End With
    End Sub

    '////////////////////////////////////////////////////////////////////////
    Private Function GetDt(ByVal sKey As String, ByVal MaxDay As Long) As String
        Dim sDT As String

        sDT = SysAD.DB(sLRDB).DBProperty(sKey)
        Do Until Val(sDT) > 0 And Val(sDT) < MaxDay
            sDT = InputBox(sKey & "日を入力してください", sKey & "日", 15, 1)

            If Len(sDT) = 0 Then
                Me.Hide()
            ElseIf Val(sDT) > 0 And Val(sDT) < MaxDay Then
                SysAD.DB(sLRDB).DBProperty(sKey) = Val(sDT)
            End If
        Loop
        GetDt = sDT
    End Function
#End Region

    Private Sub mvarTab_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles mvarTab.DrawItem
        '対象のTabControlを取得
        Dim tab As TabControl = CType(sender, TabControl)
        Dim txt As String = tab.TabPages(e.Index).Text

        'タブのテキストと背景を描画するためのブラシを決定する
        Dim foreBrush As Brush
        With CType(mvarTab.TabPages(e.Index), 諮問意見書Page)

            If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
                '選択されているタブのテキストを赤、背景を青とする
                foreBrush = Brushes.Red
            Else
                '選択されていないタブのテキストは灰色、背景を白とする
                foreBrush = Brushes.Navy
            End If
            e.DrawBackground()

            'StringFormatを作成
            Dim sf As New StringFormat
            '中央に表示する
            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center

            If .印刷 Then
                txt = Replace(txt, "□", "☑")
            End If

            'Textの描画
            e.Graphics.DrawString(txt, e.Font, foreBrush, RectangleF.op_Implicit(e.Bounds), sf)
        End With

    End Sub

    Public ReadOnly Property BottomToolStripPanel As System.Windows.Forms.ToolStripPanel Implements HimTools2012.controls.XMLLayoutContainer.BottomToolStripPanel
        Get
            Return mvarXMLLayout.Controls("ToolStripContainer1")
        End Get
    End Property

    Public ReadOnly Property Controls1 As System.Windows.Forms.Control.ControlCollection Implements HimTools2012.controls.XMLLayoutContainer.Controls
        Get
            Return Me.Controls
        End Get
    End Property

    Public Sub EventMan(ByVal s As Object, ByVal e As System.EventArgs) Implements HimTools2012.controls.XMLLayoutContainer.EventMan
        Select Case TypeName(e)

            Case "ClickEvent"

                Select Case s.name


                    Case Else
                        Stop
                End Select

            Case Else
                Stop
        End Select
    End Sub

    Public ReadOnly Property ToolStrip As HimTools2012.controls.ToolStripEx Implements HimTools2012.controls.XMLLayoutContainer.ToolStrip
        Get
            Return mvarXMLLayout.Controls("ToolStrip1")
        End Get
    End Property

    Private Sub mvarXMLLayout_ClickButton(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarXMLLayout.ClickButton
        If TypeOf sender Is HimTools2012.controls.XMLCtrl Then

            Select Case CType(sender, HimTools2012.controls.XMLCtrl).Key
                Case "Btn再読込" : SetTabData()
                Case "btn全議案選択"
                    For Each pPage As 諮問意見書Page In mvarTab.TabPages
                        pPage.印刷 = True
                    Next
                    mvarTab.Refresh()
                Case "btn全議案解除"
                    For Each pPage As 諮問意見書Page In mvarTab.TabPages
                        pPage.印刷 = False
                    Next
                    mvarTab.Refresh()
                Case "OK"
                    Dim X As New C諮問意見書Data作成(mvarTab)
                    X.Dialog.StartProc(True, False)

                    Me.DialogResult = System.Windows.Forms.DialogResult.OK
                    Me.Close()
                Case "Cancel"
                    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
                    Me.Close()
                Case Else
                    Stop
            End Select

        End If
    End Sub

    Public Sub SetXMLParam(ByVal pPNode As System.Xml.XmlNode, ByVal pLayout As HimTools2012.controls.XMLLayout) Implements HimTools2012.controls.XMLLayoutContainer.SetXMLParam
        For Each pAttr As Xml.XmlAttribute In pPNode.Attributes
            Select Case pAttr.Name
                Case "Name"
                Case "CloseMode"
            End Select
        Next
        pLayout.ChildLayout(pPNode, Me)
    End Sub

    Public Sub AddCtrl(ByVal pCtrl As Object) Implements HimTools2012.controls.XMLCtrlParent.AddCtrl

        Me.Controls.Add(pCtrl)
    End Sub
End Class

Public Class 諮問意見書Page
    Inherits TabPage
    Private WithEvents mvarPanel As SplitContainer
    Private WithEvents mvarList As DataGridView
    Private WithEvents mvar全選択 As New Button()
    Private WithEvents mvar全解除 As New Button()

    Public Property 印刷 As Boolean
        Get
            Return mvar印刷.Checked
        End Get
        Set(ByVal value As Boolean)
            mvar印刷.Checked = value
        End Set
    End Property
    Private WithEvents mvar印刷 As CheckBox

    Public Sub New()
        mvarPanel = New SplitContainer
        Me.Controls.Add(mvarPanel)
        With mvarPanel
            .Dock = DockStyle.Fill
            .BackColor = SystemColors.Window
            .Orientation = Orientation.Horizontal
            .Panel1MinSize = 60
            .SplitterDistance = 60
            .FixedPanel = FixedPanel.Panel1
            .IsSplitterFixed = True
        End With

        With AddGroup(mvarPanel.Panel1, "印刷順序")
            mvar印刷 = New CheckBox
            mvar印刷.AutoSize = True
            mvar印刷.Text = "印刷"
            mvar印刷.Checked = True
            mvar印刷.Location = New Point(0, 0)
            .Controls.Add(mvar印刷)
            mvar全選択.Text = "全選択"
            mvar全選択.Location = New Point(0, mvar印刷.Width + 5)
            .Controls.Add(mvar全選択)
            mvar全解除.Text = "全解除"
            mvar全解除.Location = New Point(0, mvar全選択.Left + mvar全選択.Width + 5)
            .Controls.Add(mvar全解除)
        End With

        mvarList = New DataGridView
        mvarPanel.Panel2.Controls.Add(mvarList)
        With mvarList
            .AllowUserToAddRows = False
            .Dock = DockStyle.Fill
            .AutoGenerateColumns = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        End With

        Dim pCol As New DataGridViewCheckBoxColumn
        pCol.Name = "選択"
        pCol.DataPropertyName = "選択"
        mvarList.Columns.Add(pCol)

        Dim pColn As New DataGridViewTextBoxColumn
        pColn.DataPropertyName = "名称"
        pColn.HeaderText = "件名"
        mvarList.Columns.Add(pColn)

        Dim pColb As New DataGridViewTextBoxColumn
        pColb.DataPropertyName = "受付番号"
        pColb.HeaderText = "受付番号"
        mvarList.Columns.Add(pColb)

        
    End Sub

    Private Function AddGroup(ByVal pPanel As Control, ByVal sText As String) As FloatGroup
        Dim pGroupBox As New FloatGroup(pPanel)
        pGroupBox.Text = sText

        Return pGroupBox
    End Function

    Public ReadOnly Property List As DataGridView
        Get
            Return mvarList
        End Get
    End Property

    Public Function SelectCount() As Integer
        Dim pCount As Integer = 0
        For Each pRow As System.Windows.Forms.DataGridViewRow In mvarList.Rows
            If Not IsDBNull(pRow.Cells("選択").Value) AndAlso pRow.Cells("選択").Value = True Then
                pCount += 1
            End If
        Next

        Return pCount
    End Function

    Private Class FloatGroup
        Inherits GroupBox
        Private mvarInnerPanel As FlowLayoutPanel

        Public Sub New(ByVal pPanel As Control)
            Me.Dock = DockStyle.Fill
            pPanel.Controls.Add(Me)
            mvarInnerPanel = New FlowLayoutPanel
            mvarInnerPanel.Dock = DockStyle.Fill
            MyBase.Controls.Add(mvarInnerPanel)
        End Sub
        Public Shadows ReadOnly Property Controls As ControlCollection
            Get
                Return mvarInnerPanel.Controls
            End Get
        End Property
        Public Sub SetFlowBreak(ByVal pCtrl As Control)
            mvarInnerPanel.SetFlowBreak(pCtrl, True)
        End Sub
    End Class

    Private Sub mvar印刷_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar印刷.CheckedChanged
        If Me.Parent IsNot Nothing Then
            Me.Parent.Refresh()
        End If
    End Sub

    Private Sub mvar全選択_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全選択.Click
        For Each pRow As DataRowView In mvarList.DataSource
            pRow.Item("選択") = True
        Next
        mvarList.Refresh()
    End Sub
    Private Sub mvar全解除_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar全解除.Click
        For Each pRow As DataRowView In mvarList.DataSource
            pRow.Item("選択") = False
        Next
        mvarList.Refresh()
    End Sub
End Class



