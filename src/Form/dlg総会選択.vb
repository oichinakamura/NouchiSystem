Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Drawing.Design


Public Class dlg総会選択
    Implements HimTools2012.controls.XMLLayoutContainer
    Implements HimTools2012.controls.XMLCtrlParent



    Public WithEvents mvarTab As TabControl

    Public DT開始年月日 As HimTools2012.controls.ToolStripDateTimePicker
    Public DT終了年月日 As HimTools2012.controls.ToolStripDateTimePicker

    Private WithEvents mvarXMLLayout As HimTools2012.controls.XMLLayout
    Private mvar総会資料作成DT As New C総会資料作成

    Public Sub New()
        InitializeComponent()

        With Me
            .SuspendLayout()
            .Text = "議案選択"
            .Width = 9990 / 15
            .Height = 7695 / 15

            mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
            With mvarXMLLayout
                .StartLayout(My.Resources.Resource1.基本画面, "議案選択")

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

    Public Sub Load総会資料()
        CType(mvarXMLLayout.Controls("mvarPropertyG"), PropertyGrid).SelectedObject = mvar総会資料作成DT
        Data初期設定()
        mvarXMLLayout.Controls("txt対象年月").Text = Strings.Right("0000" & Year(DT終了年月日.Value), 4) & Strings.Right("00" & Month(DT終了年月日.Value), 2)
        SetTabData()
    End Sub

    Private Sub SetTabData()
        mvarTab.TabPages.Clear()
        ' 23:59:59
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
                Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                    nH = 30
                Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                    nH = 50
                Case Else
                    If Val(pRow.Item("経由法人ID").ToString) > 0 AndAlso SysAD.市町村.市町村名 = "南種子町" Then
                        nH = 65
                    End If
            End Select

            Dim pTab As 申請Page
            If Not mvarTab.TabPages.ContainsKey("n." & nH) Then
                Dim St As String = ""
                Dim sTag As String = ""
                Select Case nH
                    Case enum法令.農地法3条所有権, enum法令.農地法3条所有権 : St = "3条" : sTag = "[法令]=30 Or [法令]=31"
                    Case enum法令.基盤強化法所有権 : St = "所有権設定"
                    Case enum法令.利用権設定 : St = "利用権設定"
                    Case enum法令.利用権移転 : St = "利用権移転"
                    Case enum法令.合意解約 : St = "合意解約"
                    Case enum法令.農地法18条解約, enum法令.農地法20条解約 : St = "18条解約"
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用 : St = "4条" : sTag = "[法令]=40 Or [法令]=42"
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : St = "5条" : sTag = "[法令]=50 Or [法令]=51 Or [法令]=52"
                    Case enum法令.農用地計画変更 : St = "農用地計画変更"
                    Case enum法令.事業計画変更 : St = "事業計画変更"
                    Case enum法令.農地利用目的変更 : St = "農地利用目的変更"

                    Case enum法令.非農地証明願 : St = "非農地"
                    Case enum法令.奨励金交付A, enum法令.奨励金交付B : St = "奨励金"
                    Case enum法令.買受適格耕公, enum法令.買受適格耕競, enum法令.買受適格転公, enum法令.買受適格転競
                        St = Choose(nH Mod 100, "買受耕公", "買受耕競", "買受転公", "買受転競")

                    Case enum法令.あっせん出手 : St = "あっせん(出)"
                    Case enum法令.あっせん受手 : St = "あっせん(受)"
                    Case enum法令.農地改良届 : St = "農地改良届"
                    Case enum法令.事業計画変更 : St = "転用事業計画変更"
                    Case enum法令.非農地証明願 : St = "非農地証明願"
                    Case enum法令.中間管理機構経由 : St = "中間管理機構経由"
                    Case Else
                End Select
                If Len(St) Then
                    pTab = New 申請Page
                    pTab.ImageKey = "OK"

                    pTab.Name = "n." & nH
                    pTab.Text = "□" & St
                    If Len(sTag) > 0 Then
                        pTab.Tag = sTag
                    Else
                        If SysAD.市町村.市町村名 = "南種子町" Then
                            Select Case nH
                                Case enum法令.利用権設定
                                    pTab.Tag = "[法令]=" & pRow.Item("法令") & " AND ([経由法人ID] Is Null Or [経由法人ID]=0 Or [申請者A]=[経由法人ID])"
                                Case enum法令.中間管理機構経由
                                    pTab.Tag = "[法令]=" & pRow.Item("法令") & " AND [経由法人ID] IS NOT NULL AND [経由法人ID]<>0 AND [申請者A]<>[経由法人ID]"
                                Case Else
                                    pTab.Tag = "[法令]=" & nH
                            End Select
                        Else
                            pTab.Tag = "[法令]=" & nH
                        End If
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
            mvar総会資料作成DT.総会日 = .総会日
        End With


        'If IsDate(SysAD.DB(sLRDB).DBProperty("総会開始時間")) Then
        '    '            mvarDlg.Controls("dt総会開始時間").Value = CDate(SysAD.DB(sLRDB).DBProperty("総会開始時間"))
        'Else
        '    '            mvarDlg.Controls("dt総会開始時間").Value = CDate("9:00:00")
        'End If


        ''        mvarDlg.Controls("Chk日程").Value = Val(SysAD.DB(sLRDB).DBProperty("議案の日程番号使用"))

        'CType(Me.Controls("txt発行番号"), TextBox).Text = Val(SysAD.DB(sLRDB).DBProperty("議案の発行番号"))
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



    Private Sub mvarTab_DrawItem(sender As Object, e As System.Windows.Forms.DrawItemEventArgs) Handles mvarTab.DrawItem
        '対象のTabControlを取得
        Dim tab As TabControl = CType(sender, TabControl)
        Dim txt As String = tab.TabPages(e.Index).Text

        'タブのテキストと背景を描画するためのブラシを決定する
        Dim foreBrush As Brush
        With CType(mvarTab.TabPages(e.Index), 申請Page)

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


    Public Sub EventMan(s As Object, e As System.EventArgs) Implements HimTools2012.controls.XMLLayoutContainer.EventMan
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

    Private Sub mvarXMLLayout_ClickButton(sender As Object, e As System.EventArgs) Handles mvarXMLLayout.ClickButton
        If TypeOf sender Is HimTools2012.controls.XMLCtrl Then

            Select Case CType(sender, HimTools2012.controls.XMLCtrl).Key
                Case "Btn再読込" : SetTabData()
                Case "btn全議案選択"
                    For Each pPage As 申請Page In mvarTab.TabPages
                        pPage.印刷 = True
                    Next
                    mvarTab.Refresh()
                Case "btn全議案解除"
                    For Each pPage As 申請Page In mvarTab.TabPages
                        pPage.印刷 = False
                    Next
                    mvarTab.Refresh()
                Case "OK"
                    Dim X As New C総会資料Data作成(mvarTab)
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

    Public Sub SetXMLParam(pPNode As System.Xml.XmlNode, pLayout As HimTools2012.controls.XMLLayout) Implements HimTools2012.controls.XMLLayoutContainer.SetXMLParam
        For Each pAttr As Xml.XmlAttribute In pPNode.Attributes
            Select Case pAttr.Name
                Case "Name"
                Case "CloseMode"
                    'Switch(pAttr.Value)
                    '                {
                    '                    case "": break;
                    '                    case "CloseOK":
                    '                        this.SetCloseMode(HimTools2012.controls.CloseMode.CloseOK);
                    '                        break;
                    '                    default:
                    '                        break;
                    '                }
                    '                break;
                    '            default:
                    '                break;
            End Select

        Next
        pLayout.ChildLayout(pPNode, Me)
    End Sub

    Public Sub AddCtrl(pCtrl As Object) Implements HimTools2012.controls.XMLCtrlParent.AddCtrl

        Me.Controls.Add(pCtrl)
    End Sub
End Class




Public Class C総会資料作成
    Private mvar保存パス As String
    Public Sub New()
        Dim sPath As String = SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "")
        If IO.Directory.Exists(sPath) Then
            mvar保存パス = sPath
        Else
            mvar保存パス = SysAD.OutputFolder
        End If
    End Sub

    <Category("01_総会関連")>
    Public Property 総会日 As DateTime
    <Category("01_総会関連")>
    Public Property 発行番号 As Integer
    <Category("01_総会関連")>
    Public Property 総会開始時間 As TimeSpan

    <Category("09_総会関連")> <Editor(GetType(FolderNameEditor), GetType(System.Drawing.Design.UITypeEditor))>
    Public Property 総会資料作成場所 As String
        Get
            Return mvar保存パス
        End Get
        Set(value As String)
            If IO.Directory.Exists(value) Then
                SysAD.SetXMLProperty("総会資料関連", "出力先フォルダ", value)
                mvar保存パス = value
            End If
            mvar保存パス = value
        End Set
    End Property


    Public Class PropertyGridEditorAttribute
        Inherits System.Drawing.Design.UITypeEditor

        Public Overrides Function GetEditStyle(context As System.ComponentModel.ITypeDescriptorContext) As System.Drawing.Design.UITypeEditorEditStyle
            Return Drawing.Design.UITypeEditorEditStyle.Modal
        End Function

        Public Overrides Function EditValue(context As System.ComponentModel.ITypeDescriptorContext, provider As System.IServiceProvider, value As Object) As Object
            Return MyBase.EditValue(context, provider, value)
        End Function

    End Class

    Public Class FolderNameEditor
        Inherits System.Drawing.Design.UITypeEditor

        Private folderBrowser As FolderBrowserDialog

        Public Sub New()

            MyBase.New()
        End Sub

        Public Overrides Function EditValue(context As ITypeDescriptorContext, provider As IServiceProvider, value As Object) As Object

            If folderBrowser Is Nothing Then
                folderBrowser = New FolderBrowserDialog()
                InitializeDialog(folderBrowser)
            End If


            If TypeOf value Is String Then
                folderBrowser.SelectedPath = value.ToString()
            End If
            If folderBrowser.ShowDialog() <> DialogResult.OK Then
                Return value
            End If

            Return folderBrowser.SelectedPath
        End Function

        Public Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
            Return UITypeEditorEditStyle.Modal
        End Function

        Protected Overridable Sub InitializeDialog(browserDialog As FolderBrowserDialog)
            browserDialog.Description = "対象とするフォルダへのパスを選択してください。"
        End Sub
    End Class


End Class


Public Class 申請Page
    Inherits TabPage
    Private WithEvents mvarPanel As SplitContainer
    Private WithEvents mvarList As DataGridView
    Private WithEvents mvar全選択 As New Button()
    Private WithEvents mvar全解除 As New Button()

    Public txt議案番号 As TextBox
    Public txt日程番号 As TextBox
    Public Property 印刷 As Boolean
        Get
            Return mvar印刷.Checked
        End Get
        Set(value As Boolean)
            mvar印刷.Checked = value
        End Set
    End Property
    Private WithEvents mvar印刷 As CheckBox

    Public Sub New()
        mvarPanel = New SplitContainer
        mvarPanel.Dock = DockStyle.Fill
        mvarPanel.BackColor = SystemColors.Window

        Me.Controls.Add(mvarPanel)
        mvarPanel.Orientation = Orientation.Horizontal
        mvarPanel.Panel1MinSize = 150
        mvarPanel.SplitterDistance = 150
        mvarPanel.FixedPanel = FixedPanel.Panel1

        mvarList = New DataGridView
        mvarList.AllowUserToAddRows = False
        mvarList.Dock = DockStyle.Fill
        mvarPanel.Panel2.Controls.Add(mvarList)

        mvarList.AutoGenerateColumns = False

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
        mvarPanel.FixedPanel = FixedPanel.Panel1
        mvarPanel.SplitterDistance = 160
        mvarPanel.IsSplitterFixed = True

        Dim pSp2 As SplitContainer = AddPanel(mvarPanel.Panel1, Orientation.Vertical)
        pSp2.SplitterDistance = 200
        pSp2.FixedPanel = FixedPanel.Panel1
        pSp2.IsSplitterFixed = True

        With AddGroup(pSp2.Panel1, "整理番号")
            AddLabel(.Controls, "議案番号")
            txt議案番号 = AddTextBox(.Controls)
            txt議案番号.Text = ""
            txt議案番号.TextAlign = HorizontalAlignment.Right
            .SetFlowBreak(txt議案番号)

            AddLabel(.Controls, "日程番号")
            txt日程番号 = AddTextBox(.Controls)
            txt日程番号.Text = ""
            txt日程番号.TextAlign = HorizontalAlignment.Right
        End With
        Dim pSp3 As SplitContainer = AddPanel(pSp2.Panel2, Orientation.Horizontal)
        pSp3.SplitterDistance = 50
        pSp3.IsSplitterFixed = True
        pSp3.FixedPanel = FixedPanel.Panel1


        With AddGroup(pSp3.Panel1, "印刷順序")
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
        With AddGroup(pSp3.Panel2, "ページ番号")

        End With


    End Sub

    Private Function AddPanel(pParent As Control, ByVal Ori As Orientation) As SplitContainer
        Dim pPanel As New SplitContainer
        pPanel.Dock = DockStyle.Fill
        pPanel.Orientation = Ori

        pParent.Controls.Add(pPanel)
        Return pPanel
    End Function

    Private Function AddGroup(pPanel As Control, sText As String) As FloatGroup
        Dim pGroupBox As New FloatGroup(pPanel)
        pGroupBox.Text = sText

        Return pGroupBox
    End Function

    Private Function AddLabel(pCtrls As ControlCollection, sText As String) As Label
        Dim pLabel As New Label
        pLabel.AutoSize = True
        pLabel.Margin = New Padding(0, 0, 0, 0)
        pLabel.Text = sText
        pCtrls.Add(pLabel)

        Return pLabel
    End Function
    Private Function AddTextBox(pCtrls As ControlCollection) As TextBox
        Dim pTextBox As New TextBox
        pTextBox.Margin = New Padding(0, 0, 0, 0)

        pCtrls.Add(pTextBox)

        Return pTextBox
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

        Public Sub New(pPanel As Control)
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
        Public Sub SetFlowBreak(pCtrl As Control)
            mvarInnerPanel.SetFlowBreak(pCtrl, True)
        End Sub
    End Class

    Private Sub mvar印刷_CheckedChanged(sender As Object, e As System.EventArgs) Handles mvar印刷.CheckedChanged
        If Me.Parent IsNot Nothing Then
            Me.Parent.Refresh()
        End If
    End Sub

    Private Sub mvar全選択_Click(sender As Object, e As System.EventArgs) Handles mvar全選択.Click
        For Each pRow As DataRowView In mvarList.DataSource
            pRow.Item("選択") = True
        Next
        mvarList.Refresh()
    End Sub
    Private Sub mvar全解除_Click(sender As Object, e As System.EventArgs) Handles mvar全解除.Click
        For Each pRow As DataRowView In mvarList.DataSource
            pRow.Item("選択") = False
        Next
        mvarList.Refresh()
    End Sub

End Class



