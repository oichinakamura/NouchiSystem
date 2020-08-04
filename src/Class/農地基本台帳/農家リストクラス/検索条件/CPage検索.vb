
Imports HimTools2012.controls
Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms.Design


Public Class CPanel検索
    Inherits ToolStripContainer

    Public WithEvents 検索Grid As CPropertyGridPlus
    Private mvarToolStrip As ToolStrip
    Protected WithEvents mvarTBtn As New ToolStripButton
    Protected WithEvents mvarFontP As New ToolStripButton
    Protected WithEvents mvarFontM As New ToolStripButton

    Public Event 検索(ByVal sDB検索文字列 As String, ByVal sView検索文字列 As String)
    Public mvar検索条件 As Common検索条件
    Public mvarContextMen As ContextMenu
    Public ToList As CNList農地台帳
    Public sListKey As String = ""
    Private mvarClassname As String

    Public Sub New(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件, ByVal bCloseable As Boolean, Optional pList As HimTools2012.TabPages.NListSK = Nothing)
        'MyBase.New(bCloseable, True, sName, sText)
        New共通(sText, sName, p検索条件)
        ToList = pList
    End Sub
    Public Sub New(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件, ByVal sListName As String, ByVal sClassName As String, Optional bCloseable As Boolean = False)
        'MyBase.New(bCloseable, True, sName, sText)
        New共通(sText, sName, p検索条件)
        sListKey = sListName
        mvarClassname = sClassName
    End Sub

    Private Sub New共通(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件)
        Me.検索Grid = New CPropertyGridPlus
        Me.検索Grid.Dock = DockStyle.Fill
        検索Grid.LineColor = Color.LightBlue
        検索Grid.ToolbarVisible = False

        If p検索条件 IsNot Nothing Then
            mvar検索条件 = p検索条件
            Me.検索Grid.SelectedObject = mvar検索条件
        End If

        Me.ContentPanel.Controls.Add(Me.検索Grid)

        mvarToolStrip = New ToolStrip
        mvarToolStrip.Stretch = True
        Me.TopLevelControl.Controls.Add(mvarToolStrip)

        mvarTBtn.Text = "検索開始<F3>"
        mvarTBtn.AutoSize = True
        mvarTBtn.BackColor = Color.Navy
        mvarTBtn.ForeColor = Color.White
        mvarTBtn.Alignment = ToolStripItemAlignment.Right
        mvarToolStrip.Items.Add(mvarTBtn)

        mvarFontP.Image = SysAD.ImageList16.Images("searchPlus")
        mvarFontP.ImageTransparentColor = Color.Magenta
        mvarFontP.ToolTipText = "文字を拡大"

        mvarFontM.Image = SysAD.ImageList16.Images("searchMinus")
        mvarFontM.ImageTransparentColor = Color.Magenta
        mvarFontM.ToolTipText = "文字を縮小"
        mvarToolStrip.Items.AddRange(New ToolStripItem() {mvarFontP, mvarFontM})

        mvarContextMen = New ContextMenu
        With mvarContextMen.MenuItems.Add("検索開始")
            AddHandler .Click, AddressOf mvarBtn_Click
            .Shortcut = Shortcut.F3
        End With

        検索Grid.ContextMenu = mvarContextMen
    End Sub

    Private Sub mvarBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarTBtn.Click
        Do検索()
    End Sub

    Private Sub mvarFontP_Click(sender As Object, e As System.EventArgs) Handles mvarFontP.Click
        Me.Font = New System.Drawing.Font(Me.Font.FontFamily, Me.Font.Size + 1)
    End Sub
    Private Sub mvarFontM_Click(sender As Object, e As System.EventArgs) Handles mvarFontM.Click
        Try
            Me.Font = New System.Drawing.Font(Me.Font.FontFamily, Me.Font.Size - 1)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub 検索Grid_KeyDownF3() Handles 検索Grid.KeyDownF3, 検索Grid.KeyDownEnter
        Do検索()
    End Sub

    Public Sub Do検索()
        If ToList IsNot Nothing Then
            ToList.検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
        ElseIf sListKey.Length > 0 Then
            If Not SysAD.page農家世帯.TabPageContainKey(sListKey) Then
                Dim T As System.Type = Type.GetType(mvarClassname)
                Dim pList As CNList農地台帳 = Activator.CreateInstance(T)
                SysAD.page農家世帯.中央Tab.AddPage(pList)
                pList.検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
            Else
                CType(SysAD.page農家世帯.GetItem(sListKey), CNList農地台帳).検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
            End If
        Else
            RaiseEvent 検索(mvar検索条件.ToString, mvar検索条件.View検索条件)
        End If

    End Sub

    Private Sub CPage検索_ParentChanged(sender As Object, e As System.EventArgs) Handles Me.ParentChanged
        'Me.ImageKey = "SearchItem"
    End Sub


End Class

Public Class CPage検索
    Inherits HimTools2012.TabPages.CPage検索SK
    Protected TS As New ToolStripContainer

    Public WithEvents 検索Grid As CPropertyGridPlus
    Protected WithEvents mvarTBtn As New ToolStripButton
    Protected WithEvents mvarFontP As New ToolStripButton
    Protected WithEvents mvarFontM As New ToolStripButton

    Public Event 検索(ByVal sDB検索文字列 As String, ByVal sView検索文字列 As String)
    Public mvar検索条件 As Common検索条件
    Public mvarContextMen As ContextMenu
    Public ToList As CNList農地台帳
    Public sListKey As String = ""
    Private mvarClassname As String

    Public Sub New(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件, ByVal bCloseable As Boolean, Optional pList As HimTools2012.TabPages.NListSK = Nothing)
        MyBase.New(bCloseable, sName, sText)
        New共通(sText, sName, p検索条件)
        ToList = pList
    End Sub
    Public Sub New(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件, ByVal sListName As String, ByVal sClassName As String, Optional bCloseable As Boolean = False)
        MyBase.New(bCloseable, sName, sText)
        New共通(sText, sName, p検索条件)
        sListKey = sListName
        mvarClassname = sClassName
    End Sub

    Private Sub New共通(ByVal sText As String, ByVal sName As String, ByRef p検索条件 As Common検索条件)
        Me.検索Grid = New CPropertyGridPlus
        Me.検索Grid.Dock = DockStyle.Fill
        検索Grid.LineColor = Color.LightBlue
        検索Grid.ToolbarVisible = False

        If p検索条件 IsNot Nothing Then
            mvar検索条件 = p検索条件
            Me.検索Grid.SelectedObject = mvar検索条件
        End If

        Me.ContentPanel.Controls.Add(Me.検索Grid)

        mvarTBtn.Text = "検索開始<F3>"
        mvarTBtn.AutoSize = True
        mvarTBtn.BackColor = Color.Navy
        mvarTBtn.ForeColor = Color.White
        mvarTBtn.Alignment = ToolStripItemAlignment.Right
        Me.ToolStrip.Items.Add(mvarTBtn)

        mvarFontP.Image = SysAD.ImageList16.Images("searchPlus")
        mvarFontP.ImageTransparentColor = Color.Magenta
        mvarFontP.ToolTipText = "文字を拡大"

        mvarFontM.Image = SysAD.ImageList16.Images("searchMinus")
        mvarFontM.ImageTransparentColor = Color.Magenta
        mvarFontM.ToolTipText = "文字を縮小"
        Me.ToolStrip.Items.AddRange(New ToolStripItem() {mvarFontP, mvarFontM})

        mvarContextMen = New ContextMenu
        With mvarContextMen.MenuItems.Add("検索開始")
            AddHandler .Click, AddressOf mvarBtn_Click
            .Shortcut = Shortcut.F3
        End With

        検索Grid.ContextMenu = mvarContextMen
    End Sub

    Private Sub mvarBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvarTBtn.Click
        Do検索()
    End Sub

    Private Sub mvarFontP_Click(sender As Object, e As System.EventArgs) Handles mvarFontP.Click
        Me.Font = New System.Drawing.Font(Me.Font.FontFamily, Me.Font.Size + 1)
    End Sub
    Private Sub mvarFontM_Click(sender As Object, e As System.EventArgs) Handles mvarFontM.Click
        Try
            Me.Font = New System.Drawing.Font(Me.Font.FontFamily, Me.Font.Size - 1)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub 検索Grid_KeyDownF3() Handles 検索Grid.KeyDownF3, 検索Grid.KeyDownEnter
        Do検索()
    End Sub

    Public Sub Do検索()
        If ToList IsNot Nothing Then
            ToList.検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
        ElseIf sListKey.Length > 0 Then
            If Not SysAD.page農家世帯.TabPageContainKey(sListKey) Then
                Dim T As System.Type = Type.GetType(mvarClassname)
                Dim pList As CNList農地台帳 = Activator.CreateInstance(T)
                SysAD.page農家世帯.中央Tab.AddPage(pList)
                pList.検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
            Else
                CType(SysAD.page農家世帯.GetItem(sListKey), CNList農地台帳).検索開始(mvar検索条件.ToString, mvar検索条件.View検索条件)
            End If
        Else
            RaiseEvent 検索(mvar検索条件.ToString, mvar検索条件.View検索条件)
        End If

    End Sub

    Public Overrides ReadOnly Property IconImageBass As System.Drawing.Image
        Get
            Return My.Resources.Resource1.SearchItem.ToBitmap
        End Get
    End Property

    Public Overrides Property IconKey As String
        Get
            Return "SearchItem"
        End Get
        Set(value As String)

        End Set
    End Property
End Class
