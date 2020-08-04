Imports System.ComponentModel
Imports System.Drawing.Design

''' <summary></summary>
''' <remarks>
''' 未検証 件数6
''' </remarks>
Public Class frmMain

    ''' <summary>Newのコンストラクタ</summary>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/15 14:27]</remarks>
    Public Sub New()
        MyBase.New("frmMain", "農地台帳システム")
        If Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName).Length > 1 Then
            MessageBox.Show(String.Format("{0}は既にプロセスに存在します。多重起動はできません。", "農地台帳システム"))
            'End
        End If

        SysAD = New CSystem(Me, New C農地基本台帳)
        InitializeComponent()
        InitializeMainPage(New CMainPage())

        Do Until SysAD.InitSystem()
            If MsgBox("起動時に初期化を正常に終了できませんでした。再度設定を試みますか?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                End
            End If
        Loop

        SysAD.市町村.InitLocalData()
        SysAD.市町村.InitMenu(Me.MainPage)

        mnuManual = AddMenu(mnuHelp, "mnuManual", New System.Drawing.Size(190, 22), "簡易マニュアル")
        AddHandler Me.EditMenu.Click, AddressOf ClickEditMenu
    End Sub

    ''' <summary></summary>
    ''' <param name="s">class：System.Object</param>
    ''' <param name="e">class：System.Object</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/15 14:27]</remarks>
    Public Sub ClickEditMenu(s, e)
        Dim pCMenu As New ContextMenuStrip
        pCMenu.Items.Add("マスタの編集").Enabled = False
        pCMenu.Show(MousePosition)
    End Sub

    ''' <summary></summary>
    Protected WithEvents mvarT As Timer

    ''' <summary></summary>
    ''' <param name="sender">class：System.Object</param>
    ''' <param name="e">class：System.EventArgs</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/15 14:27]</remarks>
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        mvarT = New Timer

        If SysAD.MapConnection.HasMap Then
            mvarT.Interval = 1000
            mvarT.Start()
        End If
        App農地基本台帳.TBL個人.検索文字初期化()
    End Sub

#Region "追加メニュー"
    ''' <summary></summary>
    Friend WithEvents mnuManual As System.Windows.Forms.ToolStripMenuItem

    ''' <summary></summary>
    ''' <param name="sender">class：System.Object</param>
    ''' <param name="e">class：System.EventArgs</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/15 14:27]</remarks>
    Private Sub mnuManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuManual.Click
        Dim sFile As String = SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\" & "農地台帳マニュアル.pdf"

        If IO.File.Exists(sFile) Then
            OpenPDF(sFile)
        Else
            OpenPDF(My.Application.Info.DirectoryPath & "\" & "manual_light.pdf")
        End If
    End Sub
#End Region

    ''' <summary></summary>
    ''' <param name="sender">class：System.Object</param>
    ''' <param name="e">class：System.EventArgs</param>
    ''' <remarks>Todo 検証してください！！未検証 コメント作成：[2016/09/15 14:27]</remarks>
    Private Sub mvarT_Tick(sender As Object, e As System.EventArgs) Handles mvarT.Tick
        If SysAD.MapConnection.HasMap Then

            Dim St As String = SysAD.MapConnection.GetPropStr("MapToBook", 255)
            If Len(Trim(St)) Then
                If Me.WindowState <> FormWindowState.Maximized Then
                    Me.WindowState = FormWindowState.Maximized
                End If
                Me.Activate()

                Dim sParam() As String = Split(St, ";")
                Select Case sParam(0)
                    Case "OPEN"
                        If SysAD.page農家世帯 Is Nothing OrElse Not Me.MainTabCtrl.ExistPage("農家一覧") Then
                            If MsgBox("農家検索ページがありません、開きますか?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                Open農家検索()
                                CType(ObjectMan.GetObject("農地." & sParam(2)), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
                            Else

                            End If
                        Else
                            CType(ObjectMan.GetObject("農地." & sParam(2)), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
                        End If
                    Case Else
                        Stop
                End Select
            End If
        End If
    End Sub

End Class