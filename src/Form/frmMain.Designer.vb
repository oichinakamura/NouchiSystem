<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits HimTools2012.SystemWindows.CMainFormSK
    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.mnuOpenFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.WindowToolContainer.ContentPanel.SuspendLayout()
        Me.WindowToolContainer.SuspendLayout()
        Me.SuspendLayout()
        '
        'WindowToolContainer
        '
        '
        'WindowToolContainer.ContentPanel
        '
        Me.WindowToolContainer.ContentPanel.Size = New System.Drawing.Size(950, 386)
        Me.WindowToolContainer.Size = New System.Drawing.Size(950, 437)
        Me.WindowToolContainer.TopToolStripPanelVisible = True
        Me.WindowToolContainer.BottomToolStripPanelVisible = True
        '
        'MainTabCtrl
        '
        Me.MainTabCtrl.Size = New System.Drawing.Size(950, 386)
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(950, 437)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "frmMain"
        Me.Text = "農地台帳"
        Me.WindowToolContainer.ContentPanel.ResumeLayout(False)
        Me.WindowToolContainer.ResumeLayout(False)
        Me.WindowToolContainer.PerformLayout()
        Me.ResumeLayout(False)

    End Sub



End Class
