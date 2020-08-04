

Public Class CManualPage
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private mvarB As System.Windows.Forms.WebBrowser

    Public Sub New()
        MyBase.New(True, True, "マニュアル", "マニュアル")
        mvarB = New System.Windows.Forms.WebBrowser
        mvarB.Dock = DockStyle.Fill
        Me.ControlPanel.Add(mvarB)


        mvarB.Url = New Uri(My.Application.Info.DirectoryPath & "\" & "manual_light.pdf" & "#toolbar=0&navpanes=1")
    End Sub
    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

End Class
