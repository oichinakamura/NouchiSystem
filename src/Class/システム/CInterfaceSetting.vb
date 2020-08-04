Public Class CInterfaceSetting
    Inherits CustomControlLIB.SettingParamSK

    Public Sub New(pSys As CSystem)
        MyBase.New()

        If pSys.IsClickOnceDeployed Then
            UserLock = True
        Else
            LoadPath = New IO.DirectoryInfo(My.Application.Info.DirectoryPath).Parent.Parent.FullName & "\Resources\設定XML\Interface.XML"
            SavePath = New IO.DirectoryInfo(My.Application.Info.DirectoryPath).Parent.Parent.FullName & "\Resources\設定XML\Interface.XML"
            UserLock = False
        End If
    End Sub
End Class
