

Public Class CSystemInfo農地台帳
    Inherits HimTools2012.System管理.CSystemInfoSk

    Public Sub New(ByVal sPassword As String, ByVal IsClickOnceDeploy As Boolean)
        MyBase.New()
    End Sub

    Public Overrides ReadOnly Property ApplicationDirectory As String
        Get
            Return MyBase.ApplicationDirectory
        End Get
    End Property


    Public Overrides ReadOnly Property データベース管理 As String
        Get
            Return "Access-DBMan管理"
        End Get
    End Property
End Class
