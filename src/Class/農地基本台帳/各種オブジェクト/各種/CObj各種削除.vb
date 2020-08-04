Imports HimTools2012.CommonFunc
'

Public Class CObj各種削除 : Inherits CObj各種
    Public Sub New(ByVal sKey As String)
        MyBase.New(sKey)
        Select Case GetKeyHead(sKey)
            Case "削除農地" : Me.SetRow(App農地基本台帳.TBL削除農地.FindRowByID(GetKeyCode(sKey)))
            Case "削除個人" : Me.SetRow(App農地基本台帳.TBL削除個人.FindRowByID(GetKeyCode(sKey)))
        End Select
    End Sub

    Public Overrides Function GetContextMenu(Optional ByVal pMenu As HimTools2012.controls.MenuItemEx = Nothing, Optional nDips As Integer = 1, Optional sParam() As String = Nothing) As HimTools2012.controls.MenuPlus
        Dim pMenuPlus As HimTools2012.controls.MenuPlus = CreateMenu(pMenu)

        With pMenuPlus
            Select Case Me.Key.DataClass
                Case "削除農地"
                    If pMenu Is Nothing Then
                        .AddMenu("元に戻す", , AddressOf ClickMenu)
                    End If
                    .AddMenu("履歴一覧", , AddressOf ClickMenu)
                Case "削除個人" : .AddMenu("元に戻す", , AddressOf ClickMenu)
                Case Else
                    CasePrint(Me.Key.DataClass)
            End Select
        End With

        Return pMenuPlus
    End Function

    Public Overrides Function DoCommand(ByVal sCommand As String, ByVal ParamArray sParams() As String) As Object
        Select Case Me.Key.DataClass & "-" & sCommand
            Case "削除農地-開く"
            Case "削除農地-履歴一覧"
                SysAD.page農家世帯.土地履歴リスト.検索開始("[LID]=" & Me.ID, "[LID]=" & Me.ID)
            Case "削除個人-閉じる"
            Case Else
                CasePrint(Me.Key.DataClass & "-" & sCommand)
        End Select
        Return ""
    End Function

    Public Overrides Sub ClickMenu(s As Object, e As System.EventArgs)
        Select Case Me.Key.DataClass & "-" & CType(s, ToolStripMenuItem).Text
            Case "削除農地-元に戻す" : 農地復元(Me.ID, C農地削除.enum転送先.削除農地, "")
            Case "削除個人-元に戻す" : 個人復元(Me)
            Case "削除農地-履歴一覧"
                SysAD.page農家世帯.土地履歴リスト.検索開始("[LID]=" & Me.ID, "[LID]=" & Me.ID)
            Case Else
                CasePrint(Me.Key.DataClass & "-" & CType(s, ToolStripMenuItem).Text)

                MyBase.ClickMenu(s, e)
        End Select
    End Sub
    Public Overrides ReadOnly Property DataTableWithUpdateList As HimTools2012.Data.DataTableWithUpdateList
        Get
            Return Nothing
        End Get
    End Property

End Class
