
Public Class dlgLoginForm

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        If TypeOf UsernameTextBox.SelectedItem Is SelectItem Then
            With CType(UsernameTextBox.SelectedItem, SelectItem)
                If PasswordTextBox.Text = .Password Then
                    SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(.ID, .Name, .Password, .Flag)
                    SaveSub()

                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                End If
            End With
        Else
            Select Case UsernameTextBox.Text
                Case "0555"
                    If PasswordTextBox.Text = "5555" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(0, "0555", "5555", 0)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "閲覧"
                    If PasswordTextBox.Text = "hiokic" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(0, "閲覧者", "hiokic", 0)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "本庁"
                    If PasswordTextBox.Text = "216" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(1, "本庁", "216", 1)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "0333"
                    If PasswordTextBox.Text = "3333" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(1, "0333", "3333", 1)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "管理"
                    If PasswordTextBox.Text = "him" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(2, "本庁", "216", 1)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "農業委員会"
                    If PasswordTextBox.Text = "1111" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(2, "農業委員会", "1111", 1)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "閲覧のみ"
                    If PasswordTextBox.Text = "0000" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(0, "閲覧のみ", "0000", 0)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If
                Case "管理者"
                    If PasswordTextBox.Text = "1500" Then
                        SysAD.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(1, "管理者", "1500", 1)
                        SaveSub()

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    End If


                Case Else

            End Select

        End If
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private mvarTable As DataTable = Nothing
    Public Sub New(Optional pUsers As DataTable = Nothing)
        InitializeComponent()
        If SysAD.市町村.市町村名 = "日置市" Then
            UsernameTextBox.Items.Add("閲覧")
            UsernameTextBox.Items.Add("本庁")
            UsernameTextBox.Items.Add("管理")
        ElseIf SysAD.市町村.市町村名 = "宗像市" Then
            UsernameTextBox.Items.Add("0555")
            UsernameTextBox.Items.Add("0333")
        Else
            If pUsers IsNot Nothing Then
                mvarTable = pUsers
                For Each pRow As DataRow In pUsers.Rows
                    UsernameTextBox.Items.Add(New SelectItem(pRow.Item("ID"), pRow.Item("登録者名"), pRow.Item("PW"), pRow.Item("権利フラグ")))
                Next
            Else
                UsernameTextBox.Items.Add("閲覧のみ")
                UsernameTextBox.Items.Add("農業委員会")
                UsernameTextBox.Items.Add("管理者")
            End If
        End If
        UsernameTextBox.Text = GetSetting("Avail台帳管理", "農地・農家台帳", "UserSelect", "")
    End Sub

    Private Sub UsernameTextBox_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles UsernameTextBox.SelectionChangeCommitted
        SaveSub()
    End Sub

    Private Sub SaveSub()
        SaveSetting("Avail台帳管理", "農地・農家台帳", "UserSelect", UsernameTextBox.Text)
    End Sub

    Public Class SelectItem
        Public ID As Long = 0
        Public Name As String = ""
        Public Password As String = ""
        Public Flag As Integer = 0
        Public Sub New(ByVal nID As Long, ByVal sName As String, ByVal sPass As String, ByVal pFlag As Integer)
            ID = nID
            Name = sName
            Password = sPass
            Flag = pFlag
        End Sub
        Public Overrides Function ToString() As String
            Return Me.Name
        End Function

    End Class
End Class
