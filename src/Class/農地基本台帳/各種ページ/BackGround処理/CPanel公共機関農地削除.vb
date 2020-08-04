
Imports System.ComponentModel

Public Class CPanel農地削除
    Inherits HimTools2012.TabPages.BackGroundPage

    Public 対象農地 As DataTable
    Public 転送先 As enum転送先 = enum転送先.削除農地
    Public n履歴 As Integer = 0

    Public 異動日 As DateTime = Nothing
    Public 異動事由内容 As String

    Public Enum enum転送先
        削除農地 = 1
        転用農地 = 2
    End Enum

    Public Sub New(ByVal sText As String, ByVal bCancelable As Boolean, ByVal Mess As String, ByVal s異動事由内容 As String, Optional ByVal dt異動日 As DateTime = Nothing)
        MyBase.New(False, True, Guid.NewGuid.ToString, sText, bCancelable, Mess)

        If Not IsNothing(dt異動日) Then
            異動日 = dt異動日
        Else
            異動日 = Now.Date
        End If
        異動事由内容 = s異動事由内容
    End Sub

    Public Sub Execute()
        MaxValue = 対象農地.Rows.Count

        Me.Start(AddressOf bgW_DoWork)
    End Sub

    Private Sub bgW_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

        Dim pBGW As BackgroundWorker = sender

        Select Case 転送先
            Case enum転送先.削除農地
        End Select

        For Each pRow As DataRow In 対象農地.Rows
            Dim sRet As String = ""
            Select Case 転送先
                Case enum転送先.削除農地
                    Debug.Print(App農地基本台帳.TBL削除農地.Rows.Count)
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_削除農地] WHERE [ID]=" & pRow.Item("ID"))
                    sRet = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_削除農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
                Case enum転送先.転用農地
                    SysAD.DB(sLRDB).ExecuteSQL("DELETE FROM [D_転用農地] WHERE [ID]=" & pRow.Item("ID"))
                    sRet = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO D_転用農地 SELECT [D:農地Info].* FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
            End Select

            If sRet = "" Or sRet = "OK" Then
                Dim St As String = SysAD.DB(sLRDB).ExecuteSQL("DELETE [D:農地Info].ID FROM [D:農地Info] WHERE ((([D:農地Info].ID)=" & pRow.Item("ID") & "));")
                If St.Length = 0 Or St = "OK" Then
                    Dim ppRow As DataRow = App農地基本台帳.TBL農地.Rows.Find(pRow.Item("ID"))

                    If ppRow IsNot Nothing Then
                        Try
                            App農地基本台帳.TBL農地.Rows.Remove(ppRow)
                        Catch ex As Exception

                        End Try
                    End If

                    Dim s異動日 As String = String.Format("#{0}/{1}/{2}#", 異動日.Month, 異動日.Day, 異動日.Year)
                    Select Case n履歴
                        Case 261
                            If 異動事由内容 = "" Then
                                異動事由内容 = "地図システムより非農地確定"
                            End If
                        Case 844
                            If 異動事由内容 = "" Then
                                異動事由内容 = "換地処理"
                            End If
                        Case Else
                            If 異動事由内容 = "" Then
                                異動事由内容 = "システムより削除"
                            End If
                    End Select
                    Dim sRetX As String = SysAD.DB(sLRDB).ExecuteSQL("INSERT INTO [D_土地履歴]([LID],[異動事由],[内容],[異動日],[更新日],[入力日]) VALUES({0},261,'" & 異動事由内容 & "'," & s異動日 & "," & s異動日 & "," & s異動日 & ")", pRow.Item("ID"))
                End If
            End If
            If _Cancel Then
                Exit For
            End If
            Me.IncrementProgress(1)
        Next
    End Sub

End Class



