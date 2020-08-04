
Imports System.ComponentModel
Imports HimTools2012.TypeConverterCustom

Module mod農地台帳
    Public sLRDB As String = "LRDB"
    Public s地図情報 As String = "地図情報"
    Public mvarMainForm As frmMain
    Public Sub Main()
        mvarMainForm = New frmMain

    End Sub

    Public Sub 期別農業委員会の追加()
        With New HimTools2012.PropertyGridDialog(New 期別農業委員会管理(1), "期別の農業委員会追加")
            If .ShowDialog = DialogResult.OK Then
                With CType(.ResultProperty, 期別農業委員会管理)
                    App農地基本台帳.DataMaster.AddData("期別農業委員会", .ID, .名称, 0, .ToString)
                    SysAD.page農家世帯.FolderTree.InitTree()
                End With
            End If
        End With
    End Sub

    Public Sub 期別農業委員会の編集(pOBJ As HimTools2012.TargetSystem.CTargetObjectBase)
        With New HimTools2012.PropertyGridDialog(New 期別農業委員会管理(1), "期別の農業委員会追加")
            With CType(.ResultProperty, 期別農業委員会管理)
                Dim sParam As String = App農地基本台帳.DataMaster.Rows.Find({pOBJ.Key.ID, pOBJ.Key.DataClass}).Item("sParam")
                'SysAD.DB(sLRDB).pOBJ.Key.KeyValue()
                Dim sA() As String = Split(sParam, ",")
                .期番号 = Val(sA(0))
                .任期開始 = CDate(sA(1))
                If IsDate(sA(2)) Then
                    .任期終了 = CDate(sA(2))
                End If
            End With

            If .ShowDialog = DialogResult.OK Then
                With CType(.ResultProperty, 期別農業委員会管理)
                    App農地基本台帳.DataMaster.AddData("期別農業委員会", .ID, .名称, 0, .ToString)
                    SysAD.page農家世帯.FolderTree.InitTree()
                End With
            End If
        End With
    End Sub
    Public Sub 期別農業委員会の削除(pOBJ As CObj各種)
        If MsgBox("選択中の農業委員管理情報を削除しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Try
                If App農地基本台帳.DataMaster.DeleteData(pOBJ.Key.DataClass, pOBJ.Key.ID) Then
                    Dim FNode() As TreeNode = SysAD.page農家世帯.FolderTree.Body.Nodes.Find(pOBJ.Key.KeyValue, True)
                    If FNode.Length > 0 Then
                        SysAD.page農家世帯.FolderTree.Body.Nodes.Remove(FNode(0))
                    End If
                End If

            Catch ex As Exception

            End Try
        End If
    End Sub


#Region "農地関連処理"
    Public Sub Sub農地追加(ByRef p個人 As CObj個人, ByRef p農家 As CObj農家)
        Dim pInput As New CInput農地
        Dim pDlg As New dlgInputMulti(pInput, "農地追加", "追加する農地の地番を入力してください")

        If p個人 IsNot Nothing Then
            pInput.所有者ID = p個人.ID
            pInput.所有世帯ID = p個人.世帯ID
        ElseIf p農家 IsNot Nothing Then
            pInput.所有者ID = p農家.世帯主ID
            pInput.所有世帯ID = p農家.ID
        Else
            MsgBox("正しく所有者・所有世帯が指定されていません。")
            Return
        End If

        pDlg.ShowDialog()
        If pDlg.DialogResult = DialogResult.OK Then
            Dim nID As Long = pInput.AddRecord()
            If nID <> 0 Then
                CType(ObjectMan.GetObject("農地." & nID), HimTools2012.TargetSystem.CTargetObjWithView).OpenDataViewNext(SysAD.page農家世帯.DataViewCollection)
            End If
        End If

    End Sub
#End Region

    Public Sub Open農家検索()
        If SysAD.page農家世帯 Is Nothing OrElse Not SysAD.MainForm.MainTabCtrl.TabPages.Contains(SysAD.page農家世帯) Then
            SysAD.page農家世帯 = New classPage農家世帯
            SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
        ElseIf Not SysAD.MainForm.MainTabCtrl.TabPages.Contains(SysAD.page農家世帯) Then
            SysAD.MainForm.MainTabCtrl.TabPages.Add(SysAD.page農家世帯)
        End If
        SysAD.MainForm.MainTabCtrl.SelectedTab = SysAD.page農家世帯
    End Sub
End Module

<TypeConverter(GetType(PropertyOrderConverter))>
Public Class 期別農業委員会管理
    Inherits HimTools2012.InputSupport.CInputSupport

    Public Sub New(Optional n期 As Integer = 0, Optional dt開始 As Object = Nothing, Optional dt終了 As Object = Nothing)
        MyBase.New(農地基本台帳.App農地基本台帳.DataMaster.Body)

        If dt開始 IsNot Nothing AndAlso IsDate(dt開始) Then
            任期開始 = dt開始
        Else
            任期開始 = Nothing
        End If

        If dt終了 IsNot Nothing AndAlso IsDate(dt終了) Then
            任期終了 = dt終了
        Else
            任期終了 = Nothing
        End If

    End Sub

    <Category("農業委員会情報")>
    Public ReadOnly Property 名称 As String
        Get
            Dim sRetStr As New System.Text.StringBuilder(CStr("第" & IIf(期番号 > 0, 期番号.ToString, "x") & "期"))

            sRetStr.Append("[")
            If Not IsNothing(任期開始) AndAlso IsDate(任期開始) AndAlso 任期開始.Year > 1 Then
                sRetStr.Append(任期開始.Year & "/" & 任期開始.Month)
            End If
            sRetStr.Append("～")
            If Not IsNothing(任期終了) AndAlso IsDate(任期終了) AndAlso 任期終了.Year > 1 Then
                sRetStr.Append(任期終了.Year & "/" & 任期終了.Month)
            End If
            sRetStr.Append("]")

            Return sRetStr.ToString
        End Get
    End Property
    <Category("農業委員会情報")>
    Public Property 期番号 As Integer = 0
    <Category("農業委員会任期情報")>
    Public Property 任期開始 As DateTime = Nothing
    <Category("農業委員会任期情報")>
    Public Property 任期終了 As DateTime = Nothing
    Public Overrides Function DataCompleate() As Boolean
        If Not IsNothing(任期開始) AndAlso IsDate(任期開始) AndAlso 任期開始.Year > 1 Then
            Return True
        Else

            Return False
        End If
    End Function
    <Browsable(False)>
    Public ReadOnly Property ID As Integer
        Get
            If Not IsNothing(任期開始) AndAlso IsDate(任期開始) AndAlso 任期開始.Year > 1 Then
                Return 任期開始.Year * 100 + 任期開始.Month
            End If

            Return 0
        End Get
    End Property
    Public Overrides Function ToString() As String
        Dim sRetStr As New System.Text.StringBuilder(期番号.ToString)

        sRetStr.Append(",")
        If Not IsNothing(任期開始) AndAlso IsDate(任期開始) AndAlso 任期開始.Year > 1 Then
            sRetStr.Append(任期開始.Year & "/" & 任期開始.Month & "/" & 任期開始.Day)
        End If

        sRetStr.Append(",")
        If Not IsNothing(任期終了) AndAlso IsDate(任期終了) AndAlso 任期終了.Year > 1 Then
            sRetStr.Append(任期終了.Year & "/" & 任期終了.Month & "/" & 任期終了.Day)
        End If

        Return sRetStr.ToString
    End Function
End Class
