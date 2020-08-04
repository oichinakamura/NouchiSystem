
Imports System.ComponentModel
Imports HimTools2012.CommonFunc

Public Class classPage農家世帯
    Inherits HimTools2012.controls.CanDragTabControlCollection
    Implements HimTools2012.controls.XMLLayoutContainer

    Protected WithEvents mvarXMLLayout As HimTools2012.controls.XMLLayout

    Public Sub New(Optional ByVal sSQL As String = "")
        MyBase.New(True)

        Me.Name = "農家一覧"
        Me.Text = "農地・農家検索"
        PageInit()

        App農地基本台帳.Set台帳Menu(Me.ToolStrip)
    End Sub

    Private Sub PageInit()

        Try
            Me.Body.SuspendLayout()
            mvarXMLLayout = New HimTools2012.controls.XMLLayout(SysAD.DB(sLRDB), ObjectMan, Me)
            With mvarXMLLayout
                Try
                    .Param.Add("#オリエンテーション", ConvertOrientation(SysAD.GetXMLProperty("農地台帳検索", "スプリッタオリエンテーション", Orientation.Horizontal)))
                Catch ex As Exception

                End Try

                .StartLayout(SysAD.SystemInfo.画面設定, "基本画面")
                FolderTree.Active()

                Try
                    Me.AddTabCtrls(.Controls("TabA"), .Controls("TabB"), .Controls("TabC"), .Controls("TabD"), .Controls("BlockPanelCol"))
                Catch ex As Exception
                    MsgBox(ex.StackTrace, MsgBoxStyle.Critical, "不正な処理")
                End Try
            End With

            If Not SysAD.IsClickOnceDeployed Then
                CType(mvarXMLLayout.Controls("TaskList"), CTaskList).Active()
            End If

            Me.Body.ResumeLayout()
            Me.Body.Refresh()
        Catch ex As Exception
            MsgBox(ex.StackTrace, MsgBoxStyle.Critical, "不正な処理")
        End Try
    End Sub

    Public ReadOnly Property BlockPanelCtrl2() As HimTools2012.controls.BlockPanelControl
        Get
            Return mvarXMLLayout.Controls("BlockPanelCol2")
        End Get
    End Property


    Private Sub mvarXMLLayout_ClickButton(sender As Object, e As System.EventArgs) Handles mvarXMLLayout.ClickButton
        Select Case sender.Name
            Case "早見表" : If Not TabPageContainKey("早見表", True) Then BlockPanelCtrl.BlockPanels.Add(New HimTools2012.Gadget.CPage早見表(True), True)
            Case "総会資料" : mod農地基本台帳.総会資料作成()
            Case "諮問意見書作成" : mod農地基本台帳.諮問意見書作成()
            Case "固定資産検索", "固定資産検索2" : Add各リスト("固定資産検索", "固定資産検索", "固定資産検索リスト", "固定資産検索リスト", GetType(C固定資産農地検索), GetType(C土地台帳リスト))
            Case "転用農地検索", "転用農地検索2" : Add各リスト("転用農地検索", "転用農地検索", "転用農地検索リスト", "転用農地検索リスト", GetType(C転用農地検索), GetType(C転用農地リスト))
            Case "削除農地検索", "削除農地検索2" : Add各リスト("削除農地検索", "削除農地検索", "削除農地検索リスト", "削除農地検索リスト", GetType(C削除農地検索), GetType(C削除農地リスト))
            Case "住記取込検索", "住記取込検索2" : Add各リスト("住記取込検索", "住記取込検索", "住記取込検索リスト", "住記取込検索リスト", GetType(C住記録検索), GetType(C住民記録リスト))
            Case "削除個人検索", "削除個人検索2" : Add各リスト("削除個人検索", "削除個人検索", "削除個人検索リスト", "削除個人検索リスト", GetType(C削除個人検索), GetType(C削除個人リスト))
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(sender.Name)
                    Stop
                End If

        End Select
    End Sub


    Public Overrides Sub GetCommnad(ByRef s As Object, e As System.EventArgs)
        Select Case s.Name
            Case "早見表" : If Not TabPageContainKey("早見表", True) Then BlockPanelCtrl.BlockPanels.Add(New HimTools2012.Gadget.CPage早見表(True), True)
            Case "総会資料" : mod農地基本台帳.総会資料作成()
            Case "諮問意見書作成" : mod農地基本台帳.諮問意見書作成()
            Case "固定資産検索", "固定資産検索2" : Add各リスト("固定資産検索", "固定資産検索", "固定資産検索リスト", "固定資産検索リスト", GetType(C固定資産農地検索), GetType(C土地台帳リスト))
            Case "転用農地検索", "転用農地検索2" : Add各リスト("転用農地検索", "転用農地検索", "転用農地検索リスト", "転用農地検索リスト", GetType(C転用農地検索), GetType(C転用農地リスト))
            Case "削除農地検索", "削除農地検索2" : Add各リスト("削除農地検索", "削除農地検索", "削除農地検索リスト", "削除農地検索リスト", GetType(C削除農地検索), GetType(C削除農地リスト))
            Case "住記取込検索", "住記取込検索2" : Add各リスト("住記取込検索", "住記取込検索", "住記取込検索リスト", "住記取込検索リスト", GetType(C住記録検索), GetType(C住民記録リスト))
            Case "削除個人検索", "削除個人検索2" : Add各リスト("削除個人検索", "削除個人検索", "削除個人検索リスト", "削除個人検索リスト", GetType(C削除個人検索), GetType(C削除個人リスト))
            Case Else
                If Not SysAD.IsClickOnceDeployed Then
                    Debug.Print(s.Name)
                    Stop
                End If
        End Select
    End Sub

    Public ReadOnly Property FolderTree() As CFolderTree
        Get
            Return mvarXMLLayout.Controls("フォルダ")
        End Get
    End Property

    Private Sub Add各リスト(ByVal sName As String, ByVal sText As String, ByVal sListName As String, ByVal sListText As String, ByRef type条件 As System.Type, ByVal pListType As Type)
        Dim pList As HimTools2012.TabPages.NListSK
        Dim pListNew As Boolean = False
        If Not TabPageContainKey(sListName) Then
            pList = Activator.CreateInstance(pListType, {sListName, sListText})
            pListNew = True
            中央Tab.AddPage(pList)
        Else
            pList = GetItem(sListName)
        End If
        If Not TabPageContainKey(sName) Then

            BlockPanelCtrl.BlockPanels.Add(New CPage検索(sName, sText, Activator.CreateInstance(type条件), True, pList), True)
            BlockPanelCtrl.Refresh()
        ElseIf pListNew Then
            With CType(CType(GetItem(sName), HimTools2012.controls.BlockPanel).Body, CPage検索)
                .ToList = pList
            End With
        End If
    End Sub

    Public ReadOnly Property 中央Tab As HimTools2012.controls.TabControlBase
        Get
            Return mvarXMLLayout.Controls("TabB").Body
        End Get
    End Property
    Public ReadOnly Property 詳細Tab As HimTools2012.controls.TabControlBase
        Get
            Return mvarXMLLayout.Controls("TabD").Body
        End Get
    End Property
    Public ReadOnly Property DataViewCollection As HimTools2012.TargetSystem.CDataViewCollection
        Get
            Return mvarXMLLayout.DataViewCollection
        End Get
    End Property

    Public ReadOnly Property 農家リスト As C農家リスト
        Get
            Return mvarXMLLayout.Controls("農家リスト")
        End Get
    End Property
    Public ReadOnly Property 農地リスト As C農地リスト
        Get
            Return mvarXMLLayout.Controls("農地リスト")
        End Get
    End Property
    Public ReadOnly Property 個人リスト As C個人リスト
        Get
            Return mvarXMLLayout.Controls("個人リスト")
        End Get
    End Property
    Public ReadOnly Property 土地履歴リスト As C土地履歴リスト
        Get
            Return mvarXMLLayout.Controls("土地履歴リスト")
        End Get
    End Property
    Public ReadOnly Property BlockPanelCtrl As HimTools2012.controls.BlockPanelControl
        Get
            Return mvarXMLLayout.Controls("BlockPanelCol")
        End Get
    End Property

End Class
