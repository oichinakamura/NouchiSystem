Imports HimTools2012.CommonFunc

Public Class CTabPage現地調査表作成
    Inherits HimTools2012.controls.CTabPageWithToolStrip

    Private WithEvents mvarTree As TreeView
    Private mvar総会番号 As New ToolStripTextBoxWithLabel("総会番号")

    Private mvar受付範囲開始 As New ToolStripDateTimePickerWithlabel("受付範囲")
    Private mvar受付範囲終了 As New ToolStripDateTimePickerWithlabel("～")
    Private mvar調査年月日 As New ToolStripDateTimePickerWithlabel("調査年月日")
    Private WithEvents mvar検索開始 As New ToolStripButton("検索開始")
    Private WithEvents mvar作成開始 As New ToolStripButton("作成開始")
    Private p農地3条現地調査 As 現地調査表
    Private p農地4条現地調査 As 現地調査表
    Private p農地5条現地調査 As 現地調査表
    Private p農地非農地現地調査 As 現地調査表
    Private p農地農用地除外現地調査 As 現地調査表
    Private p農地農用地編入現地調査 As 現地調査表
    Private p農地農用地区分変更現地 As 現地調査表
    Private p農地事業計画変更 As 現地調査表
    Private p農地利用目的変更 As 現地調査表

    Public Sub New()
        MyBase.New(True, True, "現地調査表作成.0", "現地調査表作成")

        mvarTree = New TreeView
        mvarTree.Dock = DockStyle.Fill
        mvarTree.Font = New Font(mvarTree.Font.FontFamily, 10)
        mvarTree.CheckBoxes = True

        With New 議案書作成パラメータ
            mvar総会番号.Text = Now.Month
            mvar受付範囲開始.Value = .開始年月日
            mvar受付範囲終了.Value = .終了年月日
            mvar調査年月日.Value = Now.Date
            Me.ToolStrip.Items.AddRange({mvar総会番号, mvar受付範囲開始, mvar受付範囲終了, mvar調査年月日, mvar検索開始, mvar作成開始})
        End With

        CleateTree()
        Me.ControlPanel.Add(mvarTree)
    End Sub

    Private Sub mvar検索開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar検索開始.Click
        CleateTree()
    End Sub

    Public Sub CleateTree()
        Dim sWhere As String = String.Format("[状態]=0 AND [受付年月日]>=#{1}/{2}/{0}# AND [受付年月日]<=#{4}/{5}/{3}#",
               mvar受付範囲開始.Value.Year, mvar受付範囲開始.Value.Month, mvar受付範囲開始.Value.Day,
               mvar受付範囲終了.Value.Year, mvar受付範囲終了.Value.Month, mvar受付範囲終了.Value.Day)
        Dim pTableA As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM D_申請 WHERE {0}", sWhere)

        App農地基本台帳.TBL申請.MergePlus(pTableA)
        Dim pView As New DataView(App農地基本台帳.TBL申請.Body, sWhere, "[法令],[現地調査番号],[受付番号]", DataViewRowState.CurrentRows)
        Dim pALL As New TreeNode("全体")
        pALL.Name = "全体.0"

        mvarTree.Nodes.Clear()
        mvarTree.Nodes.Add(pALL)
        Dim p現地調査番号Max As New Dictionary(Of enum法令, Integer)
        Dim n議案番号 As Integer = 2

        For Each pRow As DataRowView In pView
            Dim f番号 As enum法令 = 0
            Select Case CType(pRow.Item("法令"), enum法令)
                Case enum法令.農地法3条耕作権 : f番号 = enum法令.農地法3条所有権
                Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : f番号 = enum法令.農地法5条所有権
                Case Else
                    f番号 = CType(pRow.Item("法令"), enum法令)
            End Select

            If Not p現地調査番号Max.ContainsKey(f番号) Then
                p現地調査番号Max.Add(f番号, 0)
            End If
            If Not IsDBNull(pRow.Item("現地調査番号")) Then
                Select Case CType(pRow.Item("法令"), enum法令)
                    Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                        If p現地調査番号Max.Item(enum法令.農地法3条所有権) > pRow.Item("現地調査番号") Then
                            p現地調査番号Max.Item(enum法令.農地法3条所有権) = pRow.Item("現地調査番号")
                        End If
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                        If p現地調査番号Max.Item(enum法令.農地法5条所有権) > pRow.Item("現地調査番号") Then
                            p現地調査番号Max.Item(enum法令.農地法5条所有権) = pRow.Item("現地調査番号")
                        End If
                    Case Else
                        If p現地調査番号Max.Item(pRow.Item("法令")) > pRow.Item("現地調査番号") Then
                            p現地調査番号Max.Item(pRow.Item("法令")) = pRow.Item("現地調査番号")
                        End If
                End Select
            End If
        Next

        For Each pRow As DataRowView In pView
            Dim n番号 As Integer = 0
            Try
                Dim f番号 As enum法令 = 0
                Select Case CType(pRow.Item("法令"), enum法令)
                    Case enum法令.農地法3条耕作権 : f番号 = enum法令.農地法3条所有権
                    Case enum法令.農地法5条貸借, enum法令.農地法5条一時転用 : f番号 = enum法令.農地法5条所有権
                    Case Else
                        f番号 = CType(pRow.Item("法令"), enum法令)
                End Select


                If IsDBNull(pRow.Item("現地調査番号")) Then
                    pRow.Item("現地調査番号") = p現地調査番号Max.Item(f番号) + 1
                    p現地調査番号Max.Item(f番号) = pRow.Item("現地調査番号")
                    n番号 = pRow.Item("現地調査番号")
                Else
                    n番号 = pRow.Item("現地調査番号")
                End If

            Catch ex As Exception
                Stop
            End Try

            Try
                Select Case CType(pRow.Item("法令"), enum法令)
                    Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                        With AddGroupNode(n議案番号, pALL, "法令.3", "農地法3条")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.農地法4条, enum法令.農地法4条一時転用
                        With AddGroupNode(n議案番号, pALL, "法令.4", "農地法4条")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                        With AddGroupNode(n議案番号, pALL, "法令.5", "農地法5条")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.農用地計画変更
                        With AddGroupNode(n議案番号, pALL, "法令.302", "農用地利用計画変更")
                            Dim s区分 As String = ""
                            Select Case Val(pRow.Item("区分").ToString)
                                Case 1 : s区分 = "除外"
                                Case 3 : s区分 = "編入"
                                Case 2 : s区分 = "用途区分"
                            End Select

                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[" & s区分 & "]" & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[" & s区分 & "]" & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.事業計画変更
                        With AddGroupNode(n議案番号, pALL, "法令.303", "事業計画変更")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.農地利用目的変更
                        With AddGroupNode(n議案番号, pALL, "法令.500", "利用目的変更")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                    Case enum法令.非農地証明願
                        With AddGroupNode(n議案番号, pALL, "法令.602", "非農地願い")
                            Dim pChild As TreeNode
                            If IsDBNull(pRow.Item("調査年月日")) Then : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:なし]")
                            Else : pChild = .Nodes.Add(n番号 & ":" & pRow.Item("名称") & "[調査年月日:" & 和暦Format(pRow.Item("調査年月日")) & "]")
                            End If
                            pChild.Name = "申請." & pRow.Item("ID")
                        End With
                End Select
            Catch ex As Exception
                Stop
            End Try
        Next
    End Sub

    Private Function AddGroupNode(ByRef n議案番号 As Integer, ByRef pParent As TreeNode, ByVal sKey As String, ByVal sTitle As String) As TreeNode
        Dim pNodes() As TreeNode = pParent.Nodes.Find(sKey, True)
        If pNodes.Length = 1 Then
            Return pNodes(0)
        Else
            Dim pNode As New TreeNode(String.Format("{0}(議案番号:{1})", sTitle, n議案番号))
            n議案番号 += 1

            pNode.Name = sKey
            pParent.Nodes.Add(pNode)
            Return pNode
        End If
    End Function

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.CloseOK
        End Get
    End Property

#Region "TreeView"
    Private NowCheanging As Boolean = False
    Private Sub mvarTree_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles mvarTree.AfterCheck
        If Not NowCheanging Then
            NowCheanging = True
            ChildSet(e.Node)
            NowCheanging = False
        End If
    End Sub
    Private Sub ChildSet(ByRef ppNode As TreeNode)
        For Each pNode As TreeNode In ppNode.Nodes
            pNode.Checked = ppNode.Checked
            If pNode.Nodes.Count > 0 Then
                ChildSet(pNode)
            End If
        Next
    End Sub

    Private Sub mvarTree_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles mvarTree.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim pItem As TreeViewHitTestInfo = mvarTree.HitTest(e.Location)
            If pItem.Node IsNot Nothing Then
                Select Case GetKeyHead(pItem.Node.Name)
                    Case "申請"
                        Dim p申請 As CObj申請 = ObjectMan.GetObject(pItem.Node.Name)
                        Dim pMenu = p申請.GetContextMenu(Nothing)
                        If pMenu IsNot Nothing Then
                            pMenu.Show(mvarTree.PointToScreen(e.Location))
                        End If
                    Case "法令"
                        Dim pMenu As New ContextMenuStrip
                        mvarItem = pItem.Node
                        AddHandler pMenu.Items.Add("議案番号の変更").Click, AddressOf 番号変更
                        pMenu.Show(mvarTree, e.Location)
                    Case Else

                End Select
            Else
                Dim pMenu As New ContextMenuStrip
                AddHandler pMenu.Items.Add("最新状態にする").Click, AddressOf CleateTree

                pMenu.Show(mvarTree.PointToScreen(e.Location))
            End If
        End If
    End Sub
#End Region

    Private Sub mvar作成開始_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mvar作成開始.Click

        With New 議案書作成パラメータ
            Dim sFolder As String = SysAD.OutputFolder & String.Format("\総会資料{0}_{1}", .n対象年, .n対象月)
            If IO.Directory.Exists(SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "")) Then
                sFolder = SysAD.GetXMLProperty("総会資料関連", "出力先フォルダ", "") & SysAD.OutputFolder & String.Format("\総会資料{0}_{1}", .n対象年, .n対象月)
            End If
        End With

        p農地3条現地調査 = New 現地調査表("\農地法3条.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_農地法第３条許可申請.xml")
        p農地4条現地調査 = New 現地調査表("\農地法4条.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_農地法第４条許可申請.xml")
        p農地5条現地調査 = New 現地調査表("\農地法5条.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_農地法第５条許可申請.xml")
        p農地非農地現地調査 = New 現地調査表("\非農地.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_非農地証明願い.xml")

        p農地農用地除外現地調査 = New 現地調査表("\農用地除外.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_除外.xml")
        p農地農用地編入現地調査 = New 現地調査表("\農用地編入.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_編入.xml")
        p農地農用地区分変更現地 = New 現地調査表("\農用地区分変更.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_用途区分変更.xml")

        p農地事業計画変更 = New 現地調査表("\事業計画変更.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_事業計画変更.xml")
        p農地利用目的変更 = New 現地調査表("\利用目的変更.xml", SysAD.CustomReportFolder(SysAD.市町村.市町村名) & "\現地調査_農地利用目的変更願い.xml")

        If p農地3条現地調査.IsExists * p農地4条現地調査.IsExists * p農地5条現地調査.IsExists * p農地非農地現地調査.IsExists * p農地農用地除外現地調査.IsExists * p農地農用地編入現地調査.IsExists * p農地農用地区分変更現地.IsExists = 0 Then
            MsgBox("現地調査の様式が不足しています。")
        Else
            Dim sDeskTopFolder As String = SysAD.OutputFolder & String.Format("\現地調査{0}_{1}", Now.Year, Now.Month)
            If Not IO.Directory.Exists(sDeskTopFolder) Then
                IO.Directory.CreateDirectory(sDeskTopFolder)
            End If

            SetData(mvarTree.Nodes)
            If p農地3条現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地3条現地調査.Filename, Replace(p農地3条現地調査.sFrame, "{X}", Join(p農地3条現地調査.sSheet.ToArray, "")))
            End If
            If p農地4条現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地4条現地調査.Filename, Replace(p農地4条現地調査.sFrame, "{X}", Join(p農地4条現地調査.sSheet.ToArray, "")))
            End If
            If p農地5条現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地5条現地調査.Filename, Replace(p農地5条現地調査.sFrame, "{X}", Join(p農地5条現地調査.sSheet.ToArray, "")))
            End If
            If p農地非農地現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地非農地現地調査.Filename, Replace(p農地非農地現地調査.sFrame, "{X}", Join(p農地非農地現地調査.sSheet.ToArray, "")))
            End If
            If p農地農用地除外現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地農用地除外現地調査.Filename, Replace(p農地農用地除外現地調査.sFrame, "{X}", Join(p農地農用地除外現地調査.sSheet.ToArray, "")))
            End If
            If p農地農用地編入現地調査.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地農用地編入現地調査.Filename, Replace(p農地農用地編入現地調査.sFrame, "{X}", Join(p農地農用地編入現地調査.sSheet.ToArray, "")))
            End If
            If p農地農用地区分変更現地.sSheet.Count > 0 Then
                HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地農用地区分変更現地.Filename, Replace(p農地農用地区分変更現地.sFrame, "{X}", Join(p農地農用地区分変更現地.sSheet.ToArray, "")))
            End If

            If p農地事業計画変更.IsExists * p農地利用目的変更.IsExists = 0 Then
            Else
                If p農地事業計画変更.sSheet.Count > 0 Then
                    HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地事業計画変更.Filename, Replace(p農地事業計画変更.sFrame, "{X}", Join(p農地事業計画変更.sSheet.ToArray, "")))
                End If
                If p農地利用目的変更.sSheet.Count > 0 Then
                    HimTools2012.TextAdapter.SaveTextFile(sDeskTopFolder & p農地利用目的変更.Filename, Replace(p農地利用目的変更.sFrame, "{X}", Join(p農地利用目的変更.sSheet.ToArray, "")))
                End If
            End If
            System.Diagnostics.Process.Start(sDeskTopFolder)
        End If
    End Sub

    Private n議案番号 As Integer = 0
    Private Sub SetData(ByVal pNodes As TreeNodeCollection)
        For Each pNode As TreeNode In pNodes
            If GetKeyHead(pNode.Name) = "法令" Then
                n議案番号 = Val(Mid(pNode.Text, InStr(pNode.Text, "議案番号:") + 5))
            End If

            If pNode.Nodes.Count > 0 Then
                SetData(pNode.Nodes)
            End If
            If pNode.Checked AndAlso GetKeyHead(pNode.Name) = "申請" Then
                Dim pRow As DataRow = App農地基本台帳.TBL申請.FindRowByID(GetKeyCode(pNode.Name))
                If pRow IsNot Nothing Then
                    Select Case CType(pRow.Item("法令"), enum法令)
                        Case enum法令.農地法3条所有権, enum法令.農地法3条耕作権
                            p農地3条現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                        Case enum法令.農地法4条, enum法令.農地法4条一時転用
                            p農地4条現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                        Case enum法令.農地法5条所有権, enum法令.農地法5条貸借, enum法令.農地法5条一時転用
                            p農地5条現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                        Case enum法令.非農地証明願
                            p農地非農地現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                        Case enum法令.農用地計画変更
                            Select Case Val(pRow.Item("区分").ToString)
                                Case 1
                                    p農地農用地除外現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                                Case 3
                                    p農地農用地編入現地調査.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                                Case 2
                                    p農地農用地区分変更現地.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                            End Select
                        Case enum法令.事業計画変更
                            p農地事業計画変更.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                        Case enum法令.農地利用目的変更
                            p農地利用目的変更.AddPage(Val(mvar総会番号.Text), n議案番号, pRow, mvar調査年月日.Value)
                    End Select
                End If
            End If
        Next
    End Sub
    Private Class 現地調査表
        Public ReadOnly Property IsExists As Boolean
            Get
                Return sBook.Length > 0
            End Get
        End Property

        Public sBook As String = ""
        Public sDefault As String = ""
        Public sFrame As String = ""
        Public sSheet As New List(Of String)
        Public Filename As String


        Public Sub New(ByVal sFileName As String, ByVal sPath As String)
            Filename = sFileName
            If IO.File.Exists(sPath) Then
                sBook = HimTools2012.TextAdapter.LoadTextFile(sPath)
                sDefault = Mid(sBook, InStr(sBook, " <Worksheet ss:"))
                sDefault = Strings.Left(sDefault, InStr(sDefault, " </Worksheet>") + Len(" </Worksheet>") + 1)
                sFrame = Replace(sBook, sDefault, "{X}")
            End If
        End Sub

        Public Sub AddPage(ByVal n総会番号 As Integer, ByVal n議案番号 As Integer, ByRef pRow As DataRow, ByVal dt調査年月日 As Date)
            Dim sPage As String = sDefault

            Select Case Val(pRow.Item("法令").ToString)
                Case enum法令.農地法5条所有権 : sPage = Replace(sPage, "{形態}", "所有権")
                Case enum法令.農地法5条貸借 : sPage = Replace(sPage, "{形態}", "貸借")
                Case Else : sPage = Replace(sPage, "{法令}", "")
            End Select

            sPage = Replace(sPage, "{総会番号}", n総会番号)
            sPage = Replace(sPage, "{議案番号}", n議案番号)
            sPage = Replace(sPage, "{受付番号}", pRow.Item("受付番号"))
            sPage = Replace(sPage, "{現地調査番号}", pRow.Item("現地調査番号").ToString)
            sPage = Replace(sPage, "{調査年月日}", 和暦Format(dt調査年月日))

            Check調査員(pRow, sPage)

            sPage = Replace(sPage, "{譲受人}", pRow.Item("氏名B").ToString)
            sPage = Replace(sPage, "{譲受人住所}", pRow.Item("住所B").ToString)
            sPage = Replace(sPage, "{譲受年齢}", pRow.Item("年齢B").ToString)
            sPage = Replace(sPage, "{譲渡人}", pRow.Item("氏名A").ToString)
            sPage = Replace(sPage, "{譲渡人住所}", pRow.Item("住所A").ToString)

            sPage = Replace(sPage, "{当初計画者}", pRow.Item("氏名C").ToString) '事業計画変更
            sPage = Replace(sPage, "{当初計画者住所}", pRow.Item("住所C").ToString)
            sPage = Replace(sPage, "{事業計画者}", pRow.Item("氏名A").ToString)
            sPage = Replace(sPage, "{事業計画者住所}", pRow.Item("住所A").ToString)

            Dim sList As String = pRow.Item("農地リスト").ToString
            sList = Replace(sList, "転用農地.", "")
            sList = Replace(Replace(sList, "農地.", ""), ";", ",")

            Dim s申請地 As String = ""
            Dim 田面積 As Decimal = 0
            Dim 畑面積 As Decimal = 0

            sPage = Replace(sPage, "{転用目的}", pRow.Item("申請理由A").ToString)
            sPage = Replace(sPage, "{転用事由}", pRow.Item("申請理由B").ToString)

            sPage = Replace(sPage, "{変更目的}", pRow.Item("申請理由A").ToString) '事業計画変更
            sPage = Replace(sPage, "{変更前目的}", pRow.Item("予備2").ToString)

            sPage = Replace(sPage, "{利用目的}", pRow.Item("用途").ToString) '利用目的変更
            sPage = Replace(sPage, "{変更事由}", pRow.Item("申請理由A").ToString)

            Dim nCount As Integer = -1
            If sList.Length > 0 Then
                Dim pTBL As DataTable = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D:農地Info] WHERE [ID] In (" & sList & ")")
                App農地基本台帳.TBL農地.MergePlus(pTBL)
                pTBL = SysAD.DB(sLRDB).GetTableBySqlSelect("SELECT * FROM [D_転用農地] WHERE [ID] In (" & sList & ")")
                App農地基本台帳.TBL農地.MergePlus(pTBL)

                For Each pRV As DataRowView In New DataView(App農地基本台帳.TBL農地.Body, "[ID] In (" & sList & ")", "", DataViewRowState.CurrentRows)
                    nCount += 1
                    If s申請地 = "" Then
                        s申請地 = pRV.Item("土地所在").ToString
                    End If
                    田面積 += Val(pRV.Item("田面積").ToString)
                    畑面積 += Val(pRV.Item("畑面積").ToString)
                Next
            End If

            sPage = Replace(sPage, "{申請地}", s申請地 & IIF(nCount > 0, " 外 " & nCount & " 筆", ""))
            sPage = Replace(sPage, "{田面積}", 田面積.ToString("#,###"))
            sPage = Replace(sPage, "{畑面積}", 畑面積.ToString("#,###"))
            sPage = Replace(sPage, "{計面積}", (田面積 + 畑面積).ToString("#,###"))
            sPage = Replace(sPage, "{経営面積}", Val(pRow.Item("経営面積B").ToString).ToString("#,###"))

            Select Case Val(pRow.Item("所有権移転の種類").ToString)
                Case 1 : sPage = Replace(sPage, "{種類}", "売買")
                Case 2 : sPage = Replace(sPage, "{種類}", "贈与")
                Case 3 : sPage = Replace(sPage, "{種類}", "交換")
                Case Else
                    sPage = Replace(sPage, "{種類}", "")
            End Select

            Select Case Val(pRow.Item("農地区分").ToString)
                Case 1
                    sPage = Replace(sPage, "{農地区分}", "第１種農地")
                    sPage = Replace(sPage, "{立地基準農地区分1}", "農地区分は" & "第１種農地であるが、第１種農地の不許可の例外である")
                    sPage = Replace(sPage, "{立地基準農地区分2}", "に該当する為、問題なし。")
                Case 2
                    sPage = Replace(sPage, "{農地区分}", "第２種農地")
                    sPage = Replace(sPage, "{立地基準農地区分1}", "農地区分は" & "第２種農地であるが、")
                    sPage = Replace(sPage, "{立地基準農地区分2}", "に該当する為、問題なし。")
                Case 3
                    sPage = Replace(sPage, "{農地区分}", "第３種農地")
                    sPage = Replace(sPage, "{立地基準農地区分1}", "農地区分は" & "第３種農地で問題なし。")
                    sPage = Replace(sPage, "{立地基準農地区分2}", "")
                Case 5
                    sPage = Replace(sPage, "{農地区分}", "農用地区域内農地")
                    sPage = Replace(sPage, "{立地基準農地区分1}", "農地区分は" & "農用地区域内農地であるが、除外後は第　 種農地となり")
                    sPage = Replace(sPage, "{立地基準農地区分2}", "に該当する為、問題なし。")
                Case Else
                    sPage = Replace(sPage, "{農地区分}", "")
                    sPage = Replace(sPage, "{立地基準農地区分1}", "")
                    sPage = Replace(sPage, "{立地基準農地区分2}", "")
            End Select

            sPage = Replace(sPage, "SH01", pRow.Item("受付番号").ToString & "_" & pRow.Item("名称").ToString)
            sSheet.Add(sPage)
        End Sub

        Public Sub Check調査員(ByRef pRow As DataRow, ByRef sPage As String)
            If Not IsDBNull(pRow.Item("農業委員1")) Then
                Dim pRowX As DataRow = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員1"), "農業委員"})
                If pRowX IsNot Nothing Then
                    sPage = Replace(sPage, "{調査委員}", pRowX.Item("名称").ToString)
                    sPage = Replace(sPage, "{調査委員1}", pRowX.Item("名称").ToString)

                    If Not IsDBNull(pRow.Item("農業委員2")) Then
                        pRowX = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員2"), "農業委員"})
                        If pRowX IsNot Nothing Then
                            sPage = Replace(sPage, "{調査委員2}", pRowX.Item("名称").ToString)

                            If Not IsDBNull(pRow.Item("農業委員3")) Then
                                pRowX = App農地基本台帳.DataMaster.Rows.Find({pRow.Item("農業委員3"), "農業委員"})
                                If pRowX IsNot Nothing Then
                                    sPage = Replace(sPage, "{調査委員3}", pRowX.Item("名称").ToString)
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If Not IsDBNull(pRow.Item("調査員A")) Then
                    sPage = Replace(sPage, "{調査委員}", pRow.Item("調査員A"))
                    sPage = Replace(sPage, "{調査委員1}", pRow.Item("調査員A"))
                End If
            End If

            sPage = Replace(sPage, "{調査委員}", "")
            sPage = Replace(sPage, "{調査委員1}", "")
            sPage = Replace(sPage, "{調査委員2}", "")
            sPage = Replace(sPage, "{調査委員3}", "")
        End Sub
    End Class

    Private mvarItem As TreeNode = Nothing

    Private Sub 番号変更()
        If mvarItem IsNot Nothing Then
            n議案番号 = Val(Mid(mvarItem.Text, InStr(mvarItem.Text, "議案番号:") + 5))
            Dim s As String = InputBox("議案番号を入力してください", "議案番号の変更", n議案番号)
            If Val(s) > 0 Then
                Dim St As String = HimTools2012.StringF.Left(mvarItem.Text, InStr(mvarItem.Text, "議案番号:") + 4) & Val(s) & ")"
                mvarItem.Text = St
            End If
        End If

        mvarItem = Nothing
    End Sub

    'Private Sub mvarTree_NodeMouseClick(sender As Object, e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles mvarTree.NodeMouseClick
    '    Dim pItem As Windows.Forms.TreeViewHitTestInfo = mvarTree.HitTest(e.Location)
    '    If pItem.Node IsNot Nothing Then
    '        Select Case e.Button
    '            Case Windows.Forms.MouseButtons.Right

    '        End Select
    '    End If
    'End Sub
End Class
