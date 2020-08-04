Imports System.Windows.Forms

Public Class dlg住記選択
    Public Property nID As Integer = 0

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If ListView1.SelectedItems.Count = 1 Then
            nID = Val(ListView1.SelectedItems(0).SubItems(3).Text)

            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Else


        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub 検索開始(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim sWhere As String = ""

        If txtKana.Text.Length > 0 Then
            sWhere = sWhere & IIf(sWhere.Length > 0, " AND ", "") & "[フリガナ] Like '" & Replace(txtKana.Text.Replace("*", "%") & "%'", "%%", "%")
        End If

        If TextBox1.Text.Length > 0 Then
            sWhere = sWhere & IIf(sWhere.Length > 0, " AND ", "") & "[氏名] Like '" & Replace(TextBox1.Text.Replace("*", "%") & "%'", "%%", "%")
        End If

        If CheckBox1.Checked AndAlso DateTimePicker1.Value < Now Then
            sWhere = sWhere & IIf(sWhere.Length > 0, " AND ", "") & "[生年月日] = #" & DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & "#"
        End If


        If sWhere.Length > 0 Then
            Dim pTBL As DataTable
            With SysAD.DB(sLRDB)
                pTBL = .GetTableBySqlSelect("SELECT *,[M_住民情報].ID AS [住民番号],[V_住民区分].名称 AS 住記区分 FROM [M_住民情報] LEFT JOIN [V_住民区分] ON [M_住民情報].[住民区分]=[V_住民区分].ID WHERE [住民区分] IN (0,1) AND " & sWhere)
            End With

            ListView1.Items.Clear()

            For Each pRow As DataRow In pTBL.Rows
                Dim pItem As ListViewItem = ListView1.Items.Add(pRow.Item("氏名"))

                pItem.SubItems.Add(pRow.Item("フリガナ").ToString)
                pItem.SubItems.Add(pRow.Item("住所").ToString)
                pItem.SubItems.Add(pRow.Item("住民番号").ToString)
                pItem.SubItems.Add(pRow.Item("世帯ID").ToString)

                If mvarTable IsNot Nothing Then
                    Dim pFRow As DataRow = mvarTable.Rows.Find(pRow.Item("住民番号"))
                    If pFRow IsNot Nothing Then
                        pItem.SubItems.Add("○")
                    Else
                        pItem.SubItems.Add("-")

                    End If
                End If
                pItem.SubItems.Add(pRow.Item("住記区分").ToString)
                pItem.SubItems.Add(pRow.Item("住記区分").ToString)
                If Not IsDBNull(pRow.Item("生年月日")) Then
                    pItem.SubItems.Add(和暦Format(pRow.Item("生年月日")))
                End If
            Next
        End If


        Dim ch As ColumnHeader
        For Each ch In ListView1.Columns
            ch.Width = -1
        Next ch
    End Sub

    Dim mvarTable As DataTable
    Public Sub New(ByVal pTable As DataTable)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        mvarTable = pTable
    End Sub

    Private Sub ListView1_ColumnClick(ByVal sender As Object, _
        ByVal e As ColumnClickEventArgs) Handles ListView1.ColumnClick
        'ListViewItemSorterを指定する
        ListView1.ListViewItemSorter = _
            New ListViewItemComparer(e.Column)
        '並び替える（ListViewItemSorterを設定するとSortが自動的に呼び出される）
        'ListView1.Sort()
    End Sub
    ''' <summary>
    ''' ListViewの項目の並び替えに使用するクラス
    ''' </summary>
    Public Class ListViewItemComparer
        Implements IComparer
        Private _column As Integer

        ''' <summary>
        ''' ListViewItemComparerクラスのコンストラクタ
        ''' </summary>
        ''' <param name="col">並び替える列番号</param>
        Public Sub New(ByVal col As Integer)
            _column = col
        End Sub

        'xがyより小さいときはマイナスの数、大きいときはプラスの数、
        '同じときは0を返す
        Public Function Compare(ByVal x As Object, ByVal y As Object) _
                As Integer Implements System.Collections.IComparer.Compare
            'ListViewItemの取得
            Dim itemx As ListViewItem = CType(x, ListViewItem)
            Dim itemy As ListViewItem = CType(y, ListViewItem)

            'xとyを文字列として比較する
            Return String.Compare(itemx.SubItems(_column).Text, _
                itemy.SubItems(_column).Text)
        End Function
    End Class


    Private Sub DateTimePicker1_ValueChanged(sender As System.Object, e As System.EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub dlg住記選択_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        DateTimePicker1.Value = Now.Date
    End Sub
End Class
