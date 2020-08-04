Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.CodeDom.Compiler
Imports System.Reflection

Public Class dlgTest

    Private Sub OK_Click(sender As System.Object, e As System.EventArgs) Handles OK.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
    Private Sub Cancel_Click(sender As System.Object, e As System.EventArgs) Handles Cancel.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        X()
    End Sub
    Dim oDynamicCompiledInstance As Object
    Public Sub X()
        '動的に変更できるVBのプログラムソースコードを文字列化する
        Dim sbSource As New System.Text.StringBuilder
        sbSource.AppendLine("Imports Microsoft.VisualBasic")
        sbSource.AppendLine("Imports System.ComponentModel")
        sbSource.AppendLine("Public Class DynamicClass")
        sbSource.AppendLine("  <Browsable(True)>")
        sbSource.AppendLine("  Public Visible as boolean")
        sbSource.AppendLine("  Public Sub DynamicMethod(byref s1 As object,byval s2 As String)")
        sbSource.AppendLine("    MsgBox(typename(s1))")
        sbSource.AppendLine("  End Sub")
        sbSource.AppendLine("End Class")

        Dim sSource = sbSource.ToString()

        'コンパイルを実行する
        Dim oCompilerParameters As New CompilerParameters
        oCompilerParameters.GenerateExecutable = False
        oCompilerParameters.GenerateInMemory = True

        Dim oVBCompiler As New VBCodeProvider
        Dim oCompilerResults As CompilerResults
        oCompilerResults = oVBCompiler.CompileAssemblyFromSource(oCompilerParameters, sSource)

        '動的コンパイルしたソースにエラーがあった場合、エラーを表示して、終了
        If oCompilerResults.Errors.Count >= 1 Then
            Debug.Print("動的に与えられたソースコードにコンパイルエラーがありました。")
            For Each oCompilerError As System.CodeDom.Compiler.CompilerError In oCompilerResults.Errors
                Debug.Print(oCompilerError.ToString())
            Next
            MsgBox("ソースコードにエラーがあります。イミディエイトウィンドウにエラーを表示しました。")
            Exit Sub
        End If

        'コンパイルしたアセンブリをインスタンス化して実行
        Dim oDynamicCompiledAssembly As Assembly = oCompilerResults.CompiledAssembly
        Dim oDynamicCompiledClassType As Type = oDynamicCompiledAssembly.GetType("DynamicClass")
        Dim oDynamicCompiledMethodInfo As MethodInfo = oDynamicCompiledClassType.GetMethod("DynamicMethod")

        oDynamicCompiledInstance = Activator.CreateInstance(oDynamicCompiledClassType)
        PropertyGrid1.SelectedObject = oDynamicCompiledInstance

        oDynamicCompiledMethodInfo.Invoke(oDynamicCompiledInstance, New Object() {oDynamicCompiledInstance, "456789"})
        oDynamicCompiledInstance.Visible = True

    End Sub
  
    'マウスのクリック位置を記憶
    Private mousePoint As Point
    Private mouseMode As Integer = 0
    Private mvarSize As Size
    'Form1のMouseDownイベントハンドラ
    'マウスのボタンが押されたとき
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            mouseMode = -2 * (e.X > Me.ClientRectangle.Width - 10) - 4 * (e.Y > Me.ClientRectangle.Height - 10)

            Select Case mouseMode
                Case 0
                    mouseMode = 1
                    Me.Cursor = Cursors.Hand
                Case 2
                    Me.Cursor = Cursors.SizeWE
                Case 4
                    Me.Cursor = Cursors.SizeNS
                Case 6
                    Me.Cursor = Cursors.SizeAll
            End Select
            mvarSize = Me.Size
            mousePoint = New Point(e.X, e.Y)
        End If
    End Sub

    'Form1のMouseMoveイベントハンドラ
    'マウスが動いたとき
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            Select Case mouseMode
                Case 1
                    Me.Left += e.X - mousePoint.X
                    Me.Top += e.Y - mousePoint.Y
                Case 2
                    Me.Width = mvarSize.Width + (e.X - mousePoint.X)
                Case 4
                    Me.Height = mvarSize.Height + (e.Y - mousePoint.Y)
                Case 6
                    Me.Width = mvarSize.Width + (e.X - mousePoint.X)
                    Me.Height = mvarSize.Height + (e.Y - mousePoint.Y)
            End Select
        End If
    End Sub

    Private Sub dlgTest_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        mouseMode = 0
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub dlgTest_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        Me.SuspendLayout()
        Dim rect As New Rectangle(0, 0, Me.Width, Me.Height)
        Dim graphicPath As New System.Drawing.Drawing2D.GraphicsPath
        Dim radius As Integer = 20

        graphicPath.StartFigure()

        graphicPath.AddArc(rect.Left, rect.Top, radius * 2, radius * 2, 180, 90)
        graphicPath.AddLine(rect.Left + radius, rect.Top, rect.Right - radius, rect.Top)
        graphicPath.AddArc(rect.Right - radius * 2, rect.Top, radius * 2, radius * 2, 270, 90)
        graphicPath.AddLine(rect.Right, rect.Top + radius, rect.Right, rect.Bottom - radius)
        graphicPath.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90)
        graphicPath.AddLine(rect.Right - radius, rect.Bottom, rect.Left + radius, rect.Bottom)
        graphicPath.AddArc(rect.Left, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90)
        graphicPath.AddLine(rect.Left, rect.Bottom - radius, rect.Left, rect.Top + radius)
        graphicPath.CloseFigure()

        '描画先とするImageオブジェクトを作成する
        Dim canvas As New Bitmap(rect.Width, rect.Height)
        Dim g As Graphics = Graphics.FromImage(canvas)

        '縦に白から黒へのグラデーションのブラシを作成
        Dim gb As New LinearGradientBrush(g.VisibleClipBounds,
            Color.LightBlue,
            Color.Navy,
            LinearGradientMode.Horizontal)

        '四角を描く
        g.FillRectangle(gb, g.VisibleClipBounds)
        Dim NavyPen As New Pen(Color.Navy, 2)

        g.DrawPath(NavyPen, graphicPath)
        NavyPen.Dispose()

        'リソースを解放する
        gb.Dispose()
        g.Dispose()

        'PictureBox1に表示する
        Me.BackgroundImage = canvas
        Me.Region = New Region(graphicPath)
        Me.ResumeLayout()
    End Sub




    'Private Sub CalenderCtrl1_ClickDayItem(s As Object, pDate As DateTime, pButton As System.Windows.Forms.MouseButtons) Handles CalenderCtrl1.ClickDayItem
    '    If pButton = Windows.Forms.MouseButtons.Right Then
    '        Dim Dt As DateTime = pDate
    '        If MsgBox(Dt.ToString("yyyy/MM/dd") & "を休日に設定しますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
    '            Me.Visible = False
    '            Dim sText As String = InputBox("休日(祝日)名を入力してください", "名称入力")
    '            Me.Visible = True
    '            If sText.Length > 0 Then
    '                Dim pRow As DataRow = CType(CalenderCtrl1.PublicHolidaySource, DataTable).NewRow
    '                pRow.Item("日時") = Dt
    '                pRow.Item("内容") = sText
    '                With CType(CalenderCtrl1.PublicHolidaySource, DataTable)
    '                    .Rows.Add(pRow)
    '                    If SysAD.IsClickOnceDeployed Then
    '                        .WriteXml(SysAD.ClickOnceSetupPath & "\HolidayTable.xml", System.Data.XmlWriteMode.WriteSchema)
    '                    Else
    '                        .WriteXml(My.Application.Info.DirectoryPath & "\HolidayTable.xml", System.Data.XmlWriteMode.WriteSchema)
    '                    End If
    '                End With
    '            End If
    '        End If
    '    End If
    'End Sub


   

    Public Sub New()
        InitializeComponent()

    
    End Sub


End Class
