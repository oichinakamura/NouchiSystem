

Public Class CadastralMaps
    Inherits HimTools2012.controls.CTabPageWithToolStrip
    Private WithEvents mvarMap As New PictureBox


    Public Sub New()
        MyBase.New(True, True, "地籍図", "地籍図")
        mvarMap.Dock = DockStyle.Fill
        mvarMap.BackColor = Color.Black
        Me.ControlPanel.Add(mvarMap)
    End Sub

    Public Draw農地Col As New Dictionary(Of Long, CObj農地)

    Private nBTop As Long
    Private nBLeft As Long

    Public Sub DrawLandBoundary(ByRef p農地 As CObj農地)
        Dim pView As New DataView(App農地基本台帳.TBL筆情報.Body, "[XID]=" & p農地.ID, "", DataViewRowState.CurrentRows)
        If pView.Count > 0 Then
        Else
            Dim pTBLX As DataTable = SysAD.DB(s地図情報).GetTableBySqlSelect("SELECT * FROM [D:LotProperty] WHERE [OAZA]=" & p農地.大字ID & " AND [Name]='" & p農地.地番 & "'")
            For Each pRow As DataRow In pTBLX.Rows
                pRow.Item("XID") = p農地.ID
            Next
            App農地基本台帳.TBL筆情報.MergePlus(pTBLX)
        End If
        For Each pRV As DataRowView In pView
            Dim pXY As DataTable = SysAD.DB(s地図情報).GetTableBySqlSelect("SELECT * FROM [D:XY] WHERE [X]>={0} AND [X]<={1} AND [Y]>={2} AND [Y]<={3}", pRV("MinX"), pRV("MaxX"), pRV("MinY"), pRV("MaxY"))
            App農地基本台帳.TBL点情報.Merge(pXY, False, MissingSchemaAction.AddWithKey)

            nBTop = pRV("MinX") / 100
            nBLeft = pRV("MinY") / 100

            Dim sList As String = pRV.Item("List").ToString

            If sList.Length > 0 Then
                For i As Integer = 1 To sList.Length Step 8
                    Dim n As Integer = Val("&H" & HimTools2012.StringF.Mid(sList, i, 8))

                    Dim pXYRow As DataRow = App農地基本台帳.TBL点情報.Rows.Find(n)
                    If pXYRow IsNot Nothing Then
                        p農地.構成点.Add(New Point(pXYRow.Item("Y") / 100 - nBLeft, pXYRow.Item("X") / 100 - nBTop))
                    End If
                Next
                Draw農地Col.Add(p農地.ID, p農地)
            End If

        Next
    End Sub

    Public Overrides ReadOnly Property PageCloseMode As HimTools2012.controls.CloseMode
        Get
            Return HimTools2012.controls.CloseMode.NoMessage
        End Get
    End Property

    Private Sub mvarMap_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles mvarMap.Paint

        e.Graphics.TranslateTransform(0, mvarMap.Height - 1)
        e.Graphics.ScaleTransform(0.1, -0.1)

        '/
        For Each p農地 As CObj農地 In Draw農地Col.Values
            e.Graphics.DrawPolygon(Pens.White, p農地.構成点.ToArray)
        Next
    End Sub
End Class
