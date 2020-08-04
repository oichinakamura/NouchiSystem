
Imports System.Reflection
Imports System.ComponentModel
Imports System.Deployment.Application

Public Class CSystem
    Inherits HimTools2012.System管理.CSystemBase

    Public Property page農家世帯 As classPage農家世帯
    Public Property page申請処理 As classpage申請処理

    Public DatabaseProperty As New CDataBaseProperty
    Private mvarViews As New Dictionary(Of String, DataView)

    Private mvarMapConnection As New CMapConnection
    Public InterfaceSetting As CustomControlLIB.SettingParamSK


    Private mvar検索関連 As C検索関連
    Public 地図有無 As Boolean = False


    Public ReadOnly Property 検索関連 As C検索関連
        Get
            If mvar検索関連 Is Nothing Then
                mvar検索関連 = New C検索関連(Me)
            End If
            Return mvar検索関連
        End Get
    End Property

    Public ReadOnly Property MapConnection() As CMapConnection
        Get
            Return mvarMapConnection
        End Get
    End Property

    Public Overrides ReadOnly Property Application As HimTools2012.System管理.CApplication
        Get
            Return App農地基本台帳
        End Get
    End Property

    Public Function 共通書式フォルダ() As IO.DirectoryInfo
        Dim Path As String = Me.CommonReportFolder("")

        If Not IO.Directory.Exists(Path) Then
            IO.Directory.CreateDirectory(Path)
        End If

        Return New IO.DirectoryInfo(Path)
    End Function

    Public Function 市町村別書式フォルダ(s市町村名 As String) As IO.DirectoryInfo
        Return New IO.DirectoryInfo(SysAD.CustomReportFolder(s市町村名))
    End Function

    Public Sub New(ByRef pForm As frmMain, p農地基本台帳 As C農地基本台帳)
        MyBase.New("農地基本台帳", New CSystemInfo農地台帳("1500", ApplicationDeployment.IsNetworkDeployed), p農地基本台帳, pForm, "\農地基本台帳関連")

        AddHandler pForm.ViewMenu.DropDownItems.Add("オプション").Click, AddressOf Me.SystemInfo.ShowDlg

        App農地基本台帳 = p農地基本台帳
        Me.SystemInfo.ユーザー = New HimTools2012.System管理.CUser(1, "作業者", "", 1)
        Me.SetImageList(My.Resources.Resource1.ResourceManager, GetType(My.Resources.Resource1), True)

        Me.mvarSystemInfo.Load画面設定(My.Resources.Resource1.基本画面)

        DatabaseProperty.LoadPath()

        ObjectMan = New CObjectMan()
        Dim pDBX As HimTools2012.Data.CDataConnection = Nothing
        Me.InterfaceSetting = New CInterfaceSetting(Me)

        If Me.IsClickOnceDeployed Then
            Try
                Dim sFileName As String = Me.ClickOnceSetupPath() & "\SVADDR.TXT"
                If IO.File.Exists(sFileName) Then
                    pDBX = New HimTools2012.Data.CDataConnection(sLRDB)
                    pDBX.IPAddress = HimTools2012.TextAdapter.LoadTextFile(sFileName)
                    pDBX.DBMode = HimTools2012.Data.typeDBMode.ServerAccess
                Else
                    Dim sFileName2 As String = Me.ClickOnceSetupPath() & "\DBPath.TXT"
                    If IO.File.Exists(sFileName2) Then

                        Dim sPath As String = HimTools2012.TextAdapter.LoadTextFile(sFileName2)
                        DatabaseProperty.Path = sPath
                        pDBX = New HimTools2012.Data.CDataConnection(sLRDB)
                        pDBX.DBMode = HimTools2012.Data.typeDBMode.LoacalAccess
                        pDBX.LocalPath = sPath
                    Else
                        MsgBox("見つかりません　File2=" & sFileName2)
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            If Not IO.File.Exists("C:\Avail\農家台帳\SVADDR.TXT") AndAlso DatabaseProperty.IsDBExist() Then
                pDBX = New HimTools2012.Data.CDataConnection(sLRDB)
                pDBX.LocalPath = DatabaseProperty.Path
                pDBX.DBMode = HimTools2012.Data.typeDBMode.LoacalAccess
            ElseIf IO.File.Exists("C:\Avail\農家台帳\SVADDR.TXT") Then
                Dim sTXT As String = HimTools2012.TextAdapter.LoadTextFile("C:\Avail\農家台帳\SVADDR.TXT")

                Dim pDBTBL As DataTable = HimTools2012.Data.CDataConnection.GetDatabaseFiles(sTXT, "9001")

                Dim sList As New List(Of String)
                For Each pRow As DataRowView In New DataView(pDBTBL, "[名称] LIKE '*LRDB*'", "", DataViewRowState.CurrentRows)
                    sList.Add(pRow.Item("名称"))
                Next

                sLRDB = HimTools2012.CommonFunc.OptionSelect(Join(sList.ToArray, ";"), "選択してください.")
                If sTXT.Length > 0 AndAlso sLRDB.Length > 0 Then
                    pDBX = New HimTools2012.Data.CDataConnection(sLRDB)
                    pDBX.IPAddress = sTXT
                    pDBX.DBMode = HimTools2012.Data.typeDBMode.ServerAccess

                    s地図情報 = Replace(sLRDB, "LRDB", "地図情報")
                    Dim pV As New DataView(pDBTBL, "[名称]='" & s地図情報 & "'", "", DataViewRowState.CurrentRows)
                    If pV.Count > 0 Then
                        地図有無 = True
                        Dim pDBXT As New HimTools2012.Data.CDataConnection(s地図情報)
                        pDBXT.IPAddress = sTXT
                        pDBXT.DBMode = HimTools2012.Data.typeDBMode.ServerAccess
                        Me.DB.Add(s地図情報, pDBXT)
                    End If
                Else
                    End
                End If
            ElseIf DatabaseProperty.Path.Length = 0 Then
                With New OpenFileDialog
                    .Filter = "*.MDB|*.MDB"
                    If .ShowDialog() = DialogResult.OK Then
                        DatabaseProperty.SavePath(.FileName)
                        pDBX = New HimTools2012.Data.CDataConnection(sLRDB)
                        pDBX.LocalPath = DatabaseProperty.Path
                        pDBX.DBMode = HimTools2012.Data.typeDBMode.ServerAccess
                    End If
                End With
            End If

            HimTools2012.FileManager.MakeDir("C:\Avail\農家台帳")
        End If
        Me.DB.Add(sLRDB, pDBX)
        Me.DB.DefaultSelection = sLRDB

        For Each sName As String In {"農地基本台帳様式.xml", "旧農地基本台帳様式.xml"}
            If Not IO.File.Exists(Me.共通書式フォルダ.FullName & "\" & sName) Then
                IO.File.Copy(My.Application.Info.DirectoryPath & "\" & sName, Me.共通書式フォルダ.FullName & "\" & sName)
            End If
        Next
    End Sub

    Public Function Auto書式フォルダ(sFileName As String) As String
        If IO.File.Exists(Me.市町村別書式フォルダ(Me.市町村.市町村名).FullName & "\" & sFileName) Then
            Return Me.市町村別書式フォルダ(Me.市町村.市町村名).FullName & "\" & sFileName
        ElseIf IO.File.Exists(Me.共通書式フォルダ.FullName & "\" & sFileName) Then
            Return Me.共通書式フォルダ.FullName & "\" & sFileName
        ElseIf IO.File.Exists(My.Application.Info.DirectoryPath & "\" & sFileName) Then
            Return My.Application.Info.DirectoryPath & "\" & sFileName
        Else
            Return ""
        End If
    End Function



    Public ReadOnly Property MasterView(ByVal sClass As String) As DataView
        Get
            If mvarViews.ContainsKey(sClass) Then
                Return mvarViews.Item(sClass)
            Else
                Dim pView As New DataView(App農地基本台帳.DataMaster.Body, String.Format("[Class]='{0}'", sClass), "ID", DataViewRowState.CurrentRows)

                mvarViews.Add(sClass, pView)
                Return pView
            End If
        End Get
    End Property

    Public ReadOnly Property MasterFind(ByVal sClass As String, ByVal nID As Integer) As String
        Get

            Dim pView As New DataView(App農地基本台帳.DataMaster.Body, String.Format("[Class]='{0}' AND [ID]={1}", sClass, nID), "ID", DataViewRowState.CurrentRows)
            If pView.Count = 0 Then
                Return ""
            Else
                Return pView(0).Item("名称").ToString
            End If
        End Get
    End Property

    Public Overrides Function GetAssembly() As Object
        Return Assembly.GetExecutingAssembly()
    End Function


    Public Overrides Function InitSystem() As Boolean
        With SysAD.DB(sLRDB)
            App農地基本台帳.InitSystem()

            Try
                Dim s市町村名 As String = .DBProperty("市町村名")
                Dim pCLASSBASEType As Type = Type.GetType(String.Format("農地基本台帳.C{0}", s市町村名))
                mvar市町村 = CType(Activator.CreateInstance(pCLASSBASEType), C市町村別)
            Catch ex As Exception
                Stop
            End Try

        End With

        Return True
    End Function

    Public Function GetMapWindowHandle() As IntPtr
        For Each p As Process In Process.GetProcesses()
            If Not p.MainWindowHandle.Equals(IntPtr.Zero) Then
                If InStr(p.ProcessName, ":地図") Then
                    Return p.Handle
                End If
            End If
        Next

        Return 0
    End Function

    Public Sub OptionDlg(s As Object, e As EventArgs)
        mvarSystemInfo.ShowDlg()
    End Sub


    Public Overrides Sub InitDataSet()

    End Sub
End Class

Public Class CDataBaseProperty
    <Category("ファイル情報")>
    Private mvarPath As String = ""

    Public Property Path As String
        Get
            If mvarPath = "" Then
                LoadPath()
            End If
            Return mvarPath
        End Get
        Set(ByVal value As String)
            SavePath(value)
            mvarPath = value
        End Set
    End Property

    Public Function IsDBExist() As Boolean
        If Path.Length = 0 Then
            LoadPath()
        End If
        If Path.Length > 0 AndAlso Not IO.File.Exists(Path) Then
            Select Case MsgBox("指定されたパス[" & Path & "]が見つかりません。再度修正=はい、設定を解除する=いいえ、システムを終了する=キャンセル", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                    With New OpenFileDialog()
                        .Title = "データベース"
                        .Filter = "*.MDB|*.MDB"
                        If .ShowDialog = DialogResult.OK Then
                            Path = .FileName
                            SavePath(Path)
                            Return Me.IsDBExist()
                        Else
                            End
                        End If
                    End With
                Case MsgBoxResult.No
                    SavePath("")
                    End
                Case MsgBoxResult.Cancel
                    End
            End Select
        End If

        Return IO.File.Exists(Path)
    End Function

    Public Sub LoadPath()
        Path = GetSetting("Avail台帳管理", "農地・農家台帳", "Database", "")
    End Sub
    Public Sub SavePath(ByVal sPath As String)
        SaveSetting("Avail台帳管理", "農地・農家台帳", "Database", sPath)
    End Sub

End Class



