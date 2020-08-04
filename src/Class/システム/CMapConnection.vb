Public Class CMapConnection
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Int32, ByVal lpString As String) As UInt32
    Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Integer, ByVal lpString As String, ByVal hData As Int32) As Integer
    Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Integer, ByVal lpString As String) As Integer
    Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

    Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
    Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Int32) As Int32
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Integer) As Integer

    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Integer, ByVal hwndChildAfter As Integer, ByVal lpszClass As String, ByVal lpszWindow As String) As Integer
    Private hWnd As IntPtr = 0

    Public Function SelectMap(ByVal sIDList As String) As Object
        Try
            For Each p As Process In Process.GetProcesses
                If p.ProcessName = "地図システム" Then
                    hWnd = p.MainWindowHandle
                    ListToMap(sIDList)
                    SetForegroundWindow(p.MainWindowHandle)

                    Exit For
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Nothing
    End Function

    Public Function HasMap() As Boolean
        For Each p As Process In Process.GetProcesses
            If InStr(p.MainWindowTitle, "地図情報システム") AndAlso p.ProcessName = "地図システム" Then

                hWnd = p.MainWindowHandle
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub BookToMap(ByVal hwnd As Integer, ByVal sMessage As String)
        Try
            Dim lngAtom As Integer = 0
            Dim strText As String

            If hwnd Then
                Try
                    lngAtom = GetProp(hwnd, "BookToMap")

                    If lngAtom Then
                        GlobalDeleteAtom(lngAtom)
                    End If
                Catch ex As Exception

                End Try
                strText = sMessage

                Dim bName() As Byte = System.Text.Encoding.ASCII.GetBytes("BookToMap")
                Dim bAscii() As Byte = System.Text.Encoding.ASCII.GetBytes(sMessage)
                strText = System.Text.Encoding.Default.GetString(bAscii)


                lngAtom = GlobalAddAtom(strText)

                If lngAtom <> 0 Then
                    SetProp(hwnd, System.Text.Encoding.Default.GetString(bName), lngAtom)
                End If
            End If
        Catch ex As Exception
            MsgBox(":=" & ex.Message)
        End Try
    End Sub

    Private Sub ListToMap(ByVal sMessage As String)
        Dim pErr As String = ""
        Try


            Dim strText As String = ""

            If sMessage.Length > 0 Then
                Dim sPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) & "\MapAndBook"

                If Not IO.Directory.Exists(sPath) Then
                    IO.Directory.CreateDirectory(sPath)
                End If

                Dim sFileName As String = sPath & "\BookToMap.txt"

                If MsgBox("地図を呼びだしますか", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    HimTools2012.TextAdapter.SaveTextFile(sFileName, sMessage, "Ascii")
                End If

                My.Application.DoEvents()


            End If
        Catch ex As Exception
            MsgBox(pErr & ":=" & ex.Message)
        End Try
    End Sub

    Public Function GetPropStr(ByVal lpString As String, Optional ByVal buffLen As Integer = 255) As String
        Dim lngAtom As Integer
        Dim strText As String
        If HasMap() Then
            lngAtom = GetProp(hWnd, lpString)
            If lngAtom Then
                strText = Space(buffLen)
                GlobalGetAtomName(lngAtom, strText, buffLen)
                GlobalDeleteAtom(lngAtom)
                RemoveProp(hWnd, lpString)
                Try
                    Return Trim(Left$(strText, InStr(strText, Chr(0)) - 1))

                Catch ex As Exception

                End Try
            End If
        End If

        Return ""
    End Function

    Public Function FindWindow(ByVal hWndParent As Integer, ByVal hwndChildAfter As Integer, Optional ByVal sClass As String = "", Optional ByVal sWindow As String = "") As Integer
        If sClass = "" Then sClass = vbNullString
        If sWindow = "" Then sWindow = vbNullString
        FindWindow = FindWindowEx(hWndParent, hwndChildAfter, sClass, sWindow)
    End Function
End Class
