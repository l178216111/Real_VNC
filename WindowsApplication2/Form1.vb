Imports System.IO
Public Class Form1
    Dim a, b, c, d, mac, part, cl, st, version, currentversion As String
    'Dim Wmi As New System.Management.ManagementObjectSearcher("SELECT * FROM Win32_NetworkAdapterConfiguration")
    Dim Conn As OleDb.OleDbConnection
    Dim Cmd As OleDb.OleDbCommand
    Dim Rd As OleDb.OleDbDataReader
    Dim SQL As String
    Public Declare Auto Function RegisterHotKey Lib "user32.dll" Alias "RegisterHotKey" (ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Boolean
    Public Declare Auto Function UnRegisterHotKey Lib "user32.dll" Alias "UnregisterHotKey" (ByVal hwnd As IntPtr, ByVal id As Integer) As Boolean
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As IntPtr) As IntPtr
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
    Dim Database = "Data source=" & Application.StartupPath & "\Database.mdb"
    '  Dim Database = "Data source=\\10.192.130.43\Probe\Test_probe\Equipment_Engineering\6.0 Equip Staff folder\Teradyne team\LiuZX\Database4.mdb"
    Sub link()
        Conn = New OleDb.OleDbConnection(Provider & ";" & Database)
        Conn.Open()
        SQL = "Select IP,passcode ,Tool From " & part & " where Tool ='" & TextBox1.Text & "'"
        Cmd = New OleDb.OleDbCommand(SQL, Conn)
        Rd = Cmd.ExecuteReader()
        While (Rd.Read())
            If Not IsDBNull(Rd.GetValue(0)) Then
                b = Rd.GetValue(0)
                c = Rd.GetValue(1)
            End If
        End While
        Rd.Close()
        Conn.Close()
    End Sub
    Sub change()
        If c = "welcome" Then
            d = "lifei88"
        Else
            d = "welcome"
        End If
        Conn = New OleDb.OleDbConnection(Provider & ";" & Database)
        Conn.Open()
        SQL = "UPDATE " & part & " set passcode='" & d & "' where Tool ='" & TextBox1.Text & "'"
        Cmd = New OleDb.OleDbCommand(SQL, Conn)
        Cmd.ExecuteNonQuery()
        Conn.Close()
    End Sub
    Sub task()
        Dim hdwne As Integer
        hdwne = FindWindow("rfb::win32::DesktopWindowClass", vbNullString)
        SendMessage(hdwne, 16, 0, 0)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Call task()
    End Sub
    Sub Button()
        Call link()
        Shell(Application.StartupPath & "\vncviewer.exe", AppWinStyle.Hide)
        ' Shell("\\10.192.130.43\Probe\Test_probe\Equipment_Engineering\6.0 Equip Staff folder\Teradyne team\LiuZX\vncviewer")
        Dim hdWn As Integer
        Dim hdwn1, hdwn2, hdwn3, hdwn4, x2, x1, i As Integer
        Do While (hdWn = 0 Or hdwn1 = 0 Or hdwn2 = 0 Or hdwn3 = 0)
            hdWn = FindWindow("#32770", "VNC Viewer : Connection Details")
            hdwn1 = FindWindowEx(hdWn, 0, "Combobox", vbNullString)
            hdwn3 = FindWindowEx(hdWn, 0, vbNullString, "ok")
            hdwn2 = FindWindowEx(hdwn1, 0, "Edit", vbNullString)
        Loop
        SendMessage(hdwn2, 12, 0, b)
        SendMessage(hdwn3, 245, 0, 0)
        hdwn1 = 0
        hdwn2 = 0
        hdwn3 = 0
        x2 = 0
        While (hdwn1 = 0 Or hdwn2 = 0 Or hdwn3 = 0 Or hdwn4 = 0)
            i = i + 1
            If i > 10000 Then
                st = 1
                Exit While
            End If
            st = 0
            x2 = FindWindow("#32770", "VNC Viewer : Error")
            hdwn3 = FindWindowEx(hdwn1, 0, "button", "ok")
            If x1 <> 0 Or x2 <> 0 Then
                st = 1
                Exit While
            End If
            x1 = FindWindow("#32770", "VNC Viewer : Question")
            hdwn1 = FindWindow("#32770", "VNC Viewer : Authentication [No Encryption]")
            hdwn2 = FindWindowEx(hdwn1, 0, "Edit", vbNullString)
            hdwn4 = FindWindowEx(hdwn1, hdwn2, "Edit", vbNullString)
        End While
        SendMessage(hdwn4, 12, 0, c)
        SendMessage(hdwn3, 245, 0, 0)
    End Sub
    Sub passcode()

    End Sub
 
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call ReadTextFromFile()
        Call WriteTextToFile()
        part = "J750"
        Dim isResult1 As Boolean
        isResult1 = RegisterHotKey(Handle, 0, 0, Keys.F1)
        If isResult1 = False Then
            MsgBox("Press F1 to show your VNC")
            End
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim count As Integer = System.Text.Encoding.Default.GetByteCount(Me.TextBox1.Text)
        If count = 2 Then
            Call task()
            Call Button()
            If st = 0 Then
                Call mos()
            End If
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Static i
        i += 1
        Select Case (i Mod 3)
            Case 0
                Button1.Text = "J750"
                part = "J750"
            Case 1
                Button1.Text = "MST"
                part = "MST"
            Case 2
                Button1.Text = "FLEX"
                part = "FLEX"
        End Select
    End Sub
    Sub mos()
        Dim pos, z1, z2, z3, z4, z5, z6, i As Integer
        pos = 0
        z1 = 0
        z2 = 0
        z3 = 0
        z4 = 0
        While (pos = 0)
            pos = FindWindow("rfb::win32::DesktopWindowClass", vbNullString)
            z1 = FindWindow("#32770", "VNC Viewer : Question")
            If z1 <> 0 Then
                Call change()
                While (z2 = 0)
                    z2 = FindWindowEx(z1, 0, "button", "&Yes")
                End While
                SendMessage(z2, 245, 0, 0)
                While (z3 = 0 Or z4 = 0 Or z5 = 0 Or z6 = 0)
                    i = i + 1
                    If i > 10000 Then
                        st = 1
                        Exit While
                    End If
                    z3 = FindWindow("#32770", "VNC Viewer : Authentication [No Encryption]")
                    z4 = FindWindowEx(z3, 0, "Edit", vbNullString)
                    z5 = FindWindowEx(z3, z4, "Edit", vbNullString)
                    z6 = FindWindowEx(z3, 0, "button", "ok")
                End While
                SendMessage(z5, 12, 0, d)
                SendMessage(z6, 245, 0, 0)
                While (pos = 0)
                    pos = FindWindow("rfb::win32::DesktopWindowClass", vbNullString)
                End While
                Exit While
            End If
        End While
        SetForegroundWindow(pos)
        Me.TopMost = True
    End Sub
    Protected Overrides Sub WndProc(ByRef m As Message)
        Dim too As Integer
        Dim a As Boolean
        If m.Msg = 786 Then
            If Visible Then
                Hide()
                a = False
            Else
                Show()
                a = True
            End If
            If a = True Then
                While (too = 0)
                    too = FindWindow(vbNullString, "Real VNC-LZ")
                End While
                SetForegroundWindow(too)
                Me.TextBox1.Focus()
            End If
        End If

        MyBase.WndProc(m)
    End Sub
    Private Sub Form1_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Dim isResult As Boolean
        isResult = UnRegisterHotKey(Handle, 0)
        'SetHotkey(1, "", "Del")
        'SetHotkey(2, "", "Del")
    End Sub
    Sub WriteTextToFile()
        If currentversion.Trim <> version.Trim Then
            FileCopy("\\10.192.130.43\Probe\Test_probe\Equipment_Engineering\2.4 Equip Performance\J750 team\VNC\Database.mdb", Application.StartupPath & "\Database.mdb")
            Dim file As New System.IO.StreamWriter(Application.StartupPath & "\version.txt")
            file.WriteLine(version)
            file.Close()
            MsgBox("Update New Datebase Done", MsgBoxStyle.ApplicationModal)
        End If
    End Sub
    Sub ReadTextFromFile()
        Dim file As New System.IO.StreamReader("\\10.192.130.43\Probe\Test_probe\Equipment_Engineering\2.4 Equip Performance\J750 team\VNC\version.txt")
        version = file.ReadToEnd().Trim
        file.Close()
        Dim read As New System.IO.StreamReader(Application.StartupPath & "\version.txt")
        currentversion = read.ReadToEnd().Trim
        read.Close()
    End Sub
End Class