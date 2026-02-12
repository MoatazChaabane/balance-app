Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.IO

Public Class Form1
    Dim _client As TcpClient
    Dim lastReceivedTime As DateTime
    Dim pingTimer As Timer
    Dim ip As String
    Dim port As Integer
    Dim filePath As String

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function SetThreadExecutionState(esFlags As EXECUTION_STATE) As EXECUTION_STATE
    End Function

    <Flags()>
    Private Enum EXECUTION_STATE As UInteger
        ES_CONTINUOUS = &H80000000UI
        ES_SYSTEM_REQUIRED = &H1
        ES_AWAYMODE_REQUIRED = &H40
    End Enum

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS Or EXECUTION_STATE.ES_SYSTEM_REQUIRED Or EXECUTION_STATE.ES_AWAYMODE_REQUIRED)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SetThreadExecutionState(EXECUTION_STATE.ES_CONTINUOUS)
        pingTimer?.Dispose()
        _client?.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ip = TextBox1.Text
        port = Integer.Parse(TextBox2.Text)
        filePath = GetFilePathFromIP(ip)
        CheckForIllegalCrossThreadCalls = False

        Try
            _client = New TcpClient(ip, port)
            lastReceivedTime = DateTime.Now

            MsgBox("Connexion réussie")

            pingTimer = New Timer(AddressOf SendPing, Nothing, Timeout.Infinite, Timeout.Infinite)
            ThreadPool.QueueUserWorkItem(AddressOf ReceiveLoop)

            Button1.Text = "Balance connectée"
            Button1.Enabled = False
        Catch ex As Exception
            MsgBox("Erreur de connexion : " & ex.Message)
            ThreadPool.QueueUserWorkItem(AddressOf ReceiveLoop)
        End Try
    End Sub

    Private Sub SendPing(state As Object)
        Try
            If _client Is Nothing OrElse Not _client.Connected Then
                Reconnect()
                Return
            End If

            If (DateTime.Now - lastReceivedTime).TotalSeconds > 10 Then
                Dim ns As NetworkStream = _client.GetStream()
                Dim pingMessage As Byte() = Encoding.ASCII.GetBytes("ping")
                ns.Write(pingMessage, 0, pingMessage.Length)
            Else
                pingTimer.Change(Timeout.Infinite, Timeout.Infinite)
            End If
        Catch
            Reconnect()
        End Try
    End Sub

    Private Sub ReceiveLoop(state As Object)
        While True
            Try
                If _client Is Nothing OrElse Not _client.Connected Then
                    Reconnect()
                End If

                Dim ns As NetworkStream = _client.GetStream()
                ns.ReadTimeout = 15000

                Dim toReceive(1024) As Byte
                Dim bytesRead As Integer = ns.Read(toReceive, 0, toReceive.Length)

                If bytesRead > 0 Then
                    lastReceivedTime = DateTime.Now
                    pingTimer.Change(Timeout.Infinite, Timeout.Infinite)

                    Dim txt As String = Encoding.ASCII.GetString(toReceive, 0, bytesRead).Trim()
                    UpdateUI(txt)
                    SaveToFile(txt)
                End If

                pingTimer.Change(10000, 5000)
            Catch ex As IOException
                Reconnect()
            Catch ex As Exception
                Reconnect()
            End Try
        End While
    End Sub

    Private Sub Reconnect()
        Do
            Try
                _client?.Close()
            Catch
            End Try

            Try
                _client = New TcpClient(ip, port)
                lastReceivedTime = DateTime.Now
                Exit Do
            Catch
                Thread.Sleep(5000)
            End Try
        Loop
    End Sub

    Private Sub UpdateUI(txt As String)
        If RichTextBox1.TextLength > 0 Then
            RichTextBox1.AppendText(vbNewLine & txt)
        Else
            RichTextBox1.Text = txt
        End If
    End Sub

    Private Sub SaveToFile(txt As String)
        If Not String.IsNullOrWhiteSpace(txt) Then
            Try
                Using file As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(filePath, True)
                    file.WriteLine(txt.Replace(" ", "").Replace(Environment.NewLine, ""))
                End Using
            Catch
                ' Ignorer les erreurs de fichier
            End Try
        End If
    End Sub

    Private Function GetFilePathFromIP(ip As String) As String
        Select Case ip
            Case "192.168.245.111"
                Return "C:\balance\906\1\export\article.txt"
            Case "192.168.245.112"
                Return "C:\balance\906\2\export\article.txt"
            Case "192.168.92.111"
                Return "C:\balance\905\1\export\article.txt"
            Case "192.168.92.114"
                Return "C:\balance\905\2\export\article.txt"
            Case Else
                Return "C:\balance\default\export\article.txt"
        End Select
    End Function

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
    End Sub
End Class
