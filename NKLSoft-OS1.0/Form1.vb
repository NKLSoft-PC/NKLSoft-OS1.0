Imports System.Net
Imports System.Management.Automation
Imports System.Management.Automation.Runspaces
Imports System.Collections.ObjectModel
Imports System.Text

Public Class Form1

    Dim Offset As Point
    Private WithEvents MyProcess As Process
    Private Delegate Sub AppendOutputTextDelegate(ByVal text As String)

    Private Sub CloseMenu_Click(sender As Object, e As EventArgs) Handles CloseMenu.Click
        MenuPanel.Visible = False
        OpenMenu.Visible = True
        CloseMenu.Visible = False

    End Sub

    Private Sub OpenMenu_Click(sender As Object, e As EventArgs) Handles OpenMenu.Click
        MenuPanel.Visible = True
        OpenMenu.Visible = False
        CloseMenu.Visible = True

    End Sub


    Private Sub MyProcess_ErrorDataReceived(ByVal sender As Object, ByVal e As System.Diagnostics.DataReceivedEventArgs) Handles MyProcess.ErrorDataReceived
        AppendOutputText(vbCrLf & "Error: " & e.Data)
    End Sub

    Private Sub MyProcess_OutputDataReceived(ByVal sender As Object, ByVal e As System.Diagnostics.DataReceivedEventArgs) Handles MyProcess.OutputDataReceived
        AppendOutputText(vbCrLf & e.Data)
    End Sub

    Private Sub ExecuteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExecuteButton.Click
        MyProcess.StandardInput.WriteLine(InputTextBox.Text)
        MyProcess.StandardInput.Flush()
        InputTextBox.Text = ""
    End Sub

    Private Sub AppendOutputText(ByVal text As String)
        Try
            If OutputTextBox.InvokeRequired Then
                Dim myDelegate As New AppendOutputTextDelegate(AddressOf AppendOutputText)
                Me.Invoke(myDelegate, text)
            Else
                OutputTextBox.AppendText(text)
            End If
        Catch ex As Exception
            If OutputTextBox.InvokeRequired Then
                Dim myDelegate As New AppendOutputTextDelegate(AddressOf AppendOutputText)
                Me.Invoke(myDelegate, text)
            Else
                OutputTextBox.AppendText(text)
            End If
        End Try

    End Sub


    Private Sub MenuBox_MouseDown(sender As Object, e As MouseEventArgs) Handles MenuBox.MouseDown

    End Sub

    Private Sub MoveBox1_Click(sender As Object, e As EventArgs) Handles MoveBox1.Click

    End Sub

    Private Sub MoveBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles MoveBox1.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim Pos As Point = Me.PointToClient(MousePosition)
            Pos.Offset(Offset.X, Offset.Y)
            MenuPanel.Location = Pos
            Me.Text = "MoveBox1 is at " & Pos.X.ToString & "," & Pos.Y.ToString
            Dim ViewableAreaLocation As Point = New Point(Pos.X + Math.Abs(Offset.X), Pos.Y + Math.Abs(Offset.Y))
            Me.Text &= " Mouse pointer is at " & ViewableAreaLocation.X.ToString & "," & ViewableAreaLocation.Y.ToString
        End If
    End Sub

    Private Sub MoveBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles MoveBox1.MouseDown
        Offset = New Point(-e.X, -e.Y)
    End Sub

    Private Sub XLABEL_Click(sender As Object, e As EventArgs) Handles XLABEL.Click
        Application.Exit()
    End Sub



    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles Me.Click

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized
        Me.TopMost = True
        Timer1.Start()

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TIMELBL.Text = "{ " + TimeOfDay.ToString("hh:mm:ss tt") + "}"
        DATELBL.Text = "{ " + Date.Now.Date.Date + "}"
        DATELBL2.Text = "{ " + Date.Now.Date.Date + "}"
        TIMELBL2.Text = "{ " + TimeOfDay.ToString("hh:mm:ss tt") + "}"
        TIMELBL3.Text = "{ " + TimeOfDay.ToString("hh:mm:ss tt") + "}"
    End Sub



    Private Sub OpenCalBox_Paint(sender As Object, e As PaintEventArgs)




    End Sub

    Private Sub CloseCalBox_Paint(sender As Object, e As PaintEventArgs)



    End Sub

    Private Sub DATELBL_Click(sender As Object, e As EventArgs) Handles DATELBL.Click
        CalandarBox.Visible = True
        DATELBL.Visible = False
        TIMELBL2.Visible = False
        DATELBL2.Visible = True
        TIMELBL3.Visible = True

    End Sub

    Private Sub TIMELBL2_Click(sender As Object, e As EventArgs) Handles TIMELBL2.Click
        CalandarBox.Visible = True
        DATELBL.Visible = False
        TIMELBL2.Visible = False
        DATELBL2.Visible = True
        TIMELBL3.Visible = True
    End Sub

    Private Sub DATELBL2_Click(sender As Object, e As EventArgs) Handles DATELBL2.Click
        CalandarBox.Visible = False
        DATELBL.Visible = True
        TIMELBL2.Visible = True
        DATELBL2.Visible = False
        TIMELBL3.Visible = False


    End Sub

    Private Sub TIMELBL3_Click(sender As Object, e As EventArgs) Handles TIMELBL3.Click
        CalandarBox.Visible = False
        DATELBL.Visible = True
        TIMELBL2.Visible = True
        DATELBL2.Visible = False
        TIMELBL3.Visible = False
    End Sub

    Private Sub QUITSHELLBOX_Click(sender As Object, e As EventArgs) Handles QUITSHELLBOX.Click
        SHELLBOX.Visible = False
        OpenShellBox.Visible = True
        CloseShellBox.Visible = False
        OutputTextBox.Clear()
        MyProcess.Close()
        OutputTextBox.Clear()
    End Sub

    Private Sub OpenShellBox_Click(sender As Object, e As EventArgs) Handles OpenShellBox.Click
        MenuPanel.Visible = False
        OpenMenu.Visible = True
        CloseMenu.Visible = False
        Me.AcceptButton = ExecuteButton
        MyProcess = New Process
        With MyProcess.StartInfo
            .FileName = "CMD.EXE"
            .UseShellExecute = False
            .CreateNoWindow = True
            .RedirectStandardInput = True
            .RedirectStandardOutput = True
            .RedirectStandardError = True
        End With
        MyProcess.Start()

        MyProcess.BeginErrorReadLine()
        MyProcess.BeginOutputReadLine()
        'AppendOutputText("Process Started at: " & MyProcess.StartTime.ToString)

        '  MyProcess.StandardInput.WriteLine("ipconfig") 'send an EXIT command to the Command Prompt
        '  MyProcess.StandardInput.Flush()
        SHELLBOX.Visible = True
        OpenShellBox.Visible = False
        CloseShellBox.Visible = True
        Dim Pos As Point = Me.PointToClient(MousePosition)
        Pos.Offset(Offset.X, Offset.Y)
        SHELLBOX.Location = Pos
        POS1.Text = Pos.X.ToString & "::" & Pos.Y.ToString
        Dim ViewableAreaLocation As Point = New Point(Pos.X + Math.Abs(Offset.X), Pos.Y + Math.Abs(Offset.Y))
        POS1.Text &= ViewableAreaLocation.X.ToString & "::" & ViewableAreaLocation.Y.ToString
    End Sub

    Private Sub CloseShellBox_Click(sender As Object, e As EventArgs) Handles CloseShellBox.Click
        SHELLBOX.Visible = False
        OpenShellBox.Visible = True
        CloseShellBox.Visible = False
        OutputTextBox.Clear()
        MyProcess.Close()
        OutputTextBox.Clear()


    End Sub



    Private Sub MoveShellBox_MouseMove(sender As Object, e As MouseEventArgs) Handles MoveShellBox.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Dim Pos As Point = Me.PointToClient(MousePosition)
            Pos.Offset(Offset.X, Offset.Y)
            SHELLBOX.Location = Pos
            'POS1.Text = Pos.X.ToString & "::" & Pos.Y.ToString
            Dim ViewableAreaLocation As Point = New Point(Pos.X + Math.Abs(Offset.X), Pos.Y + Math.Abs(Offset.Y))
            ' POS1.Text &= ViewableAreaLocation.X.ToString & "::" & ViewableAreaLocation.Y.ToString
        End If
    End Sub

    Private Sub MoveShellBox_MouseDown(sender As Object, e As MouseEventArgs) Handles MoveShellBox.MouseDown
        Offset = New Point(-e.X, -e.Y)
    End Sub

    Private Sub BackIMG_Click(sender As Object, e As EventArgs) Handles BackIMG.Click
        MenuPanel.Visible = False
        OpenMenu.Visible = True
        CloseMenu.Visible = False
    End Sub

    Private Sub CloseMenuBar_Click(sender As Object, e As EventArgs) Handles CloseMenuBar.Click
        MenuPanel.Visible = False
        OpenMenu.Visible = True
        CloseMenu.Visible = False
    End Sub

    Private Sub RunPS1SCRIPTFromURLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RunPS1SCRIPTFromURLToolStripMenuItem.Click
        'iex ((New-Object System.Net.WebClient).DownloadString("https://www.dropbox.com/s/b0qun9euotd576k/NKLBrowser.txt?dl=1"))

        If PSURLCOMBOBOX1.Text = "" Then

        Else
            MyProcess.StandardInput.WriteLine("start /wait powershell -exec bypass -command iex ((New-Object System.Net.WebClient).DownloadString('" + PSURLCOMBOBOX1.Text + "'))") 'send an EXIT command to the Command Prompt
            MyProcess.StandardInput.Flush()
        End If

    End Sub

    Private Sub GetIPV4ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetIPV4ToolStripMenuItem.Click

        MyProcess.StandardInput.WriteLine("ipconfig |find " + UNION.Text + "IPv4" + UNION.Text + "")
        MyProcess.StandardInput.Flush()
    End Sub

    Private Sub PSRunToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PSRunToolStripMenuItem.Click
        txtOutput.Clear()
        Dim openFileDialog As OpenFileDialog = New OpenFileDialog
        If (openFileDialog.ShowDialog = DialogResult.OK) Then
            Dim powerShell As PowerShell = PowerShell.Create
            powerShell.AddScript(System.IO.File.ReadAllText(openFileDialog.FileName))
            powerShell.AddCommand("Out-String")
            Dim PSOutput As Collection(Of PSObject) = powerShell.Invoke
            Dim stringBuilder As StringBuilder = New StringBuilder
            For Each pSObject As PSObject In PSOutput
                stringBuilder.AppendLine(pSObject.ToString)
            Next
            txtOutput.Text = stringBuilder.ToString
        End If
    End Sub

    Private Sub OpenPS1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenPS1ToolStripMenuItem.Click
        txtOutput.Clear()
        Dim openFileDialog As OpenFileDialog = New OpenFileDialog
        If (openFileDialog.ShowDialog = DialogResult.OK) Then
            Dim powerShell As PowerShell = PowerShell.Create
            powerShell.AddScript(System.IO.File.ReadAllText(openFileDialog.FileName))
            powerShell.AddCommand("Out-String")
            Dim PSOutput As Collection(Of PSObject) = powerShell.Invoke
            Dim stringBuilder As StringBuilder = New StringBuilder
            For Each pSObject As PSObject In PSOutput
                stringBuilder.AppendLine(pSObject.ToString)
            Next
            txtOutput.Text = stringBuilder.ToString
        End If
    End Sub
End Class
