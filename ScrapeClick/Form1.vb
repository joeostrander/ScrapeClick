Imports System.IO
Imports System.Text
Imports System.Net

Public Class Form1

    Private intCheckInterval As Integer = 1 'minutes
    Private dtLastCheck As DateTime
    Private ElapsedTime As TimeSpan
    Private SourceDir As String = New DirectoryInfo(Directory.GetCurrentDirectory).FullName & "\"
    Private strFilename As String = SourceDir & "desktop.bmp"
    Private strReferenceFile As String = SourceDir & "reference.bmp"
    'Private strReferenceFile As String = "c:\temp\reference.bmp"

    Public boolTrayExit As Boolean = False

    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, _
            ByVal lpKeyName As String, _
            ByVal lpDefault As String, _
            ByVal lpReturnedString As StringBuilder, _
            ByVal nSize As Integer, _
            ByVal lpFileName As String) As Integer

    Declare Sub mouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

    Const LEFTDOWN As Integer = &H2
    Const LEFTUP As Integer = &H4


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not IO.File.Exists(strReferenceFile) Then
            MsgBox("Reference file not found:" & vbCrLf & vbCrLf & strReferenceFile, MsgBoxStyle.Exclamation)
            boolTrayExit = True
            Application.Exit()
        End If

        System.Threading.Thread.Sleep(5000)
        ScanNow()
        dtLastCheck = Now
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ElapsedTime = Now.Subtract(dtLastCheck)
        LabelNextCheck.Text = "Next Check:  " & FormatNumber(intCheckInterval * 60 - ElapsedTime.TotalSeconds, 0).ToString & " seconds"
        If ElapsedTime.TotalMinutes >= intCheckInterval Then ScanNow()
    End Sub

    Private Sub ScanNow()
        Try

            LabelStatus.Text = Now & " - " & "Scanning..."
            Debug.WriteLine(Now & vbTab & "Scanning pixels...")


            Dim SC As New ScreenShot.ScreenCapture
            Dim bmpDesktop As Bitmap = SC.CaptureScreen
            Dim bmpReference As Bitmap = Bitmap.FromFile(strReferenceFile)

            bmpDesktop.Save(strFilename, Imaging.ImageFormat.Bmp)


            dtLastCheck = Now

            ''might not be needed... make bitmap from file w/o locking file:
            'Dim MyBitmap As Bitmap = New Bitmap(bmpTemp.Width, bmpTemp.Height)
            'Dim gr As Graphics = Graphics.FromImage(bmpTemp)
            'gr.DrawImage(MyBitmap, 0, 0)
            'bmpTemp.Dispose()


            Dim width_desktop As Integer = bmpDesktop.Width
            Dim height_desktop As Integer = bmpDesktop.Height

            Dim width_reference As Integer = bmpReference.Width
            Dim height_reference As Integer = bmpReference.Height

            Dim boolFound As Boolean = False

            For y_desk = 0 To height_desktop - 1 - height_reference
                If boolFound = True Then Exit For

                For x_desk = 0 To width_desktop - 1 - width_reference
                    'Set a flag - when it doesn't match set false
                    Dim boolMatch As Boolean = True

                    'Compare to reference bitmap
                    Dim y_ref As Integer = 0
                    Do While y_ref < height_reference - 1
                        For x_ref = 0 To width_reference - 1
                            'compare the reference pixel with the offset pixel of the desktop
                            'Get the current desktop pixel color
                            Dim color_pixel_desktop As Color = bmpDesktop.GetPixel(x_desk + x_ref, y_desk + y_ref)
                            Dim color_pixel_reference As Color = bmpReference.GetPixel(x_ref, y_ref)

                            If color_pixel_desktop.ToString = color_pixel_reference.ToString Then
                                'Debug.WriteLine("match: " & x_desk & "," & y_desk)
                            Else
                                'Debug.WriteLine(x_desk & "," & y_desk & "..." & color_pixel_desktop.ToString & " <> " & color_pixel_reference.ToString)
                                boolMatch = False
                                Exit Do
                            End If
                        Next
                        y_ref += 1
                    Loop

                    If boolMatch = True Then
                        boolFound = True
                        Debug.WriteLine(Now & vbTab & "Found at:  " & x_desk & "x" & y_desk)
                        Me.Cursor = New Cursor(Cursor.Current.Handle)   'may not be needed...
                        Dim oldPosition As Point = New Point(Cursor.Position)
                        Cursor.Position = New Point(x_desk, y_desk)

                        'MsgBox("Found at:  " & x_desk & "x" & y_desk, MsgBoxStyle.Information)
                        DoClick()
                        LabelStatus.Text = Now & " - " & "Clicked at " & x_desk & "x" & y_desk
                        Cursor.Position = oldPosition
                        Exit For
                    End If

                Next
            Next


            If boolFound = False Then
                Debug.WriteLine(Now & vbTab & "Couldn't find it")
                LabelStatus.Text = Now & " - " & "Not found"
            End If

        Catch ex As Exception
            LabelStatus.Text = ex.Message
        End Try

        Debug.WriteLine(Now & vbTab & "Done scanning.")

    End Sub

    Private Sub DoClick()
        mouse_event(LEFTDOWN, 0, 0, 0, 0)
        mouse_event(LEFTUP, 0, 0, 0, 0)


    End Sub

    Private Function GetIniValue(ByVal section As String, ByVal key As String) As String
        Dim sb As StringBuilder = New StringBuilder(500)
        Dim ret As Integer = GetPrivateProfileString(section, key, "", sb, sb.Capacity, SourceDir & "settings.ini")
        Return sb.ToString.Trim
    End Function



    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If boolTrayExit = False Then
            e.Cancel = True
            Me.WindowState = FormWindowState.Minimized
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        boolTrayExit = True
        Application.Exit()
    End Sub


    Private Sub NotifyIcon1_BalloonTipClicked(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.BalloonTipClicked
        ShowMe()
    End Sub

    Private Sub ShowMe()
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.Activate()
    End Sub

    Private Sub NotifyIcon1_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick
        If e.Button = MouseButtons.Right Then
            'NotifyIcon1.ContextMenu = ContextMenuIcon
        End If
        If e.Button = MouseButtons.Left Then
            ShowMe()
        End If

    End Sub

    Private Sub Form1_Move(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Move

        If Me.WindowState = FormWindowState.Minimized Then
            'NotifyIcon1.ShowBalloonTip(600, Application.ProductName, "Click for details", ToolTipIcon.Info)
            Me.Hide()
        Else
            Me.Show()
        End If
    End Sub


    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        If Me.WindowState = FormWindowState.Minimized Then Me.Hide()
    End Sub


    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    
End Class
