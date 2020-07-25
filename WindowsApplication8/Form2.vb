Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging


Public Class Form2
    Inherits System.Windows.Forms.Form

    Private cam As SnapShot.Capture
    Private m_ip As IntPtr = IntPtr.Zero

    Const VIDEODEVICE As Integer = 0         ' zero based index of video capture device to use
    Const VIDEOWIDTH As Integer = 640        ' Depends on video device caps
    Const VIDEOHEIGHT As Integer = 480       ' Depends on video device caps
    Const VIDEOBITSPERPIXEL As Integer = 24  ' BitsPerPixel values determined by device

    Private Sub prepare()
        cam = New SnapShot.Capture(VIDEODEVICE, VIDEOWIDTH, VIDEOHEIGHT, VIDEOBITSPERPIXEL, PictureBox2)
    End Sub

    Private Sub button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.prepare()
    End Sub

    Private Sub takePic()
        Cursor.Current = Cursors.WaitCursor

        ' Release any previous buffer
        If m_ip <> IntPtr.Zero Then
            Marshal.FreeCoTaskMem(m_ip)
            m_ip = IntPtr.Zero
        End If

        ' capture image
        m_ip = cam.Click()
        Dim b As New Bitmap(cam.Width, cam.Height, cam.Stride, PixelFormat.Format24bppRgb, m_ip)

        ' If the image is upsidedown
        b.RotateFlip(RotateFlipType.RotateNoneFlipY)
        PictureBox1.Image = b

        Cursor.Current = Cursors.[Default]
    End Sub

    Private Sub button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        Me.takePic()
    End Sub

    Private Sub closeDx()
        Try
            cam.Dispose()

            If m_ip <> IntPtr.Zero Then
                Marshal.FreeCoTaskMem(m_ip)
                m_ip = IntPtr.Zero
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Me.closeDx()
    End Sub

    
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        prepare()
        System.Threading.Thread.Sleep(1000)
        takePic()
        closeDx()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class