Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class Form1
    Inherits Form

    Public Sub New()
        InitializeComponent()

        ' Load the blurred desktop wallpaper as the background of Form1
        Me.BackgroundImage = ApplyBlurEffect()
        Me.BackgroundImageLayout = ImageLayout.Stretch
    End Sub

    Private Function CaptureWindowScreenshot() As Bitmap
        ' Create a bitmap to store the captured window screenshot
        Dim windowScreenshot As New Bitmap(Me.Width, Me.Height)

        ' Capture the window screenshot
        Using g As Graphics = Graphics.FromImage(windowScreenshot)
            g.CopyFromScreen(Me.Location, New Point(0, 0), Me.Size)
        End Using

        Return windowScreenshot
    End Function

    Private Sub ApplySimpleBlur(ByRef image As Bitmap, Optional ByVal BlurForce As Integer = 1)
        ' Declare an ImageAttributes to use it when drawing
        Dim att As New ImageAttributes
        ' Declare a ColorMatrix
        Dim m As New ColorMatrix
        ' Set Matrix33 to 0.5, which represents the opacity, making the drawing semi-transparent.
        m.Matrix33 = 0.5F
        ' Setting this ColorMatrix to the ImageAttributes.
        att.SetColorMatrix(m)

        ' We get a graphics object from the image
        Dim g As Graphics = Graphics.FromImage(image)

        ' Drawing the image on itself, but not in the same coordinates, 
        ' in a way that every pixel will be drawn on the pixels around it.
        For x = -BlurForce To BlurForce Step 24
            For y = -BlurForce To BlurForce Step 24
                ' Drawing image on itself using our ImageAttributes to draw it semi-transparent.
                g.DrawImage(image, New Rectangle(x, y, image.Width, image.Height), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, att)
            Next
        Next

        ' Disposing ImageAttributes and Graphics. The effect is then applied. 
        att.Dispose()
        g.Dispose()
    End Sub

    Private Function ApplyBlurEffect() As Bitmap
        ' Capture the desktop screenshot
        Dim windowScreenshot As Bitmap = CaptureWindowScreenshot()

        ' Apply blur effect to the captured image
        ApplySimpleBlur(windowScreenshot, 32) ' Adjust blur force as needed

        Return windowScreenshot
    End Function
    Private Sub UpdateBackground()
        ' Capture the window screenshot covered by the form
        Dim windowScreenshot As Bitmap = CaptureWindowScreenshot()

        ' Apply blur effect to the captured image
        ApplySimpleBlur(windowScreenshot, 32) ' Adjust blur force as needed

        ' Set the blurred image as the background of Form1
        Me.BackgroundImage = windowScreenshot
    End Sub

    Private backgroundUpdateCounter As Integer = 0

    Private Sub Form1_LocationChanged(sender As Object, e As EventArgs) Handles Me.LocationChanged
        ' Reset the counter when the form location changes
        backgroundUpdateCounter = 0

        ' Update the background immediately
        UpdateBackground()

        ' Start the loop to update the background an additional 5 times
        Dim timer As New Timer()
        AddHandler timer.Tick, AddressOf Timer_Tick
        timer.Interval = 100 ' Adjust the interval as needed
        timer.Start()
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs)
        ' Update the background while the counter is less than 5
        If backgroundUpdateCounter < 10 Then
            UpdateBackground()
            backgroundUpdateCounter += 1
        Else
            ' Stop the timer when the counter reaches 5
            Dim timer As Timer = DirectCast(sender, Timer)
            timer.Stop()
            timer.Dispose()
        End If
    End Sub

End Class
