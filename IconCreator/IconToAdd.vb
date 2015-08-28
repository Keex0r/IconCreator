''' <summary>
''' Holds information about a PNG icon to add to the container
''' </summary>
Public Class IconToAdd
    ''' <summary>
    ''' The size of the PNG Data
    ''' </summary>
    Public FileSize As Integer
    ''' <summary>
    ''' The PNG binary data
    ''' </summary>
    Public ImageBytes As Byte()
    ''' <summary>
    ''' The Width/Height of this image
    ''' </summary>
    Public ImageSize As Integer

    ''' <summary>
    ''' Creates a new IconToAdd
    ''' </summary>
    ''' <param name="Source">The Source Image</param>
    ''' <param name="Size">The size of the icon</param>
    ''' <returns></returns>
    Public Shared Function Create(Source As Image, Size As Integer) As IconToAdd

        Using bmp As New Bitmap(Size, Size, Imaging.PixelFormat.Format32bppArgb)
            Using g As Graphics = Graphics.FromImage(bmp)
                'Set image interpolation to high quality
                g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(Source, 0, 0, Size, Size)
            End Using
            'Save the image in PNG format in a memorystream to extract the byte data
            Using ms As New IO.MemoryStream
                bmp.Save(ms, Imaging.ImageFormat.Png)
                'Return a new IconToAdd object
                Dim ico As New IconToAdd With {.FileSize = CInt(ms.Length), .ImageBytes = ms.ToArray, .ImageSize = Size}
                Return ico
            End Using
        End Using
    End Function
End Class
