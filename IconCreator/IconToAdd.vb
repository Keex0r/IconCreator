Public Class IconToAdd
    Public FileSize As Integer
    Public ImageBytes As Byte()
    Public ImageSize As Integer
    Public Shared Function Create(Source As Image, Size As Integer) As IconToAdd

        Using bmp As New Bitmap(Size, Size, Imaging.PixelFormat.Format32bppArgb)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(Source, 0, 0, Size, Size)
            End Using
            Using ms As New IO.MemoryStream
                bmp.Save(ms, Imaging.ImageFormat.Png)
                Dim ico As New IconToAdd With {.FileSize = CInt(ms.Length), .ImageBytes = ms.ToArray, .ImageSize = Size}
                Return ico
            End Using
        End Using
    End Function
End Class
