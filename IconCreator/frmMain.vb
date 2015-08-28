Public Class frmMain

    Private IconSizes As Integer() = {16, 24, 32, 48, 96, 128, 256, 512}
    Private CheckBoxes As List(Of CheckBox)
    Private SourceImage As Image

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        flpSizes.Controls.Clear()
        CheckBoxes = New List(Of CheckBox)
        For i = 0 To IconSizes.Count - 1
            Dim cb As New CheckBox
            cb.Text = IconSizes(i).ToString & "x" & IconSizes(i).ToString
            cb.AutoSize = True
            cb.Checked = True
            AddHandler cb.CheckedChanged, Sub(s, e1) UpdatePreview()
            CheckBoxes.Add(cb)
            cb.Tag = IconSizes(i)
            flpSizes.Controls.Add(cb)
        Next
    End Sub

    Private Sub UpdatePreview()
        For Each c As Control In flpPreview.Controls
            CType(c, PictureBox).Image.Dispose()
            c.Dispose()
        Next
        flpPreview.Controls.Clear()
        If SourceImage Is Nothing Then Exit Sub
        For Each cb In CheckBoxes
            If cb.Checked Then
                Dim size As Integer = CInt(cb.Tag)
                Dim pb As New PictureBox
                pb.Size = New Size(size, size)
                Dim bmp As New Bitmap(size, size)
                Using g As Graphics = Graphics.FromImage(bmp)
                    Dim filled As Boolean = True
                    For x = 0 To size Step 8
                        Dim filledfirst As Boolean = filled
                        For y = 0 To size Step 8
                            Dim brush As SolidBrush
                            If filled Then
                                brush = New SolidBrush(Color.Gray)
                            Else
                                brush = New SolidBrush(Color.White)
                            End If
                            g.FillRectangle(brush, x, y, 8, 8)
                            filled = Not filled
                        Next
                        filled = Not filledfirst
                    Next


                    g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    g.DrawImage(SourceImage, 0, 0, size, size)
                End Using
                pb.Image = bmp
                flpPreview.Controls.Add(pb)
            End If
        Next
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Using ofd As New OpenFileDialog
            ofd.Filter = "PNG Files|*.png"
            ofd.Multiselect = False
            If ofd.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub
            If SourceImage IsNot Nothing Then SourceImage.Dispose()
            SourceImage = New Bitmap(ofd.FileName)
            UpdatePreview()
            tbSource.Text = ofd.FileName
        End Using
    End Sub

    Private Sub CreateIcons()
        If SourceImage Is Nothing Then
            MessageBox.Show("Select a source image first!")
            Exit Sub
        End If
        Dim Filename As String = ""
        Using sfd As New SaveFileDialog
            sfd.Filter = "ICO File|*.ico"
            If sfd.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub
            Filename = sfd.FileName
        End Using
        Dim iconstoadd As New List(Of IconToAdd)
        For Each cb As CheckBox In CheckBoxes
            If cb.Checked Then
                Dim size = CInt(cb.Tag)
                iconstoadd.Add(IconToAdd.Create(SourceImage, size))
            End If
        Next
        Using fs As New IO.FileStream(Filename, IO.FileMode.Create)
            Using bw As New IO.BinaryWriter(fs)
                Dim ByteVal As Byte = 0
                Dim ShortVal As Short = 0
                Dim IntVal As Integer = 0
                bw.Write(ShortVal)
                ShortVal = 1
                bw.Write(ShortVal)
                ShortVal = CShort(iconstoadd.Count)
                bw.Write(ShortVal)
                '6 bytes am Anfang
                '16 bytes pro image entry
                '--> Start bei alle sizes davor additer+22 (-1 weil 0 basiert)
                Dim addiOffset As Integer = 6 + 16 * iconstoadd.Count
                For count = 0 To iconstoadd.Count - 1
                    Dim i = iconstoadd(count)

                    ByteVal = If(i.ImageSize < 256, CByte(i.ImageSize), CByte(0))
                    bw.Write(ByteVal) 'Width
                    bw.Write(ByteVal) 'Height
                    ByteVal = 0
                    bw.Write(ByteVal) 'Color Palette
                    bw.Write(ByteVal) 'Reserved
                    ShortVal = 0
                    bw.Write(ShortVal) 'Color planes (??)
                    ShortVal = 32
                    bw.Write(ShortVal) 'Bit depth
                    IntVal = i.FileSize
                    bw.Write(IntVal) 'Image size
                    IntVal = addiOffset
                    bw.Write(IntVal) 'Image position
                    addiOffset += i.FileSize
                Next
                'Write all images one after another to the file
                For count = 0 To iconstoadd.Count - 1
                    Dim i = iconstoadd(count)
                    bw.Write(i.ImageBytes)
                Next
            End Using
        End Using
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        CreateIcons()
    End Sub
End Class
