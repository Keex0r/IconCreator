Public Class frmMain

#Region "Constructor"
    Public Sub New()

        InitializeComponent()

        'Create checkboxes for each supported icon size
        flpSizes.Controls.Clear()
        CheckBoxes = New List(Of CheckBox)
        For i = 0 To IconSizes.Count - 1
            Dim cb As New CheckBox
            'Text should be WidthxHeight
            cb.Text = IconSizes(i).ToString & "x" & IconSizes(i).ToString
            cb.AutoSize = True
            cb.Checked = True
            'Update the preview when the checkbox is checked/unchecked
            AddHandler cb.CheckedChanged, Sub(s, e1) UpdatePreview()
            'Hold the size the checkbox represents in the Tag for easier usage
            cb.Tag = IconSizes(i)
            'Add to array and display on the form
            CheckBoxes.Add(cb)
            flpSizes.Controls.Add(cb)
        Next
    End Sub
#End Region


#Region "Private Fields"
    ''' <summary>
    ''' Holds the supported icon sizes.
    ''' </summary>
    Private IconSizes As Integer() = {16, 24, 32, 48, 96, 128, 256, 512}
    ''' <summary>
    ''' Holds the created checkboxes for the icon sizes
    ''' </summary>
    Private CheckBoxes As List(Of CheckBox)
    ''' <summary>
    ''' Holds the source image that the icons are created from
    ''' </summary>
    Private SourceImage As Image
#End Region

#Region "Image Preview"
    ''' <summary>
    ''' Updates the icon preview section by adding pictureboxes to <see cref="flpPreview"/>
    ''' </summary>
    Private Sub UpdatePreview()
        'Dispose and clear the previous pictureboxes
        For Each c As Control In flpPreview.Controls
            Dim pb = CType(c, PictureBox)
            If pb.Image IsNot Nothing Then pb.Image.Dispose()
            c.Dispose()
        Next
        flpPreview.Controls.Clear()

        'Only continue when a valid image is selected
        If SourceImage Is Nothing Then Exit Sub

        'Create the previews for all selected icon sizes
        For Each cb In CheckBoxes
            If cb.Checked Then
                'Get the icon size
                Dim size As Integer = CInt(cb.Tag)

                'Pictureboxes and Image will be disposed when the preview is next updated
                Dim pb As New PictureBox
                pb.Size = New Size(size, size)
                Dim bmp As New Bitmap(size, size)

                Using g As Graphics = Graphics.FromImage(bmp)
                    'Draw a checkerboard pattern to indicate transparency
                    Dim filled As Boolean = True
                    For x = 0 To size Step 8
                        'Makes sure that the pattern color switches back and fourth at the beginning of each column
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

                    'Set the GDI+ Image drawing to high quality interpolation
                    g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                    g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                    'Draw the image in the correct size
                    g.DrawImage(SourceImage, 0, 0, size, size)
                End Using
                pb.Image = bmp
                'Display on the form
                flpPreview.Controls.Add(pb)
            End If
        Next
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        SelectSource()
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        CreateIcons()
    End Sub
#End Region

#Region "Source Selection"
    ''' <summary>
    ''' Lets the user select a new source image from the explorer
    ''' </summary>
    Private Sub SelectSource()
        Using ofd As New OpenFileDialog
            ofd.Filter = "Image Files|*.png;*.bmp;*.jpg;*.jpeg;*.tif"
            ofd.Multiselect = False
            If ofd.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub
            If SourceImage IsNot Nothing Then SourceImage.Dispose()
            'Update form and field
            SourceImage = New Bitmap(ofd.FileName)
            UpdatePreview()
            tbSource.Text = ofd.FileName
        End Using
    End Sub
#End Region
#Region "Icon Creation"
    ''' <summary>
    ''' Actually creates the icon from the source. It implements the .ICO format for Microsoft Windows Vista and above
    ''' where the icon images are stored as complete .PNG files in the container.
    ''' </summary>
    Private Sub CreateIcons()
        If SourceImage Is Nothing Then
            MessageBox.Show("Select a source image first!", "No source image selected", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        'Select a filename to save the icon to
        Dim Filename As String = ""
        Using sfd As New SaveFileDialog
            sfd.Filter = "ICO File|*.ico"
            If sfd.ShowDialog(Me) = DialogResult.Cancel Then Exit Sub
            Filename = sfd.FileName
        End Using

        'First the PNG image data needs to be created, because we need to now the size of each
        'image to correctly construct the header information. The data is created in the IconToAdd.Create() method.
        Dim IconsToAdd As New List(Of IconToAdd)
        For Each cb As CheckBox In CheckBoxes
            If cb.Checked Then
                Dim size = CInt(cb.Tag)
                IconsToAdd.Add(IconToAdd.Create(SourceImage, size))
            End If
        Next
        Try
            'Open the file to write to for binary writing
            Using fs As New IO.FileStream(Filename, IO.FileMode.Create)
                Using bw As New IO.BinaryWriter(fs)
                    'Helper variables to make use of Option Infer
                    Dim ByteVal As Byte = 0
                    Dim IntVal As Integer = 0
                    Dim ShortVal As Short = 0
                    '.ICO always starts with a Short 0
                    bw.Write(ShortVal)
                    'File type: 1=ICO, 2=CUR
                    ShortVal = 1
                    bw.Write(ShortVal)
                    'Count of icons contained in the container
                    ShortVal = CShort(IconsToAdd.Count)
                    bw.Write(ShortVal)

                    'Now the header information for each contained icon need to be written.
                    'This contains the offset where each PNG image is found in the container.
                    '6 Bytes for the ICO Header
                    '16 bytes per image entry
                    'Add the length of each written image
                    '--> Image is located at addiOffset
                    Dim addiOffset As Integer = 6 + 16 * IconsToAdd.Count
                    For count = 0 To IconsToAdd.Count - 1
                        Dim i = IconsToAdd(count)

                        'Size of the image: for sizes > 255 we write 0
                        ByteVal = If(i.ImageSize < 256, CByte(i.ImageSize), CByte(0))
                        bw.Write(ByteVal) 'Width
                        bw.Write(ByteVal) 'Height

                        'Color Palette
                        ByteVal = 0
                        bw.Write(ByteVal)
                        'Reserved
                        bw.Write(ByteVal)
                        'Color planes
                        ShortVal = 0
                        bw.Write(ShortVal)
                        'Bit depth, we always create 32bpp images
                        ShortVal = 32
                        bw.Write(ShortVal)
                        'Length of PNG data
                        IntVal = i.FileSize
                        bw.Write(IntVal)
                        'Offset of PNG data in the container
                        IntVal = addiOffset
                        bw.Write(IntVal)
                        'Increment offset counter by this PNG's file size
                        addiOffset += i.FileSize
                    Next
                    'Write all images one after another to the file
                    For count = 0 To IconsToAdd.Count - 1
                        Dim i = IconsToAdd(count)
                        bw.Write(i.ImageBytes)
                    Next
                End Using
            End Using
            MessageBox.Show("The icon was created successfully!", "Sucess", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("An error was encountered during the creation of the ICON:" & vbCrLf & vbCrLf &
                ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region


End Class
