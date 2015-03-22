Public Class Exif

    Private _fileName As String = String.Empty
    Private _createDatetime As Date
    Private _jpg As System.Drawing.Bitmap = Nothing
    Private _items As System.Drawing.Imaging.PropertyItem
    Private _index() As Integer = Nothing

    Public Sub New()

    End Sub

    Public Sub New(fileName As String)
        Me.FileName = fileName
    End Sub

    Public ReadOnly Property Width As Integer
        Get
            Try
                Return BitConverter.ToInt32(_jpg.PropertyItems(Array.IndexOf(_index, 40962)).Value, 0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Height As Integer
        Get
            Try
                Return BitConverter.ToInt32(_jpg.PropertyItems(Array.IndexOf(_index, 40963)).Value, 0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property DpiX As Integer
        Get
            Try
                Return BitConverter.ToInt16(_jpg.PropertyItems(Array.IndexOf(_index, 282)).Value, 0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property DpiY As Integer
        Get
            Try
                Return BitConverter.ToInt16(_jpg.PropertyItems(Array.IndexOf(_index, 283)).Value, 0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property ISO As Integer
        Get
            Try
                Return BitConverter.ToInt16(_jpg.PropertyItems(Array.IndexOf(_index, 34855)).Value, 0)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Exposure As Integer
        Get
            Try
                Return BitConverter.ToInt16(_jpg.PropertyItems(Array.IndexOf(_index, 34850)).Value, 0)

            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property ProgramName As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 305)).Value)

            Catch ex As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property CameraDescription As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 42036)).Value)

            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property LatitudeRef As String
        Get
            Try
                'Return Byte2String((_jpg.PropertyItems(Array.IndexOf(_index, 1))).Value)
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 1)).Value)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Latitude As String
        Get
            Try

                Return IIf(Me.LatitudeRef = "N", "", "-") & Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 2))).Value, 0) & _
                    "." & Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 2))).Value, 8) & _
                     (Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 2))).Value, 16)).ToString.Replace(".", "")

            Catch ex As Exception
                Return Nothing
            End Try


        End Get
    End Property

    Public ReadOnly Property LongitudeRef As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 3)).Value)
            Catch ex As Exception
                Return Nothing

            End Try
        End Get
    End Property

    Public ReadOnly Property Longitude As String
        Get
            Try
                Return IIf(Me.LatitudeRef = "E", "", "-") & Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 4))).Value, 0) & _
               "." & Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 4))).Value, 8) & _
                (Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 4))).Value, 16)).ToString.Replace(".", "")
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property AltitudeRef As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 5)).Value)
            Catch ex As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property Altitude As String

        Get
            Try
                Return Byte2GIS((_jpg.PropertyItems(Array.IndexOf(_index, 6))).Value, 0)


            Catch ex As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property CreateDate As DateTime
        Get
            Try
                Dim s As String = System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 36867)).Value)
                s = s.Trim(New Char() {ControlChars.NullChar})
                Return DateTime.ParseExact(s, "yyyy:MM:dd HH:mm:ss", Nothing)

            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property DeviceMaker As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 271)).Value)
            Catch ex As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Public ReadOnly Property Artist As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 315)).Value)
                ' Return Byte2String((_jpg.PropertyItems(Array.IndexOf(_index, 315))).Value)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property DeviceName As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 272)).Value)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property DeviceFullName As String
        Get
            Return (Me.DeviceMaker & " " & Me.DeviceName).Trim

        End Get
    End Property

    Public Property FileName As String
        Get
            Return Me._fileName

        End Get
        Set(value As String)
            Try
                Dim f As New System.IO.FileInfo(value)
                If f.Exists Then
                    _jpg = New System.Drawing.Bitmap(value)
                    _index = _jpg.PropertyIdList

                Else
                    Throw New Exception("image file not found [" & value & "]")
                End If
                Me._fileName = value

            Catch ex As Exception
                Me._fileName = String.Empty
            End Try

        End Set

    End Property

    Public ReadOnly Property F As String
        Get
            Try
                Return System.Text.Encoding.ASCII.GetString(_jpg.PropertyItems(Array.IndexOf(_index, 64)).Value)
            Catch ex As Exception
                Return Nothing
            End Try

        End Get
    End Property

    Protected Overrides Sub Finalize()
        If _jpg IsNot Nothing Then _jpg.Dispose()
        MyBase.Finalize()
    End Sub

    Public Sub dump()

        If Me.FileName = String.Empty Then Return

        Dim f As New jp.polestar.io.Datatable2TSV
        Dim dt As New System.Data.DataTable("exif")

        dt.Columns.Add("filename", GetType(String))
        dt.Columns.Add("type", GetType(String))
        dt.Columns.Add("id", GetType(String))
        dt.Columns.Add("raw", GetType(Byte()))
        dt.Columns.Add("value", GetType(String))

        For Each i As Integer In _index
            Dim row As System.Data.DataRow = dt.NewRow
            Dim b() As Byte = _jpg.PropertyItems(Array.IndexOf(_index, i)).Value

            row(0) = Me.FileName
            row(1) = (_jpg.PropertyItems(Array.IndexOf(_index, i))).Type
            row(2) = i
            row(3) = b

            Select Case row(1)
                Case 1
                    ' ?
                    'row(4) = System.Text.Encoding.ASCII.GetString(b)
                Case 2
                    row(4) = System.Text.Encoding.ASCII.GetString(b)
                Case 3
                    row(4) = BitConverter.ToInt16(b, 0)
                Case 4
                    row(4) = BitConverter.ToInt32(b, 0)
                Case 5
                    For h As Integer = 0 To b.Length Step 4
                        If h < b.Length Then row(4) += BitConverter.ToInt16(b, h) & ":"
                    Next h
                    row(4) = row(4).ToString.Trim(":")

                Case 7
                    'row(4) = System.Text.Encoding.ASCII.GetString(b)
                Case 10
                    row(4) = BitConverter.ToInt64(b, 0)
            End Select
            dt.Rows.Add(row)
        Next
        f.dt2tsv(dt, "dumpExif.txt", True, io.FileAction.Overwrite)

    End Sub

    'GPSのデータをSingle値にする
    Private Function Byte2GIS(ByVal val() As Byte, ByVal ofs As Integer) As Single
        Dim aaa As Single = BitConverter.ToInt32(val, ofs)
        Dim bbb As Single = BitConverter.ToInt32(val, ofs + 4)
        Return aaa / bbb
    End Function

    Public Sub FlashGIS(Optional createBackupfile = True)

        Dim pi As System.Drawing.Imaging.PropertyItem
        Dim result() As Byte = Nothing
        Dim f As New System.IO.FileInfo(Me.FileName)
        Dim backupFileName As String = String.Empty

        If createBackupfile = True Then
            backupFileName = f.FullName.Substring(0, f.FullName.Length - f.Extension.Length) & "-back" & f.Extension
            f = New System.IO.FileInfo(backupFileName)
        Else
            backupFileName = f.FullName
        End If
        If f.Exists Then f.Delete()

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 1))
        ReDim result(1)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 2))
        ReDim result(23)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 3))
        ReDim result(1)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 4))
        ReDim result(23)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 5))
        ReDim result(1)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        pi = _jpg.PropertyItems(Array.IndexOf(_index, 6))
        ReDim result(7)
        pi.Value = result
        pi.Len = pi.Value.Length
        _jpg.SetPropertyItem(pi)

        _jpg.Save(backupFileName, System.Drawing.Imaging.ImageFormat.Jpeg)

        If f is Nothing then f=nothing

    End Sub


End Class
