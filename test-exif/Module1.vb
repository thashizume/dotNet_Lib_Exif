Module Module1

    Sub Main()
        Dim ex As jp.polestar.image.Exif = Nothing

        ' E:\tmp\exif
        ' C:\Users\toshi\OneDrive\Pictures\カメラ ロール

        For Each f As System.IO.FileInfo In (New System.IO.DirectoryInfo("E:\tmp\exif").GetFiles)
            ex = New jp.polestar.image.Exif
            ex.FileName = f.FullName
            ex.dump()
            ex.FlashGIS()
            Console.WriteLine(f.FullName)
        Next

    End Sub

End Module
