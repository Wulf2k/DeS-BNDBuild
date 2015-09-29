Imports System.IO
Imports System.IO.Compression


Public Class DesBNDBuild
    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String

    Public Shared extractPath As String

    Public Shared BinderID As String
    Public Shared flags As UInteger
    Public Shared numFiles As UInteger
    Public Shared namesEndLoc As UInteger


    Private Function StrFromBytes(ByVal loc As UInteger) As String
        Dim Str As String = ""
        Dim cont As Boolean = True

        While cont
            If bytes(loc) > 0 Then
                Str = Str + Convert.ToChar(bytes(loc))
                loc += 1
            Else
                cont = False
            End If
        End While

        Return Str
    End Function
    Private Function UIntFromBytes(ByVal loc As UInteger) As UInteger
        Dim tmpUint As UInteger = 0

        For i = 0 To 3
            tmpUint += Convert.ToUInt16(bytes(loc + i)) * &H100 ^ (3 - i)
        Next

        Return tmpUint
    End Function

    Private Sub UINTToBytes(ByVal loc As UInteger, val As UInteger)
        Dim BArr(3) As Byte

        For i = 0 To 3
            BArr(i) = Math.Floor(val / (&H100 ^ (3 - i))) Mod &H100
        Next

        Array.Copy(BArr, 0, bytes, loc, 4)
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim openDlg As New OpenFileDialog()

        openDlg.Filter = "DeS DCX/BND File|*BND;*DCX"
        openDlg.Title = "Open your BND file"

        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtBNDfile.Text = openDlg.FileName
        End If


    End Sub

    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        Dim currFileSize As UInteger
        Dim currFileOffset As UInteger
        Dim currFileID As UInteger
        Dim currFileNameOffset As UInteger
        Dim currFileName As String
        Dim currFilePath As String
        Dim currFileBytes() As Byte = {}

        Dim fileList As String = ""

        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        bytes = File.ReadAllBytes(filepath & filename)

        BinderID = StrFromBytes(&H0)
        flags = UIntFromBytes(&HC)
        numFiles = UIntFromBytes(&H10)
        namesEndLoc = UIntFromBytes(&H14)

        For i As UInteger = 0 To numFiles - 1
            currFileSize = UIntFromBytes(&H24 + i * &H18)
            currFileOffset = UIntFromBytes(&H28 + i * &H18)
            currFileID = UIntFromBytes(&H2C + i * &H18)
            currFileNameOffset = UIntFromBytes(&H30 + i * &H18)

            currFileName = StrFromBytes(currFileNameOffset)

            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

            fileList += currFileID & "," & currFileName & Environment.NewLine

            If (Not System.IO.Directory.Exists(currFilePath)) Then
                System.IO.Directory.CreateDirectory(currFilePath)
            End If


            ReDim currFileBytes(currFileSize - 1)
            Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
            File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
        Next
        File.WriteAllText(filepath & filename & ".extract\FileList.txt", fileList)
        txtInfo.Text += TimeOfDay & " - Extracted." & Environment.NewLine
    End Sub
    Private Sub btnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        Dim currFileSize As UInteger
        Dim currFileOffset As UInteger
        Dim currFileNameOffset As UInteger
        Dim currFileName As String
        Dim currFilePath As String
        Dim currFileBytes() As Byte = {}

        Dim padding As UInteger

        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        bytes = File.ReadAllBytes(filepath & filename)

        BinderID = StrFromBytes(&H0)
        flags = UIntFromBytes(&HC)
        numFiles = UIntFromBytes(&H10)
        namesEndLoc = UIntFromBytes(&H14)

        If namesEndLoc Mod &H10 > 0 Then
            padding = &H10 - (namesEndLoc Mod &H10)
        Else
            padding = 0
        End If


        ReDim Preserve bytes(namesEndLoc + padding - 1)


        For i As UInteger = 0 To numFiles - 1
            currFileNameOffset = UIntFromBytes(&H30 + i * &H18)

            currFileName = StrFromBytes(currFileNameOffset)
            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

            currFileBytes = File.ReadAllBytes(currFilePath & currFileName)
            currFileSize = currFileBytes.Length

            If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                padding = &H10 - (currFileSize Mod &H10)
            Else
                padding = 0
            End If

            currFileOffset = bytes.Length
            UINTToBytes(&H24 + i * &H18, currFileSize)
            UINTToBytes(&H28 + i * &H18, currFileOffset)
            UINTToBytes(&H34 + i * &H18, currFileSize)

            ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)
            Array.Copy(currFileBytes, 0, bytes, currFileOffset, currFileSize)
        Next
        File.WriteAllBytes(filepath & filename, bytes)
        txtInfo.Text += TimeOfDay & " - Rebuilt." & Environment.NewLine
    End Sub

    Public Function Decompress(ByVal cmpBytes() As Byte) As Byte()
        Dim sourceFile As MemoryStream = New MemoryStream(cmpBytes)
        Dim destFile As MemoryStream = New MemoryStream()
        Dim compStream As New DeflateStream(sourceFile, CompressionMode.Decompress)
        Dim myByte As Integer = compStream.ReadByte()

        While myByte <> -1
            destFile.WriteByte(CType(myByte, Byte))
            myByte = compStream.ReadByte()
        End While

        destFile.Close()
        sourceFile.Close()

        Return destFile.ToArray
    End Function
    Private Sub btnDecompress_Click(sender As Object, e As EventArgs) Handles btnDecompress.Click
        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        bytes = File.ReadAllBytes(filepath & filename)


        Dim newbytes(&H44D9) As Byte
        Dim decbytes(&H10000) As Byte

        Array.Copy(bytes, &H120, newbytes, 0, &H44D9)

        decbytes = Decompress(newbytes)


    End Sub
    Private Sub btnCompress_Click(sender As Object, e As EventArgs) Handles btnCompress.Click
        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        bytes = File.ReadAllBytes(filepath & filename)
        Dim newbytes() As Byte


        Dim stream = New MemoryStream()
        Dim zipStream = New DeflateStream(stream, CompressionLevel.Optimal)
        zipStream.Write(bytes, 0, 65536)
        zipStream.Close()

        newbytes = stream.ToArray()




        File.WriteAllBytes(filepath & filename & ".test", newbytes)
    End Sub

    Private Sub txt_Drop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        sender.Text = file(0)
    End Sub
    Private Sub txt_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub


    Private Sub btnBackup_Click(sender As Object, e As EventArgs) Handles btnBackup.Click
        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        If Not File.Exists(filepath & filename & ".bak") Then
            bytes = File.ReadAllBytes(filepath & filename)
            File.WriteAllBytes(filepath & filename & ".bak", bytes)
            txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
        Else
            txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
        End If
    End Sub
End Class
