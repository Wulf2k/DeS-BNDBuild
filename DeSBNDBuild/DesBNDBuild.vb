Imports System.IO
Imports System.IO.Compression


Public Class DesBNDBuild
    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String
    Public Shared extractPath As String


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

    Private Sub StrToBytes(ByVal str As String, ByVal loc As UInteger)
        Dim BArr() As Byte
        BArr = System.Text.Encoding.ASCII.GetBytes(str)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub
    Private Sub InsBytes(ByVal bytes2() As Byte, ByVal loc As UInteger)
        Array.Copy(bytes2, 0, bytes, loc, bytes2.Length)
    End Sub
    Private Sub UINTToBytes(ByVal val As UInteger, loc As UInteger)
        Dim BArr(3) As Byte

        For i = 0 To 3
            BArr(i) = Math.Floor(val / (&H100 ^ (3 - i))) Mod &H100
        Next

        Array.Copy(BArr, 0, bytes, loc, 4)
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim openDlg As New OpenFileDialog()

        openDlg.Filter = "DeS DCX/BND File|*BND;*MOWB;*DCX"
        openDlg.Title = "Open your BND file"

        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtBNDfile.Text = openDlg.FileName
        End If


    End Sub

    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        Dim currFileName As String = ""
        Dim currFilePath As String = ""
        Dim fileList As String = ""

        Dim BinderID As String = ""
        Dim namesEndLoc As UInteger = 0
        Dim flags As UInteger = 0
        Dim numFiles As UInteger = 0

        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        bytes = File.ReadAllBytes(filepath & filename)

        If Not File.Exists(filepath & filename & ".bak") Then
            bytes = File.ReadAllBytes(filepath & filename)
            File.WriteAllBytes(filepath & filename & ".bak", bytes)
            txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
        Else
            txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
        End If

        Select Case Microsoft.VisualBasic.Left(StrFromBytes(0), 4)
            Case "BND3"
                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}

                BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 12)
                flags = UIntFromBytes(&HC)
                numFiles = UIntFromBytes(&H10)
                namesEndLoc = UIntFromBytes(&H14)

                fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                For i As UInteger = 0 To numFiles - 1
                    Select Case flags
                        Case &H10100
                            currFileSize = UIntFromBytes(&H24 + i * &HC)
                            currFileOffset = UIntFromBytes(&H28 + i * &HC)
                            currFileID = i
                            currFileName = i & "." & Microsoft.VisualBasic.Left(StrFromBytes(currFileOffset), 4)
                            fileList += currFileName & Environment.NewLine
                            currFileName = filepath & filename & ".extract\" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                        Case &HE010100
                            currFileSize = UIntFromBytes(&H24 + i * &H14)
                            currFileOffset = UIntFromBytes(&H28 + i * &H14)
                            currFileID = UIntFromBytes(&H2C + i * &H14)
                            currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                            currFileName = StrFromBytes(currFileNameOffset)
                            fileList += currFileID & "," & currFileName & Environment.NewLine
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                        Case &H2E010100
                            currFileSize = UIntFromBytes(&H24 + i * &H18)
                            currFileOffset = UIntFromBytes(&H28 + i * &H18)
                            currFileID = UIntFromBytes(&H2C + i * &H18)
                            currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                            currFileName = StrFromBytes(currFileNameOffset)
                            fileList += currFileID & "," & currFileName & Environment.NewLine
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                    End Select

                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                        System.IO.Directory.CreateDirectory(currFilePath)
                    End If

                    ReDim currFileBytes(currFileSize - 1)
                    Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                    File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                Next

            Case "DCX"
                Dim newbytes(&H10000) As Byte
                Dim decbytes(&H10000) As Byte
                Dim bytes2(UIntFromBytes(&H1C) - 1) As Byte

                Dim startOffset As UInteger = UIntFromBytes(&H14) + &H20
                Dim numChunks As UInteger = UIntFromBytes(&H68)
                Dim DecSize As UInteger

                fileList = StrFromBytes(&H28) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                For i = 0 To numChunks - 1
                    If i = numChunks - 1 Then
                        DecSize = bytes2.Length - DecSize * i
                    Else
                        DecSize = &H10000
                    End If

                    Array.Copy(bytes, startOffset + UIntFromBytes(&H74 + i * &H10), newbytes, 0, UIntFromBytes(&H78 + i * &H10))
                    decbytes = Decompress(newbytes)
                    Array.Copy(decbytes, 0, bytes2, &H10000 * i, DecSize)
                Next

                currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                If (Not System.IO.Directory.Exists(currFilePath)) Then
                    System.IO.Directory.CreateDirectory(currFilePath)
                End If

                File.WriteAllBytes(currFileName, bytes2)
        End Select

        File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)
        txtInfo.Text += TimeOfDay & " - " & filename & " extracted." & Environment.NewLine
    End Sub
    Private Sub btnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        Dim currFileSize As UInteger = 0
        Dim currFileOffset As UInteger = 0
        Dim currFileNameOffset As UInteger = 0
        Dim currFileName As String = ""
        Dim currFilePath As String = ""
        Dim currFileBytes() As Byte = {}
        Dim currFileID As UInteger = 0
        Dim namesEndLoc As UInteger = 0
        Dim fileList As String() = {""}
        Dim BinderID As String = ""
        Dim flags As UInteger = 0
        Dim numFiles As UInteger = 0
        Dim tmpbytes() As Byte

        Dim padding As UInteger = 0

        filepath = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, "\"))
        filename = Microsoft.VisualBasic.Right(txtBNDfile.Text, txtBNDfile.Text.Length - filepath.Length)

        fileList = File.ReadAllLines(filepath & filename & ".extract\" & "fileList.txt")

        Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
            Case "BND3"
                ReDim bytes(&H1F)
                StrToBytes(fileList(0), 0)

                flags = fileList(1)
                numFiles = fileList.Length - 2

                For i = 2 To fileList.Length - 1
                    namesEndLoc += fileList(i).Length - InStr(fileList(i), ",") + 1
                Next

                Select Case flags
                    Case &H10100
                        namesEndLoc = &H20 + &HC * numFiles
                    Case &HE010100
                        currFileNameOffset = &H20 + &H14 * numFiles
                        namesEndLoc += &H20 + &H14 * numFiles
                    Case &H2E010100
                        currFileNameOffset = &H20 + &H18 * numFiles
                        namesEndLoc += &H20 + &H18 * numFiles
                End Select

                UINTToBytes(flags, &HC)
                UINTToBytes(numFiles, &H10)
                UINTToBytes(namesEndLoc, &H14)

                If namesEndLoc Mod &H10 > 0 Then
                    padding = &H10 - (namesEndLoc Mod &H10)
                Else
                    padding = 0
                End If

                ReDim Preserve bytes(namesEndLoc + padding - 1)

                currFileOffset = namesEndLoc + padding

                For i As UInteger = 0 To numFiles - 1
                    Select Case flags
                        Case &H10100
                            currFileName = fileList(i + 2)
                            currFileName = filepath & filename & ".extract\" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                            tmpbytes = File.ReadAllBytes(currFilePath & currFileName)
                            currFileSize = tmpbytes.Length

                            If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                                padding = &H10 - (currFileSize Mod &H10)
                            Else
                                padding = 0
                            End If

                            UINTToBytes(&H2000000, &H20 + i * &HC)
                            UINTToBytes(currFileSize, &H24 + i * &HC)
                            UINTToBytes(currFileOffset, &H28 + i * &HC)

                            ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                            InsBytes(tmpbytes, currFileOffset)

                            currFileOffset += tmpbytes.Length + padding

                        Case &HE010100
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UINTToBytes(&H2000000, &H20 + i * &H14)
                            UINTToBytes(tmpbytes.Length, &H24 + i * &H14)
                            UINTToBytes(currFileOffset, &H28 + i * &H14)
                            UINTToBytes(currFileID, &H2C + i * &H14)
                            UINTToBytes(currFileNameOffset, &H30 + i * &H14)

                            If tmpbytes.Length Mod &H10 > 0 Then
                                padding = &H10 - (tmpbytes.Length Mod &H10)
                            Else
                                padding = 0
                            End If
                            If i = numFiles - 1 Then padding = 0
                            ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                            InsBytes(tmpbytes, currFileOffset)

                            currFileOffset += tmpbytes.Length
                            If currFileOffset Mod &H10 > 0 Then
                                padding = &H10 - (currFileOffset Mod &H10)
                            Else
                                padding = 0
                            End If
                            currFileOffset += padding

                            StrToBytes(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                            currFileNameOffset += Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))).Length + 1
                        Case &H2E010100
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UINTToBytes(&H2000000, &H20 + i * &H18)
                            UINTToBytes(tmpbytes.Length, &H24 + i * &H18)
                            UINTToBytes(currFileOffset, &H28 + i * &H18)
                            UINTToBytes(currFileID, &H2C + i * &H18)
                            UINTToBytes(currFileNameOffset, &H30 + i * &H18)
                            UINTToBytes(tmpbytes.Length, &H34 + i * &H18)

                            If tmpbytes.Length Mod &H10 > 0 Then
                                padding = &H10 - (tmpbytes.Length Mod &H10)
                            Else
                                padding = 0
                            End If
                            If i = numFiles - 1 Then padding = 0
                            ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                            InsBytes(tmpbytes, currFileOffset)

                            currFileOffset += tmpbytes.Length
                            If currFileOffset Mod &H10 > 0 Then
                                padding = &H10 - (currFileOffset Mod &H10)
                            Else
                                padding = 0
                            End If
                            currFileOffset += padding

                            StrToBytes(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                            currFileNameOffset += Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))).Length + 1
                    End Select
                Next
            Case "EDGE"
                Dim chunkBytes(&H10000) As Byte
                Dim cmpChunkBytes() As Byte
                Dim zipBytes() As Byte = {}

                currFileName = filepath + filename + ".extract\" + fileList(1)
                tmpbytes = File.ReadAllBytes(currFileName)

                currFileSize = tmpbytes.Length

                ReDim bytes(&H83)

                Dim fileRemaining As Integer = tmpbytes.Length
                Dim fileDone As Integer = 0
                Dim fileToDo As Integer = 0
                Dim chunks = 0
                Dim lastchunk = 0

                While fileRemaining > 0
                    chunks += 1

                    If fileRemaining > &H10000 Then
                        fileToDo = &H10000
                    Else
                        fileToDo = fileRemaining
                    End If


                    Array.Copy(tmpbytes, fileDone, chunkBytes, 0, fileToDo)
                    cmpChunkBytes = Compress(chunkBytes)

                    lastchunk = zipBytes.Length

                    If lastchunk Mod &H10 > 0 Then
                        padding = &H10 - (lastchunk Mod &H10)
                    Else
                        padding = 0
                    End If
                    lastchunk += padding

                    ReDim Preserve zipBytes(lastchunk + cmpChunkBytes.Length)
                    Array.Copy(cmpChunkBytes, 0, zipBytes, lastchunk, cmpChunkBytes.Length)


                    fileDone += fileToDo
                    fileRemaining -= fileToDo

                    ReDim Preserve bytes(bytes.Length + &H10)

                    UINTToBytes(lastchunk, &H64 + chunks * &H10)
                    UINTToBytes(cmpChunkBytes.Length, &H68 + chunks * &H10)
                    UINTToBytes(&H1, &H6C + chunks * &H10)

                End While
                ReDim Preserve bytes(bytes.Length + zipBytes.Length)

                StrToBytes("DCX", &H0)
                UINTToBytes(&H10000, &H4)
                UINTToBytes(&H18, &H8)
                UINTToBytes(&H24, &HC)
                UINTToBytes(&H24, &H10)
                UINTToBytes(&H50 + chunks * &H10, &H14)
                StrToBytes("DCS", &H18)
                UINTToBytes(currFileSize, &H1C)
                UINTToBytes(bytes.Length - (&H70 + chunks * &H10), &H20)
                StrToBytes("DCP", &H24)
                StrToBytes("EDGE", &H28)
                UINTToBytes(&H20, &H2C)
                UINTToBytes(&H9000000, &H30)
                UINTToBytes(&H10000, &H34)

                UINTToBytes(&H100100, &H40)
                StrToBytes("DCA", &H44)
                UINTToBytes(chunks * &H10 + &H2C, &H48)
                StrToBytes("EgdT", &H4C)
                UINTToBytes(&H10100, &H50)
                UINTToBytes(&H24, &H54)
                UINTToBytes(&H10, &H58)
                UINTToBytes(&H10000, &H5C)
                UINTToBytes(tmpbytes.Length Mod &H10000, &H60)
                UINTToBytes(&H24 + chunks * &H10, &H64)
                UINTToBytes(chunks, &H68)
                UINTToBytes(&H100000, &H6C)

                Array.Copy(zipBytes, 0, bytes, &H70 + chunks * &H10, zipBytes.Length)

        End Select
        File.WriteAllBytes(filepath & filename, bytes)
        txtInfo.Text += TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine
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
    Public Function Compress(ByVal cmpBytes() As Byte) As Byte()
        Dim ms As New MemoryStream()
        Dim zipStream As Stream = Nothing

        zipStream = New DeflateStream(ms, CompressionMode.Compress, True)
        zipStream.Write(cmpBytes, 0, cmpBytes.Length)
        zipStream.Close()

        ms.Position = 0

        Dim outBytes(ms.Length - 1) As Byte

        ms.Read(outBytes, 0, ms.Length)
        Return outBytes
    End Function

    Private Sub txt_Drop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        sender.Text = file(0)
    End Sub
    Private Sub txt_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub
End Class
