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

        openDlg.Filter = "DeS DCX/BND File|*BND;*DCX"
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

        Select Case Microsoft.VisualBasic.Left(StrFromBytes(0), 3)
            Case "BND"
                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}

                BinderID = StrFromBytes(&H0)
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

        Select Case Microsoft.VisualBasic.Left(fileList(0), 3)
            Case "BND"
                ReDim bytes(&H1F)
                StrToBytes(fileList(0), 0)

                flags = fileList(1)
                numFiles = fileList.Length - 2

                For i = 2 To fileList.Length - 1
                    namesEndLoc += fileList(i).Length - InStr(fileList(i), ",") + 1
                Next

                namesEndLoc += &H20 + &H18 * numFiles

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
                currFileNameOffset = &H20 + &H18 * numFiles

                For i As UInteger = 0 To numFiles - 1
                    Select Case flags
                        Case &H10100
                            'currFileName = fileList(i + 1)
                            'currFileName = filepath & filename & ".extract\" & currFileName
                            'currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            'currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                            'currFileBytes = File.ReadAllBytes(currFilePath & currFileName)
                            'currFileSize = currFileBytes.Length

                            'If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                            '    padding = &H10 - (currFileSize Mod &H10)
                            'Else
                            '    padding = 0
                            'End If

                            'currFileOffset = bytes.Length
                            'UINTToBytes(&H24 + i * &HC, currFileSize)
                            'UINTToBytes(&H28 + i * &HC, currFileOffset)

                        Case &HE010100
                            'currFileNameOffset = UIntFromBytes(&H30 + i * &H14)

                            'currFileName = StrFromBytes(currFileNameOffset)
                            'currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            'currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            'currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                            'currFileBytes = File.ReadAllBytes(currFilePath & currFileName)
                            'currFileSize = currFileBytes.Length

                            'If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                            '    padding = &H10 - (currFileSize Mod &H10)
                            'Else
                            '    padding = 0
                            'End If

                            'currFileOffset = bytes.Length
                            'UINTToBytes(&H24 + i * &H14, currFileSize)
                            'UINTToBytes(&H28 + i * &H14, currFileOffset)
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





                            'currFileNameOffset = UIntFromBytes(&H30 + i * &H18)

                            'currFileName = StrFromBytes(currFileNameOffset)
                            'currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            'currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            'currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                            'currFileBytes = File.ReadAllBytes(currFilePath & currFileName)
                            'currFileSize = currFileBytes.Length

                            'If currFileSize Mod &H10 > 0 And i < numFiles - 1 Then
                            '    padding = &H10 - (currFileSize Mod &H10)
                            'Else
                            '    padding = 0
                            'End If

                            'currFileOffset = bytes.Length
                            'UINTToBytes(&H24 + i * &H18, currFileSize)
                            'UINTToBytes(&H28 + i * &H18, currFileOffset)
                            'UINTToBytes(&H34 + i * &H18, currFileSize)
                    End Select

                    'ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)
                    'Array.Copy(currFileBytes, 0, bytes, currFileOffset, currFileSize)
                Next
            Case "DCX"

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

    Private Sub txt_Drop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        sender.Text = file(0)
    End Sub
    Private Sub txt_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub
End Class
