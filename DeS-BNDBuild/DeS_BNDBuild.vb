Imports System.IO
Imports System.IO.Compression

Public Class Des_BNDBuild
    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String
    Public Shared extractPath As String

    Public Shared bigEndian As Boolean = True


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

        If bigEndian Then
            For i = 0 To 3
                tmpUint += Convert.ToUInt32(bytes(loc + i)) * &H100 ^ (3 - i)
            Next
        Else
            For i = 0 To 3
                tmpUint += Convert.ToUInt32(bytes(loc + 3 - i)) * &H100 ^ (3 - i)
            Next
        End If

        Return tmpUint
    End Function

    Private Sub StrToBytes(ByVal str As String, ByVal loc As UInteger)
        Dim BArr() As Byte
        BArr = System.Text.Encoding.ASCII.GetBytes(str)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub
    Private Function Str2Bytes(ByVal str As String) As Byte()
        Dim BArr() As Byte
        BArr = System.Text.Encoding.ASCII.GetBytes(str)
        Return BArr
    End Function
    Private Sub InsBytes(ByVal bytes2() As Byte, ByVal loc As UInteger)
        Array.Copy(bytes2, 0, bytes, loc, bytes2.Length)
    End Sub
    Private Sub UINTToBytes(ByVal val As UInteger, loc As UInteger)
        Dim BArr(3) As Byte

        If bigEndian Then
            For i = 0 To 3
                BArr(i) = Math.Floor(val / (&H100 ^ (3 - i))) Mod &H100
            Next
        Else
            For i = 0 To 3
                BArr(3 - i) = Math.Floor(val / (&H100 ^ (3 - i))) Mod &H100
            Next
        End If
        Array.Copy(BArr, 0, bytes, loc, 4)
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim openDlg As New OpenFileDialog()

        openDlg.Filter = "DeS DCX/BND File|*BND;*MOWB;*DCX;*TPF;*BHD5;*BHD"
        openDlg.Title = "Open your BND file"

        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtBNDfile.Text = openDlg.FileName
        End If


    End Sub


    Private Function HashFileName(filename As String) As UInteger

        REM This code copied from https://github.com/Burton-Radons/Alexandria

        If filename Is Nothing Then
            Return 0
        End If

        Dim hash As UInteger = 0

        For Each ch As Char In filename
            hash = hash * &H25 + Asc(Char.ToLowerInvariant(ch))
        Next

        Return hash
    End Function
    Private Sub WriteBytes(ByRef fs As FileStream, ByVal byt() As Byte)
        For i = 0 To byt.Length - 1
            fs.WriteByte(byt(i))
        Next
    End Sub


    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        bigEndian = True

        Dim DCX As Boolean = False

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
            Case "BHD5"
                If UIntFromBytes(&H4) = 0 Then
                    bigEndian = True
                Else
                    bigEndian = False
                End If

                fileList = "BHD5,"

                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}

                Dim count As UInteger = 0

                Dim idx As Integer
                Dim fileidx() As String = My.Resources.fileidx.Replace(Chr(&HD), "").Split(Chr(&HA))
                Dim hashidx(fileidx.Length - 1) As UInteger

                For i = 0 To fileidx.Length - 1
                    hashidx(i) = HashFileName(fileidx(i))
                Next

                flags = UIntFromBytes(&H4)
                numFiles = UIntFromBytes(&H10)

                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 5)

                If Not File.Exists(filepath & filename & ".bdt.bak") Then
                    File.Copy(filepath & filename & ".bdt", filepath & filename & ".bdt.bak")
                    txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                Else
                    txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                End If

                Dim BDTStream As New IO.FileStream(filepath & filename & ".bdt", IO.FileMode.Open)
                Dim bhdOffSet As UInteger

                BinderID = ""
                For k = 0 To &HF
                    Dim tmpchr As Char
                    tmpchr = Chr(BDTStream.ReadByte)
                    If Not Asc(tmpchr) = 0 Then
                        BinderID = BinderID & tmpchr
                    Else
                        Exit For
                    End If
                Next
                fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                For i As UInteger = 0 To numFiles - 1

                    count = UIntFromBytes(&H18 + i * &H8)
                    bhdOffSet = UIntFromBytes(&H1C + i * 8)



                    For j = 0 To count - 1
                        currFileSize = UIntFromBytes(bhdOffSet + &H4)

                        If bigEndian Then
                            currFileOffset = UIntFromBytes(bhdOffSet + &HC)
                        Else
                            currFileOffset = UIntFromBytes(bhdOffSet + &H8)
                        End If


                        ReDim currFileBytes(currFileSize - 1)

                        BDTStream.Position = currFileOffset

                        For k = 0 To currFileSize - 1
                            currFileBytes(k) = BDTStream.ReadByte
                        Next

                        currFileName = ""


                        If hashidx.Contains(UIntFromBytes(bhdOffSet)) Then
                            idx = Array.IndexOf(hashidx, UIntFromBytes(bhdOffSet))

                            currFileName = fileidx(idx)
                            currFileName = currFileName.Replace("/", "\")
                            fileList += i & "," & currFileName & Environment.NewLine

                            currFileName = filepath & filename & ".bhd5" & ".extract" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                        Else
                            idx = -1
                            currFileName = "NOMATCH-" & Hex(UIntFromBytes(bhdOffSet))
                            fileList += i & "," & currFileName & Environment.NewLine

                            currFileName = filepath & filename & ".bhd5" & ".extract\" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                        End If

                        If (Not System.IO.Directory.Exists(currFilePath)) Then
                            System.IO.Directory.CreateDirectory(currFilePath)
                        End If

                        File.WriteAllBytes(currFileName, currFileBytes)

                        bhdOffSet += &H10
                    Next


                Next
                filename = filename & ".bhd5"
                BDTStream.Close()
                BDTStream.Dispose()

            Case "BHF3"
                fileList = "BHF3,"

                If UIntFromBytes(&H10) = 0 Then
                    bigEndian = True
                Else
                    bigEndian = False
                End If

                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}

                Dim count As UInteger = 0

                flags = UIntFromBytes(&HC)
                numFiles = UIntFromBytes(&H10)

                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 3)

                If Not File.Exists(filepath & filename & "bdt.bak") Then
                    File.Copy(filepath & filename & "bdt", filepath & filename & "bdt.bak")
                    txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                Else
                    txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                End If

                Dim BDTStream As New IO.FileStream(filepath & filename & "bdt", IO.FileMode.Open)
                Dim bhdOffSet As UInteger = &H20

                BinderID = StrFromBytes(&H4)
                fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                For i As UInteger = 0 To numFiles - 1

                    currFileSize = UIntFromBytes(bhdOffSet + &H4)
                    currFileOffset = UIntFromBytes(bhdOffSet + &H8)
                    currFileID = UIntFromBytes(bhdOffSet + &HC)

                    ReDim currFileBytes(currFileSize - 1)

                    BDTStream.Position = currFileOffset

                    For k = 0 To currFileSize - 1
                        currFileBytes(k) = BDTStream.ReadByte
                    Next


                    currFileName = StrFromBytes(UIntFromBytes(bhdOffSet + &H10))
                    fileList += currFileID & "," & currFileName & Environment.NewLine

                    currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                        System.IO.Directory.CreateDirectory(currFilePath)
                    End If

                    File.WriteAllBytes(currFileName, currFileBytes)

                    bhdOffSet += &H18
                Next
                filename = filename & "bhd"
                BDTStream.Close()
                BDTStream.Dispose()
            Case "BND3"
                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}

                BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 12)
                flags = UIntFromBytes(&HC)

                If flags = &H74000000 Or flags = &H54000000 Or flags = &H70000000 Then bigEndian = False

                numFiles = UIntFromBytes(&H10)
                namesEndLoc = UIntFromBytes(&H14)

                fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                If numFiles = 0 Then
                    MsgBox("No files found in archive")
                    Exit Sub
                End If


                For i As UInteger = 0 To numFiles - 1
                    Select Case flags
                        Case &H70000000
                            currFileSize = UIntFromBytes(&H24 + i * &H14)
                            currFileOffset = UIntFromBytes(&H28 + i * &H14)
                            currFileID = UIntFromBytes(&H2C + i * &H14)
                            currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                            currFileName = StrFromBytes(currFileNameOffset)
                            fileList += currFileID & "," & currFileName & Environment.NewLine
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                        Case &H74000000, &H54000000
                            currFileSize = UIntFromBytes(&H24 + i * &H18)
                            currFileOffset = UIntFromBytes(&H28 + i * &H18)
                            currFileID = UIntFromBytes(&H2C + i * &H18)
                            currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                            currFileName = StrFromBytes(currFileNameOffset)
                            fileList += currFileID & "," & currFileName & Environment.NewLine
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(currFileName, currFileName.Length - &H3)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                            currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
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
            Case "TPF"
                'TODO:  Handle m10_9999 (PC) format
                Dim currFileSize As UInteger = 0
                Dim currFileOffset As UInteger = 0
                Dim currFileID As UInteger = 0
                Dim currFileNameOffset As UInteger = 0
                Dim currFileBytes() As Byte = {}
                Dim currFileFlags1 As UInteger = 0
                Dim currFileFlags2 As UInteger = 0

                If UIntFromBytes(&H8) = 0 Then
                    bigEndian = True
                Else
                    bigEndian = False
                End If

                BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                numFiles = UIntFromBytes(&H8)
                flags = UIntFromBytes(&HC)
                currFileNameOffset = UIntFromBytes(&H10)

                fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                For i As UInteger = 0 To numFiles - 1
                    currFileOffset = UIntFromBytes(&H10 + i * &H20)
                    currFileSize = UIntFromBytes(&H14 + i * &H20)
                    currFileFlags1 = UIntFromBytes(&H18 + i * &H20)
                    currFileFlags2 = UIntFromBytes(&H1C + i * &H20)
                    currFileNameOffset = UIntFromBytes(&H28 + i * &H20)
                    currFileName = StrFromBytes(currFileNameOffset)
                    fileList += currFileFlags1 & "," & currFileFlags2 & "," & currFileName & Environment.NewLine
                    currFileName = filepath & filename & ".extract\" & currFileName
                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                        System.IO.Directory.CreateDirectory(currFilePath)
                    End If

                    ReDim currFileBytes(currFileSize - 1)
                    Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                    File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                Next
            Case "DCX"
                Select Case StrFromBytes(&H28)
                    Case "EDGE"
                        DCX = True
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

                        currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                        currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                        If (Not System.IO.Directory.Exists(currFilePath)) Then
                            System.IO.Directory.CreateDirectory(currFilePath)
                        End If

                        File.WriteAllBytes(currFileName, bytes2)
                    Case "DFLT"
                        DCX = True
                        Dim startOffset As UInteger = UIntFromBytes(&H14) + &H22

                        Dim newbytes(UIntFromBytes(&H20) - 1) As Byte
                        Dim decbytes(UIntFromBytes(&H1C)) As Byte

                        fileList = StrFromBytes(&H28) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                        Array.Copy(bytes, startOffset, newbytes, 0, newbytes.Length - 2)



                        decbytes = Decompress(newbytes)


                        currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                        currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                        If (Not System.IO.Directory.Exists(currFilePath)) Then
                            System.IO.Directory.CreateDirectory(currFilePath)
                        End If

                        File.WriteAllBytes(currFileName, decbytes)

                End Select

        End Select

        If Not DCX Then
            File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)
        Else
            File.WriteAllText(filepath & filename & ".info.txt", fileList)
        End If

        txtInfo.Text += TimeOfDay & " - " & filename & " extracted." & Environment.NewLine
    End Sub
    Private Sub btnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        bigEndian = True

        Dim DCX As Boolean = False

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

        DCX = (Microsoft.VisualBasic.Right(filename, 4).ToLower = ".dcx")

        If Not DCX Then
            fileList = File.ReadAllLines(filepath & filename & ".extract\" & "fileList.txt")
        Else
            fileList = File.ReadAllLines(filepath & filename & ".info.txt")
        End If


        Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
            Case "BHD5"
                BinderID = fileList(0).Split(",")(1)
                flags = fileList(1)
                numFiles = fileList.Length - 2
                If flags = 0 Then
                    bigEndian = True
                Else
                    bigEndian = False
                End If

                Dim BDTFilename As String
                BDTFilename = Microsoft.VisualBasic.Left(txtBNDfile.Text, InStrRev(txtBNDfile.Text, ".")) & "bdt"



                File.Delete(BDTFilename)

                Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                BDTStream.Position = 0
                WriteBytes(BDTStream, Str2Bytes(BinderID))
                BDTStream.Position = &H10

                ReDim bytes(&H17)

                Dim bins(fileList.Length - 2) As UInteger
                Dim currBin As UInteger = 0
                Dim totBin As UInteger = 0

                Dim bdtoffset As UInteger = &H10

                For i = 0 To fileList.Length - 3
                    currBin = fileList(i + 2).Split(",")(0)
                    bins(currBin) += 1
                Next

                totBin = Val(fileList(numFiles + 1).Split(",")(0)) + 1

                StrToBytes("BHD5", 0)
                UINTToBytes(flags, &H4)
                UINTToBytes(1, &H8)
                'total file size, &HC
                UINTToBytes(totBin, &H10)
                UINTToBytes(&H18, &H14)


                ReDim Preserve bytes(&H17 + totBin * &H8)
                Dim idxOffset As UInteger
                idxOffset = &H18 + totBin * &H8


                For i As UInteger = 0 To totBin - 1
                    UINTToBytes(bins(i), &H18 + i * &H8)
                    UINTToBytes(idxOffset, &H1C + i * &H8)
                    idxOffset += bins(i) * &H10
                Next

                ReDim Preserve bytes(bytes.Length + numFiles * &H10 - 1)
                idxOffset = &H18 + totBin * &H8

                For i = 0 To numFiles - 1
                    currFileName = fileList(i + 2).Split(",")(1)
                    If currFileName(0) = "\" Then
                        UINTToBytes(HashFileName(currFileName.Replace("\", "/")), idxOffset + i * &H10)
                    Else
                        UINTToBytes(Convert.ToUInt32(currFileName.Split("-")(1), 16), idxOffset + i * &H10)
                        currFileName = "\" & currFileName
                    End If

                    Dim fStream As New IO.FileStream(filepath & filename & ".extract" & currFileName, IO.FileMode.Open)

                    UINTToBytes(fStream.Length, idxOffset + &H4 + i * &H10)
                    If bigEndian Then
                        UINTToBytes(bdtoffset, idxOffset + &HC + i * &H10)
                    Else
                        UINTToBytes(bdtoffset, idxOffset + &H8 + i * &H10)
                    End If


                    For j = 0 To fStream.Length - 1
                        BDTStream.WriteByte(fStream.ReadByte)
                    Next

                    bdtoffset = BDTStream.Position
                    If bdtoffset Mod &H10 > 0 Then
                        padding = &H10 - (bdtoffset Mod &H10)
                    Else
                        padding = 0
                    End If
                    bdtoffset += padding

                    BDTStream.Position = bdtoffset

                    fStream.Close()
                    fStream.Dispose()
                Next

                UINTToBytes(bytes.Length, &HC)

                BDTStream.Close()
                BDTStream.Dispose()

                txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

            Case "BHF3"
                BinderID = fileList(0).Split(",")(1)
                flags = fileList(1)
                numFiles = fileList.Length - 2

                Dim currNameOffset As UInteger = 0

                Dim BDTFilename As String
                BDTFilename = Microsoft.VisualBasic.Left(txtBNDfile.Text, txtBNDfile.Text.Length - 3) & "bdt"

                File.Delete(BDTFilename)

                Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                BDTStream.Position = 0
                WriteBytes(BDTStream, Str2Bytes("BDF3" & BinderID))
                BDTStream.Position = &H10

                ReDim bytes(&H1F)

                Dim bdtoffset As UInteger = &H10

                StrToBytes("BHF3" & BinderID, 0)

                UINTToBytes(flags, &HC)
                UINTToBytes(numFiles, &H10)


                ReDim Preserve bytes(&H1F + numFiles * &H18)
                Dim idxOffset As UInteger
                idxOffset = &H20


                For i = 0 To numFiles - 1
                    currFileID = fileList(i + 2).Split(",")(0)
                    currFileName = fileList(i + 2).Split(",")(1)
                    currNameOffset = bytes.Length

                    Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                    UINTToBytes(&H2000000, idxOffset + i * &H18)
                    UINTToBytes(fStream.Length, idxOffset + &H4 + i * &H18)
                    UINTToBytes(bdtoffset, idxOffset + &H8 + i * &H18)
                    UINTToBytes(currFileID, idxOffset + &HC + i * &H18)
                    UINTToBytes(currNameOffset, idxOffset + &H10 + i * &H18)
                    UINTToBytes(fStream.Length, idxOffset + &H14 + i * &H18)

                    ReDim Preserve bytes(bytes.Length + currFileName.Length)

                    StrToBytes(currFileName, currNameOffset)

                    For j = 0 To fStream.Length - 1
                        BDTStream.WriteByte(fStream.ReadByte)
                    Next

                    bdtoffset = BDTStream.Position
                    If bdtoffset Mod &H10 > 0 Then
                        padding = &H10 - (bdtoffset Mod &H10)
                    Else
                        padding = 0
                    End If
                    bdtoffset += padding

                    BDTStream.Position = bdtoffset

                    fStream.Close()
                    fStream.Dispose()
                Next

                BDTStream.Close()
                BDTStream.Dispose()

                txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine
            Case "BND3"
                ReDim bytes(&H1F)
                StrToBytes(fileList(0), 0)

                flags = fileList(1)
                numFiles = fileList.Length - 2


                For i = 2 To fileList.Length - 1
                    namesEndLoc += fileList(i).Length - InStr(fileList(i), ",") + 1
                Next

                Select Case flags
                    Case &H74000000
                        currFileNameOffset = &H20 + &H18 * numFiles
                        namesEndLoc += &H20 + &H18 * numFiles
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
                If flags = &H74000000 Then bigEndian = False

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
                        Case &H74000000
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UINTToBytes(&H40, &H20 + i * &H18)
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
            Case "TPF"
                'TODO:  Handle m10_9999 (PC) format
                Dim currFileFlags1
                Dim currFileFlags2
                Dim totalFileSize = 0
                ReDim bytes(&HF)
                StrToBytes(fileList(0), 0)

                flags = fileList(1)
                numFiles = fileList.Length - 2

                namesEndLoc = &H10 + numFiles * &H20

                For i = 2 To fileList.Length - 1
                    namesEndLoc += fileList(i).Length - InStrRev(fileList(i), ",") + 1
                Next

                UINTToBytes(numFiles, &H8)
                UINTToBytes(flags, &HC)

                If namesEndLoc Mod &H10 > 0 Then
                    padding = &H10 - (namesEndLoc Mod &H10)
                Else
                    padding = 0
                End If

                ReDim Preserve bytes(namesEndLoc + padding - 1)
                currFileOffset = namesEndLoc + padding

                UINTToBytes(currFileOffset, &H10)

                currFileNameOffset = &H10 + &H20 * numFiles

                For i = 0 To numFiles - 1
                    currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))
                    tmpbytes = File.ReadAllBytes(currFileName)

                    currFileSize = tmpbytes.Length
                    If currFileSize Mod &H10 > 0 Then
                        padding = &H10 - (currFileSize Mod &H10)
                    Else
                        padding = 0
                    End If

                    currFileFlags1 = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)
                    currFileFlags2 = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1), Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1).Length - InStr(Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1), ","))

                    UINTToBytes(currFileOffset, &H10 + i * &H20)
                    UINTToBytes(currFileSize, &H14 + i * &H20)
                    UINTToBytes(currFileFlags1, &H18 + i * &H20)
                    UINTToBytes(currFileFlags2, &H1C + i * &H20)
                    UINTToBytes(currFileNameOffset, &H28 + i * &H20)

                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                    InsBytes(tmpbytes, currFileOffset)

                    currFileOffset += currFileSize + padding
                    totalFileSize += currFileSize

                    StrToBytes(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ","))), currFileNameOffset)
                    currFileNameOffset += Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ","))).Length + 1
                Next

                UINTToBytes(totalFileSize, &H4)


            Case "EDGE"
                Dim chunkBytes(&H10000) As Byte
                Dim cmpChunkBytes() As Byte
                Dim zipBytes() As Byte = {}

                currFileName = filepath + fileList(1)
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
            Case "DFLT"
                Dim cmpBytes() As Byte
                Dim zipBytes() As Byte = {}

                currFileName = filepath + fileList(1)
                tmpbytes = File.ReadAllBytes(currFileName)

                currFileSize = tmpbytes.Length

                ReDim bytes(&H4E)


                cmpBytes = Compress(tmpbytes)

                ReDim Preserve bytes(bytes.Length + cmpBytes.Length)

                StrToBytes("DCX", &H0)
                UINTToBytes(&H10000, &H4)
                UINTToBytes(&H18, &H8)
                UINTToBytes(&H24, &HC)
                UINTToBytes(&H24, &H10)
                UINTToBytes(&H2C, &H14)
                StrToBytes("DCS", &H18)
                UINTToBytes(currFileSize, &H1C)
                UINTToBytes(cmpBytes.Length, &H20)
                StrToBytes("DCP", &H24)
                StrToBytes("DFLT", &H28)
                UINTToBytes(&H20, &H2C)
                UINTToBytes(&H9000000, &H30)

                UINTToBytes(&H10100, &H40)
                StrToBytes("DCA", &H44)
                UINTToBytes(&H8, &H48)
                UINTToBytes(&H78DA0000, &H4C)


                Array.Copy(cmpBytes, 0, bytes, &H4E, cmpBytes.Length)

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
