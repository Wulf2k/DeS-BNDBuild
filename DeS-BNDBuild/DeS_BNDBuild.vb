Imports System.IO
Imports System.IO.Compression
Imports System.Numerics
Imports System.Threading
Imports System.Security.Cryptography.RijndaelManaged
Imports System.Security.Cryptography



Public Class Des_BNDBuild

    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String
    Public Shared extractPath As String

    Public Shared bigEndian As Boolean = True

    Private trdWorker As Thread

    Dim outputLock As New Object
    Dim workLock As New Object

    Public Shared work As Boolean = False
    Public Shared outputList As New List(Of String)

    Public Structure pathHash
        Dim hash As UInteger
        Dim idx As UInteger
    End Structure

    Public Structure hashGroup
        Dim length As UInteger
        Dim idx As UInteger
    End Structure

    Private WithEvents updateUITimer As New System.Windows.Forms.Timer()

    Dim ShiftJISEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("shift_jis")

    Dim UnicodeEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("UTF-16")

    Public Sub output(txt As String)
        SyncLock outputLock
            outputList.Add(txt)
        End SyncLock
    End Sub

    Private Function EncodeFileName(ByVal filename As String) As Byte()
        Return ShiftJISEncoding.GetBytes(filename)
    End Function

    Private Function EncodeFileNameBND4(ByVal filename As String) As Byte()
        Return UnicodeEncoding.GetBytes(filename)
    End Function

    Private Sub EncodeFileName(ByVal filename As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = ShiftJISEncoding.GetBytes(filename)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub

    Private Sub EncodeFileNameBND4(ByVal filename As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = UnicodeEncoding.GetBytes(filename)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub

    Private Function DecodeFileName(ByVal loc As UInteger) As String
        Dim b As New System.Collections.Generic.List(Of Byte)
        Dim cont As Boolean = True

        While cont
            If bytes(loc) > 0 Then
                b.Add(bytes(loc))
                loc += 1
            Else
                cont = False
            End If
        End While

        Return ShiftJISEncoding.GetString(b.ToArray())
    End Function

    Private Function DecodeFileNameBND4(ByVal loc As UInteger) As String
        Dim b As New System.Collections.Generic.List(Of Byte)
        Dim cont As Boolean = True

        While True
            If bytes(loc) > 0 Then
                b.Add(bytes(loc))
                b.Add(bytes(loc + 1))
                loc += 2
            Else
                Exit While
            End If
        End While

        Return UnicodeEncoding.GetString(b.ToArray())
    End Function

    Private Function RAsciiStr(ByVal loc As UInteger) As String
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

    Private Function RAsciiStrNumBytes(ByVal loc As UInteger, ByRef num As UInteger) As String
        Dim Str As String = ""

        For i As UInteger = 0 To num - 1
            Str = Str + Convert.ToChar(bytes(loc))
            loc += 1
        Next

        Return Str
    End Function

    Private Function RUInt16(ByVal loc As UInteger) As UInt16
        Dim tmpUint As UInteger = 0
        Dim bArr(1) As Byte

        Array.Copy(bytes, loc, bArr, 0, bArr.Length)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt16(bArr, 0)

        Return tmpUint
    End Function
    Private Function RUInt32(ByVal loc As UInteger) As UInteger
        Dim tmpUint As UInteger = 0
        Dim bArr(3) As Byte

        Array.Copy(bytes, loc, bArr, 0, 4)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt32(bArr, 0)

        Return tmpUint
    End Function


    Private Function RUInt64(ByVal loc As UInteger) As ULong
        Dim tmpUint As ULong = 0
        Dim bArr(7) As Byte

        Array.Copy(bytes, loc, bArr, 0, bArr.Length)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt64(bArr, 0)

        Return tmpUint
    End Function

    Private Sub WAsciiStr(ByVal str As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = System.Text.Encoding.ASCII.GetBytes(str)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub
    Private Function StrToBytes(ByVal str As String) As Byte()
        'Return bytes of string, do not insert
        Return System.Text.Encoding.ASCII.GetBytes(str)
    End Function

    Private Sub InsBytes(ByVal bytes2() As Byte, ByVal loc As Long)
        Array.Copy(bytes2, 0, bytes, loc, bytes2.Length)
    End Sub


    Private Sub WUint32(ByVal val As UInteger, loc As UInteger)

        Dim bArr(3) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, 4)
    End Sub
    Private Sub WUInt16(ByVal val As UInt16, loc As UInteger)

        Dim bArr(1) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, bArr.Length)
    End Sub
    Private Sub WUInt64(ByVal val As ULong, loc As UInteger)

        Dim bArr(7) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, bArr.Length)
    End Sub

    Private Sub BtnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Dim openDlg As New OpenFileDialog()

        openDlg.Filter = "DeS DCX/BND File|*BND;*MOWB;*DCX;*TPF;*BHD5;*BHD"
        openDlg.Multiselect = True
        openDlg.Title = "Open your BND files"

        If openDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtBNDfile.Lines = openDlg.FileNames
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
    Private Function DecryptRegulationFile(ByRef regBytes As Byte()) As Byte()
        Dim key As Byte() = System.Text.Encoding.ASCII.GetBytes("ds3#jn/8_7(rsY9pg55GFN7VFL#+3n/)")
        Dim iv As Byte() = regBytes.Take(16).ToArray()

        Dim encBytes As Byte() = regBytes.Skip(16).ToArray()

        Dim ms As New MemoryStream()
        Dim aes As New AesManaged() With {
            .Mode = CipherMode.CBC,
            .Padding = PaddingMode.Zeros
        }
        Dim cs As New CryptoStream(ms, aes.CreateDecryptor(key, iv), CryptoStreamMode.Write)

        cs.Write(encBytes, 0, encBytes.Length)

        cs.Dispose()
        Return ms.ToArray()
    End Function
    Private Function EncryptRegulationFile(ByRef regBytes As Byte()) As Byte()
        Dim key As Byte() = System.Text.Encoding.ASCII.GetBytes("ds3#jn/8_7(rsY9pg55GFN7VFL#+3n/)")

        Dim ms As New MemoryStream()
        Dim aes As New AesManaged() With {
            .Mode = CipherMode.CBC,
            .Padding = PaddingMode.Zeros
        }
        Dim iv(15) As Byte

        Dim cs As New CryptoStream(ms, aes.CreateEncryptor(key, iv), CryptoStreamMode.Write)

        cs.Write(regBytes, 0, regBytes.Length)

        cs.FlushFinalBlock()

        Dim encBytes(iv.Length + ms.Length - 1) As Byte

        Array.Copy(ms.ToArray, 0, encBytes, 16, ms.Length)

        Return encBytes
    End Function
    Private Sub WBytes(ByRef fs As FileStream, ByVal byt() As Byte)
        'Write to stream at present location
        For i = 0 To byt.Length - 1
            fs.WriteByte(byt(i))
        Next
    End Sub


    Private Sub BtnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        trdWorker = New Thread(AddressOf Extract)
        trdWorker.IsBackground = True
        trdWorker.Start()
    End Sub
    Private Sub Extract()
        SyncLock workLock
            work = True
        End SyncLock

        Try

            For Each bndfile In txtBNDfile.Lines
                'TODO:  Confirm endian correctness for all DeS/DaS PC/PS3 formats
                'TODO:  Bitch about the massive job that is the above

                'TODO:  Do it anyway.

                'TODO:  In the endian checks, look into why you check if it equals 0
                '       Since that can't matter, since a non-zero value will still
                '       be non-zero in either endian.

                '       Seriously, what the hell were you thinking?


                bigEndian = True
                Dim DCX As Boolean = False
                Dim IsRegulation As Boolean = False

                Dim currFileName As String = ""
                Dim currFilePath As String = ""
                Dim fileList As String = ""

                Dim BinderID As String = ""
                Dim namesEndLoc As UInteger = 0
                Dim flags As UInteger = 0
                Dim numFiles As UInteger = 0

                filepath = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, "\"))
                filename = Microsoft.VisualBasic.Right(bndfile, bndfile.Length - filepath.Length)
                Try
                    bytes = File.ReadAllBytes(filepath & filename)
                Catch ex As Exception
                    MsgBox(ex.Message, MessageBoxIcon.Error)
                    SyncLock workLock
                        work = False
                    End SyncLock
                    Return
                End Try


                output(TimeOfDay & " - Beginning extraction." & Environment.NewLine)


                If Microsoft.VisualBasic.Right(filename, 3) = "bhd" Then
                    Dim firstBytes As UInteger = RUInt32(&H0)

                    If archiveDict.ContainsKey(firstBytes) Then
                        Dim archiveName = archiveDict(firstBytes)

                        If archiveName = "Data0" Then
                            Try
                                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                                bytes = File.ReadAllBytes(filepath & filename)
                            Catch ex As Exception
                                MsgBox(ex.Message, MessageBoxIcon.Error)
                                SyncLock workLock
                                    work = False
                                End SyncLock
                                Return
                            End Try

                            output(TimeOfDay & " - Beginning decryption of regulation file." & Environment.NewLine)

                            bytes = DecryptRegulationFile(bytes)

                            output(TimeOfDay & " - Finished decryption of regulation file." & Environment.NewLine)

                            IsRegulation = True
                        Else
                            If Not File.Exists(filepath & ".enc.bak") Then
                                File.WriteAllBytes(filepath & filename & ".enc.bak", bytes)
                                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                                output(TimeOfDay & " - " & filename & ".enc.bak created." & Environment.NewLine)
                            Else
                                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                                output(TimeOfDay & " - " & filename & ".enc.bak already exists." & Environment.NewLine)
                            End If
                            output(TimeOfDay & " - Beginning decryption." & Environment.NewLine)

                            Dim decStream As New IO.FileStream(filepath & archiveName & ".bhd", IO.FileMode.Create)
                            Dim idxSheet As ULong = 0
                            Dim diff As UInteger = 0
                            Dim countSheet As UInteger = 256
                            Dim exp As New BigInteger(expDict(archiveName))
                            Dim modulus As New BigInteger(modDict(archiveName))

                            While idxSheet < bytes.Length
                                diff = bytes.Length - idxSheet
                                If diff < 256 Then
                                    countSheet = diff
                                End If
                                Dim tempBlock As Byte() = (bytes.Skip(idxSheet).Take(countSheet)).ToArray()
                                Array.Reverse(tempBlock)
                                ReDim Preserve tempBlock(tempBlock.Length)

                                Dim processBlock As New BigInteger(tempBlock)

                                processBlock = BigInteger.ModPow(processBlock, exp, modulus)

                                Dim processBlockBytes As Byte() = processBlock.ToByteArray().Reverse().ToArray()

                                Dim padding = (countSheet - 1) - processBlockBytes.Length

                                If padding > 0 Then
                                    Dim paddedBlock(countSheet - 2) As Byte
                                    processBlockBytes.CopyTo(paddedBlock, padding)
                                    processBlockBytes = paddedBlock
                                ElseIf padding < 0 Then
                                    processBlockBytes = processBlockBytes.Skip(1).ToArray()
                                End If

                                decStream.Write(processBlockBytes, 0, processBlockBytes.Length)

                                idxSheet += 256
                            End While
                            decStream.Close()
                            output(TimeOfDay & " - Finished decryption." & Environment.NewLine)
                            bytes = File.ReadAllBytes(filepath & filename)
                        End If
                    End If
                End If

                If Microsoft.VisualBasic.Left(RAsciiStr(0), 4) = "DCX" Then
                    Select Case RAsciiStr(&H28)
                        Case "EDGE"
                            DCX = True
                            Dim newbytes(&H10000) As Byte
                            Dim decbytes(&H10000) As Byte
                            Dim bytes2(RUInt32(&H1C) - 1) As Byte

                            Dim startOffset As UInteger = RUInt32(&H14) + &H20
                            Dim numChunks As UInteger = RUInt32(&H68)
                            Dim DecSize As UInteger

                            fileList = DecodeFileName(&H28) & Environment.NewLine & RUInt64(&H10) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                            For i = 0 To numChunks - 1
                                If i = numChunks - 1 Then
                                    DecSize = bytes2.Length - DecSize * i
                                Else
                                    DecSize = &H10000
                                End If

                                Array.Copy(bytes, startOffset + RUInt32(&H74 + i * &H10), newbytes, 0, RUInt32(&H78 + i * &H10))
                                decbytes = Decompress(newbytes)
                                Array.Copy(decbytes, 0, bytes2, &H10000 * i, DecSize)
                            Next

                            currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            File.WriteAllBytes(currFileName, bytes2)

                            File.WriteAllText(filepath & filename & ".info.txt", fileList)
                            output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)

                            bytes = decbytes
                            filepath = currFilePath
                            filename = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - filepath.Length)

                        Case "DFLT"
                            DCX = True
                            Dim startOffset As UInteger

                            If RUInt32(&H14) = 76 Then
                                startOffset = RUInt32(&H14) + 2
                            Else
                                startOffset = RUInt32(&H14) + &H22
                            End If

                            Dim newbytes(RUInt32(&H20) - 1) As Byte
                            Dim decbytes(RUInt32(&H1C)) As Byte

                            fileList = DecodeFileName(&H28) & Environment.NewLine & RUInt64(&H10) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

                            Array.Copy(bytes, startOffset, newbytes, 0, newbytes.Length - 2)

                            decbytes = Decompress(newbytes)

                            If IsRegulation Then
                                currFileName = filepath & filename
                            Else
                                currFileName = filepath & Microsoft.VisualBasic.Left(filename, filename.Length - &H4)
                                File.WriteAllBytes(currFileName, decbytes)
                                output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)
                            End If

                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If


                            File.WriteAllText(filepath & filename & ".info.txt", fileList)


                            bytes = decbytes
                            filepath = currFilePath
                            If IsRegulation Then
                                filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"
                            Else
                                filename = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - filepath.Length)
                            End If

                    End Select
                End If



                Dim OnlyDCX = False


                Select Case Microsoft.VisualBasic.Left(RAsciiStr(0), 4)
                    Case "BHD5"

                        'DS3 BHD5 Reversing by Atvaark
                        'Credits to Atvaark and TKGP for an almost complete list of files
                        'https://github.com/Atvaark/BinderTool
                        'https://github.com/JKAnderson/UXM

                        bigEndian = False
                        If Not (RUInt32(&H4) And &HFF) = &HFF Then
                            bigEndian = True
                        End If

                        fileList = "BHD5,"

                        Dim currFileSize As UInteger = 0
                        Dim currFileOffset As ULong = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim count As UInteger = 0

                        Dim idx As Integer

                        flags = RUInt32(&H4)
                        numFiles = RUInt32(&H10)
                        Dim startOffset As UInteger = 0

                        Dim fileidx() As String
                        Dim IsDS3 As Boolean = False

                        If flags = &H1FF Then
                            fileidx = My.Resources.fileidx_ds3.Replace(Chr(&HD), "").Split(Chr(&HA))
                            IsDS3 = True
                        Else
                            fileidx = My.Resources.fileidx.Replace(Chr(&HD), "").Split(Chr(&HA))
                        End If

                        Dim hashidx(fileidx.Length - 1) As UInteger

                        For i = 0 To fileidx.Length - 1
                            hashidx(i) = HashFileName(fileidx(i))
                        Next

                        If IsDS3 Then
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4)
                        Else
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 5)
                        End If


                        If Not File.Exists(filepath & filename & ".bdt.bak") Then
                            File.Copy(filepath & filename & ".bdt", filepath & filename & ".bdt.bak")
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & ".bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger

                        BinderID = ""
                        If IsDS3 Then
                            BinderID = RAsciiStrNumBytes(&H1C, RUInt32(&H18))
                        Else
                            For k = 0 To &HF
                                Dim tmpchr As Char
                                tmpchr = Chr(BDTStream.ReadByte)
                                If Not Asc(tmpchr) = 0 Then
                                    BinderID = BinderID & tmpchr
                                Else
                                    Exit For
                                End If
                            Next
                        End If

                        Dim IsSwitch As Boolean = False
                        If RUInt32(&H14) = 0 And RUInt32(&H18) = 32 Then
                            IsSwitch = True
                            startOffset = RUInt32(&H18)
                        Else
                            startOffset = RUInt32(&H14)
                        End If

                        fileList = fileList & BinderID & Environment.NewLine & flags & "," & Convert.ToInt32(IsSwitch) & Environment.NewLine

                        For i As UInteger = 0 To numFiles - 1

                            If IsSwitch Then
                                count = RUInt32(startOffset + i * &H10)
                                bhdOffSet = RUInt32(startOffset + 8 + i * 16)
                            Else
                                count = RUInt32(startOffset + i * &H8)
                                bhdOffSet = RUInt32(startOffset + 4 + i * 8)
                            End If

                            For j = 0 To count - 1

                                currFileSize = RUInt32(bhdOffSet + &H4)

                                If bigEndian Then
                                    currFileOffset = RUInt32(bhdOffSet + &HC)
                                ElseIf IsDS3 Then
                                    currFileOffset = RUInt64(bhdOffSet + &H8)
                                Else
                                    currFileOffset = RUInt32(bhdOffSet + &H8)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                BDTStream.Position = currFileOffset
                                BDTStream.Read(currFileBytes, 0, currFileSize)

                                If IsDS3 Then
                                    Dim isEncrypted As Boolean = False
                                    Dim currFileSizeFinal As Long = 0
                                    Dim aesKeyOffset As Long = 0

                                    currFileSizeFinal = RUInt64(bhdOffSet + &H20)
                                    aesKeyOffset = RUInt64(bhdOffSet + &H18)
                                    If aesKeyOffset <> 0 Then
                                        isEncrypted = True
                                    End If

                                    If currFileSizeFinal = 0 Then
                                        Dim header(47) As Byte
                                        Array.Copy(currFileBytes, header, 48)

                                        If isEncrypted Then
                                            Dim aesKey(15) As Byte
                                            Array.Copy(bytes, aesKeyOffset, aesKey, 0, 16)

                                            Dim iv(15) As Byte

                                            Dim ms As New MemoryStream()
                                            Dim aes As New AesManaged() With {
                                                .Mode = CipherMode.ECB,
                                                .Padding = PaddingMode.None
                                            }

                                            Dim cs As New CryptoStream(ms, aes.CreateDecryptor(aesKey, iv), CryptoStreamMode.Write)

                                            cs.Write(header, 0, 48)

                                            Array.Copy(ms.ToArray(), header, 48)
                                            cs.Dispose()
                                        End If

                                        Dim tempBytes(3) As Byte
                                        Array.Copy(header, &H20, tempBytes, 0, 4)
                                        Array.Reverse(tempBytes)
                                        '76 -> DCX header size
                                        currFileSizeFinal = 76 + BitConverter.ToUInt32(tempBytes, 0)

                                    End If

                                    If isEncrypted Then
                                        Dim aesKey(15) As Byte
                                        Dim numRanges As UInteger = RUInt32(aesKeyOffset + &H10)
                                        Dim startOffsets(numRanges - 1) As Long
                                        Dim endOffsets(numRanges - 1) As Long

                                        For k = 0 To numRanges - 1
                                            startOffsets(k) = RUInt64(aesKeyOffset + &H14 + &H10 * k)
                                            endOffsets(k) = RUInt64(aesKeyOffset + &H1C + &H10 * k)
                                        Next

                                        Array.Copy(bytes, aesKeyOffset, aesKey, 0, 16)

                                        Dim iv(15) As Byte

                                        Dim ms As New MemoryStream()
                                        Dim aes As New AesManaged() With {
                                            .Mode = CipherMode.ECB,
                                            .Padding = PaddingMode.None
                                        }

                                        Dim cs As New CryptoStream(ms, aes.CreateDecryptor(aesKey, iv), CryptoStreamMode.Write)

                                        For k = 0 To numRanges - 1
                                            If startOffsets(k) > -1 And endOffsets(k) > -1 Then
                                                cs.Write(currFileBytes, startOffsets(k), endOffsets(k) - startOffsets(k))
                                                Array.Copy(ms.ToArray(), 0, currFileBytes, startOffsets(k), endOffsets(k) - startOffsets(k))
                                                ms.Position = 0
                                            End If
                                        Next
                                        If currFileSize > currFileSizeFinal Then
                                            ReDim Preserve currFileBytes(currFileSizeFinal - 1)
                                        End If
                                        cs.Dispose()
                                    End If


                                End If



                                'BDTStream.Position = currFileOffset

                                'For k = 0 To currFileSize - 1
                                '    currFileBytes(k) = BDTStream.ReadByte
                                'Next


                                currFileName = ""


                                If hashidx.Contains(RUInt32(bhdOffSet)) Then
                                    idx = Array.IndexOf(hashidx, RUInt32(bhdOffSet))

                                    currFileName = fileidx(idx)
                                    currFileName = currFileName.Replace("/", "\")
                                    fileList += i & "," & currFileName & Environment.NewLine

                                    If IsDS3 Then
                                        currFileName = filepath & filename & ".bhd" & ".extract" & currFileName
                                    Else
                                        currFileName = filepath & filename & ".bhd5" & ".extract" & currFileName
                                    End If

                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                Else
                                    idx = -1
                                    currFileName = "NOMATCH-" & Hex(RUInt32(bhdOffSet))
                                    fileList += i & "," & currFileName & Environment.NewLine

                                    If IsDS3 Then
                                        currFileName = filepath & filename & ".bhd" & ".extract\" & currFileName
                                    Else
                                        currFileName = filepath & filename & ".bhd5" & ".extract\" & currFileName
                                    End If

                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                End If

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If


                                File.WriteAllBytes(currFileName, currFileBytes)
                                output(TimeOfDay & " - Extracted " & currFileName & Environment.NewLine)

                                If IsDS3 Then
                                    bhdOffSet += &H28
                                Else
                                    bhdOffSet += &H10
                                End If
                            Next


                        Next

                        If IsDS3 Then
                            filename = filename & ".bhd"
                        Else
                            filename = filename & ".bhd5"
                        End If

                        'BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BHF3"
                        fileList = "BHF3,"


                        REM this assumes we'll always have between 1 and 16777215 files 
                        bigEndian = False
                        If RUInt32(&H10) >= &H1000000 Then
                            bigEndian = True
                        Else
                            bigEndian = False
                        End If

                        Dim currFileSize As UInteger = 0
                        'Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim count As UInteger = 0

                        flags = RUInt32(&HC)
                        numFiles = RUInt32(&H10)

                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 3)

                        If Not File.Exists(filepath & filename & "bdt.bak") Then
                            File.Copy(filepath & filename & "bdt", filepath & filename & "bdt.bak")
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & "bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger = &H20

                        BinderID = RAsciiStr(&H4)
                        fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                        For i As UInteger = 0 To numFiles - 1
                            Select Case flags
                                Case &H7C, &H5C
                                    Dim currFileOffset As ULong = 0
                                    currFileSize = RUInt32(bhdOffSet + &H4)
                                    currFileOffset = RUInt64(bhdOffSet + &H8)
                                    currFileID = RUInt32(bhdOffSet + &H10)

                                    ReDim currFileBytes(currFileSize - 1)

                                    BDTStream.Position = currFileOffset

                                    For k = 0 To currFileSize - 1
                                        currFileBytes(k) = BDTStream.ReadByte
                                    Next


                                    currFileName = DecodeFileName(RUInt32(bhdOffSet + &H14))
                                    fileList += currFileID & "," & currFileName & Environment.NewLine

                                    currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                                        System.IO.Directory.CreateDirectory(currFilePath)
                                    End If

                                    File.WriteAllBytes(currFileName, currFileBytes)

                                    bhdOffSet += &H1C
                                Case Else
                                    Dim currFileOffset As UInteger = 0
                                    currFileSize = RUInt32(bhdOffSet + &H4)
                                    currFileOffset = RUInt32(bhdOffSet + &H8)
                                    currFileID = RUInt32(bhdOffSet + &HC)

                                    ReDim currFileBytes(currFileSize - 1)

                                    BDTStream.Position = currFileOffset

                                    For k = 0 To currFileSize - 1
                                        currFileBytes(k) = BDTStream.ReadByte
                                    Next


                                    currFileName = DecodeFileName(RUInt32(bhdOffSet + &H10))
                                    fileList += currFileID & "," & currFileName & Environment.NewLine

                                    currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                                    If (Not System.IO.Directory.Exists(currFilePath)) Then
                                        System.IO.Directory.CreateDirectory(currFilePath)
                                    End If

                                    File.WriteAllBytes(currFileName, currFileBytes)

                                    bhdOffSet += &H18
                            End Select


                        Next
                        filename = filename & "bhd"
                        BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BHF4"
                        fileList = "BHF4,"


                        bigEndian = False

                        Dim currFileSize As ULong = 0
                        Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        Dim unicode As Boolean = True
                        Dim count As UInteger = 0
                        Dim type As Byte = 0
                        Dim extendedHeader As Byte = 0

                        flags = RUInt32(&H30)
                        unicode = flags And &HFF
                        type = (flags And &HFF00) >> 8
                        extendedHeader = flags >> 16

                        numFiles = RUInt32(&HC)

                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 3)

                        If Not File.Exists(filepath & filename & "bdt.bak") Then
                            File.Copy(filepath & filename & "bdt", filepath & filename & "bdt.bak")
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                            'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
                        End If

                        Dim BDTStream As New IO.FileStream(filepath & filename & "bdt", IO.FileMode.Open)
                        Dim bhdOffSet As UInteger = &H40

                        BinderID = RAsciiStr(&H18)
                        fileList = fileList & BinderID & Environment.NewLine & flags & Environment.NewLine


                        For i As UInteger = 0 To numFiles - 1

                            currFileSize = RUInt64(bhdOffSet + &H8)
                            currFileOffset = RUInt32(bhdOffSet + &H18)
                            currFileID = RUInt32(bhdOffSet + &H1C)

                            ReDim currFileBytes(currFileSize - 1)

                            BDTStream.Position = currFileOffset

                            For k = 0 To currFileSize - 1
                                currFileBytes(k) = BDTStream.ReadByte
                            Next


                            currFileName = DecodeFileNameBND4(RUInt32(bhdOffSet + &H20))
                            fileList += currFileID & "," & currFileName & Environment.NewLine

                            currFileName = filepath & filename & "bhd" & ".extract\" & currFileName
                            currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            File.WriteAllBytes(currFileName, currFileBytes)

                            bhdOffSet += &H24
                        Next
                        filename = filename & "bhd"
                        BDTStream.Close()
                        BDTStream.Dispose()

                    Case "BND3"
                        'TODO:  DeS, c0300.anibnd, no files found?
                        'that archive is busted

                        Dim currFileSize As Long = 0
                        Dim currFileOffset As Long = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}

                        BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 12)
                        flags = RUInt32(&HC)

                        If flags = &H74000000 Or flags = &H54000000 Or flags = &H70000000 Or flags = &H78000000 Or flags = &H7C000000 Or flags = &H5C000000 Then bigEndian = False

                        numFiles = RUInt32(&H10)
                        namesEndLoc = RUInt32(&H14)

                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        If numFiles = 0 Then
                            MsgBox("No files found in archive")
                            SyncLock workLock
                                work = False
                            End SyncLock
                            Exit Sub
                        End If


                        For i As UInteger = 0 To numFiles - 1
                            Select Case flags
                                Case &H70000000
                                    currFileSize = RUInt32(&H24 + i * &H14)
                                    currFileOffset = RUInt32(&H28 + i * &H14)
                                    currFileID = RUInt32(&H2C + i * &H14)
                                    currFileNameOffset = RUInt32(&H30 + i * &H14)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H74000000, &H54000000
                                    currFileSize = RUInt32(&H24 + i * &H18)
                                    currFileOffset = RUInt32(&H28 + i * &H18)
                                    currFileID = RUInt32(&H2C + i * &H18)
                                    currFileNameOffset = RUInt32(&H30 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H10100
                                    currFileSize = RUInt32(&H24 + i * &HC)
                                    currFileOffset = RUInt32(&H28 + i * &HC)
                                    currFileID = i
                                    currFileName = i & "." & Microsoft.VisualBasic.Left(DecodeFileName(currFileOffset), 4)
                                    fileList += currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &HE010100
                                    currFileSize = RUInt32(&H24 + i * &H14)
                                    currFileOffset = RUInt32(&H28 + i * &H14)
                                    currFileID = RUInt32(&H2C + i * &H14)
                                    currFileNameOffset = RUInt32(&H30 + i * &H14)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H2E010100
                                    currFileSize = RUInt32(&H24 + i * &H18)
                                    currFileOffset = RUInt32(&H28 + i * &H18)
                                    currFileID = RUInt32(&H2C + i * &H18)
                                    currFileNameOffset = RUInt32(&H30 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H78000000
                                    currFileSize = RUInt32(&H24 + i * &H18)
                                    currFileOffset = RUInt64(&H28 + i * &H18)
                                    currFileID = RUInt32(&H30 + i * &H18)
                                    currFileNameOffset = RUInt32(&H34 + i * &H18)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case &H7C000000, &H5C000000
                                    currFileSize = RUInt32(&H24 + i * &H1C)
                                    currFileOffset = RUInt64(&H28 + i * &H1C)
                                    currFileID = RUInt32(&H30 + i * &H1C)
                                    currFileNameOffset = RUInt32(&H34 + i * &H1C)
                                    currFileName = DecodeFileName(currFileNameOffset)
                                    fileList += currFileID & "," & currFileName & Environment.NewLine
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case Else
                                    output(TimeOfDay & " - Unknown BND3 type" & Environment.NewLine)
                            End Select

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            ReDim currFileBytes(currFileSize - 1)
                            Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                            File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                        Next
                    Case "BND4"
                        Dim currFileSize As ULong = 0
                        Dim currFileOffset As ULong = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}
                        Dim currFileHash As UInteger = 0
                        Dim currFileHashIdx As UInteger = 0
                        Dim extendedHeader As Byte = 0
                        Dim unicode As Byte = 0
                        Dim type As Byte = 0
                        bigEndian = False

                        BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 4) + Microsoft.VisualBasic.Left(RAsciiStr(&H18), 8)

                        numFiles = RUInt32(&HC)
                        flags = RUInt32(&H30)
                        unicode = flags And &HFF
                        type = (flags And &HFF00) >> 8
                        extendedHeader = flags >> 16
                        namesEndLoc = RUInt32(&H38)


                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        If numFiles = 0 Then
                            MsgBox("No files found in archive")
                            SyncLock workLock
                                work = False
                            End SyncLock
                            Exit Sub
                        End If


                        For i As UInteger = 0 To numFiles - 1
                            Select Case type
                                Case &H74, &H54

                                    currFileSize = RUInt64(&H48 + i * &H24)
                                    currFileOffset = RUInt32(&H58 + i * &H24)
                                    currFileID = RUInt32(&H5C + i * &H24)
                                    currFileNameOffset = RUInt32(&H60 + i * &H24)
                                    currFileName = DecodeFileNameBND4(currFileNameOffset)


                                    If extendedHeader = 4 Then
                                        fileList += currFileID & "," & currFileName & Environment.NewLine
                                    Else
                                        fileList += currFileID & "," & currFileName & Environment.NewLine
                                    End If
                                    currFileName = currFileName.Replace("N:\", "")
                                    currFileName = currFileName.Replace("n:\", "")
                                    currFileName = filepath & filename & ".extract\" & currFileName
                                    currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                    currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                                Case Else
                                    output(TimeOfDay & " - Unknown BND4 type" & Environment.NewLine)
                            End Select

                            If (Not System.IO.Directory.Exists(currFilePath)) Then
                                System.IO.Directory.CreateDirectory(currFilePath)
                            End If

                            ReDim currFileBytes(currFileSize - 1)

                            If currFileSize > 0 Then
                                For j As ULong = 0 To currFileSize - 1
                                    currFileBytes(j) = bytes(currFileOffset + j)
                                Next
                            End If

                            'Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)
                            File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                        Next


                    Case "TPF"
                        Dim currFileSize As UInteger = 0
                        Dim currFileOffset As UInteger = 0
                        Dim currFileID As UInteger = 0
                        Dim currFileNameOffset As UInteger = 0
                        Dim currFileBytes() As Byte = {}
                        Dim currFileFlags1 As UInteger = 0
                        Dim currFileFlags2 As UInteger = 0

                        Dim texWidth As UInt16 = 0
                        Dim texHeight As UInt16 = 0

                        bigEndian = False
                        If RUInt32(&H8) >= &H1000000 Then
                            bigEndian = True
                        Else
                            bigEndian = False
                        End If

                        flags = RUInt32(&HC)

                        If flags = &H2010200 Or flags = &H2010000 Then
                            ' Demon's Souls (headerless DDS)

                            BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 3)
                            numFiles = RUInt32(&H8)
                            currFileNameOffset = RUInt32(&H10)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = RUInt32(&H10 + i * &H20)
                                currFileSize = RUInt32(&H14 + i * &H20)
                                currFileFlags1 = RUInt32(&H18 + i * &H20)
                                texWidth = RUInt16(&H1C + i * &H20)
                                texHeight = RUInt16(&H1E + i * &H20)
                                currFileNameOffset = RUInt32(&H28 + i * &H20)
                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
                                fileList += currFileFlags1 & "," & texWidth & "," & texHeight & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)

                                Dim ddsheader As New DDSHeader
                                ddsheader.Width = texWidth
                                ddsheader.Height = texHeight
                                ddsheader.PitchOrLinearSize = currFileSize

                                If ((currFileFlags1 And &H5000000) > 0) Then
                                    ddsheader.pixelFormat.FourCC = &H35545844
                                End If

                                currFileBytes = ddsheader.ToBytes().Concat(currFileBytes).ToArray()

                                File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                            Next
                        ElseIf flags = &H20300 Or flags = &H20304 Then
                            ' Dark Souls

                            BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 3)
                            numFiles = RUInt32(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = RUInt32(&H10 + i * &H14)
                                currFileSize = RUInt32(&H14 + i * &H14)
                                currFileFlags1 = RUInt32(&H18 + i * &H14)
                                currFileNameOffset = RUInt32(&H1C + i * &H14)
                                currFileFlags2 = RUInt32(&H20 + i * &H14)
                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
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
                        ElseIf flags = &H2030200 Then
                            ' Dark Souls/Demon's Souls (headerless DDS)
                            bigEndian = True
                            BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 3)
                            numFiles = RUInt32(&H8)

                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = RUInt32(&H10 + i * &H20)
                                currFileSize = RUInt32(&H14 + i * &H20)
                                currFileFlags1 = RUInt32(&H18 + i * &H20)
                                texWidth = RUInt16(&H1C + i * &H20)
                                texHeight = RUInt16(&H1E + i * &H20)
                                currFileNameOffset = RUInt32(&H28 + i * &H20)

                                currFileName = DecodeFileName(currFileNameOffset) & ".dds"
                                fileList += currFileFlags1 & "," & texWidth & "," & texHeight & "," & currFileName & Environment.NewLine
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)

                                If (Not System.IO.Directory.Exists(currFilePath)) Then
                                    System.IO.Directory.CreateDirectory(currFilePath)
                                End If

                                ReDim currFileBytes(currFileSize - 1)
                                Array.Copy(bytes, currFileOffset, currFileBytes, 0, currFileSize)

                                Dim ddsheader As New DDSHeader
                                ddsheader.Width = texWidth
                                ddsheader.Height = texHeight
                                ddsheader.PitchOrLinearSize = currFileSize

                                If ((currFileFlags1 And &H5000000) > 0) Then
                                    ddsheader.pixelFormat.FourCC = &H35545844
                                End If

                                currFileBytes = ddsheader.ToBytes().Concat(currFileBytes).ToArray()

                                File.WriteAllBytes(currFilePath & currFileName, currFileBytes)
                            Next


                        ElseIf flags = &H10300 Then
                            ' Dark Souls III

                            BinderID = Microsoft.VisualBasic.Left(RAsciiStr(&H0), 3)
                            numFiles = RUInt32(&H8)
                            fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                            For i As UInteger = 0 To numFiles - 1
                                currFileOffset = RUInt32(&H10 + i * &H14)
                                currFileSize = RUInt32(&H14 + i * &H14)
                                currFileFlags1 = RUInt32(&H18 + i * &H14)
                                currFileNameOffset = RUInt32(&H1C + i * &H14)
                                currFileFlags2 = RUInt32(&H20 + i * &H14)

                                currFileName = DecodeFileNameBND4(currFileNameOffset) & ".dds"
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
                        ElseIf flags = &H10304 Then
                            ' Dark Souls III (headerless DDS)
                            output(TimeOfDay & " - TPF format not implemented" & Environment.NewLine)
                        Else
                            output(TimeOfDay & " - Unknown TPF format" & Environment.NewLine)
                        End If

                    Case Else
                        OnlyDCX = True

                End Select

                If Not OnlyDCX Then
                    File.WriteAllText(filepath & filename & ".extract\filelist.txt", fileList)
                    output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)
                End If


            Next



        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)
            output(TimeOfDay & " - Unhandled exception - " & ex.Message & ex.StackTrace & Environment.NewLine)
        End Try

        SyncLock workLock
            work = False
        End SyncLock
        'txtInfo.Text += TimeOfDay & " - " & filename & " extracted." & Environment.NewLine
    End Sub
    Private Sub BtnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        trdWorker = New Thread(AddressOf Rebuild) With {
            .IsBackground = True
        }
        trdWorker.Start()
    End Sub
    Private Sub Rebuild()
        'TODO:  Confirm endian before each rebuild.

        'TODO:  List of non-DCXs that don't rebuild byte-perfect
        '   DeS, facegen.tpf
        '   DeS, i7006.tpf
        '   DeS, m07_9990.tpf
        '   DaS, m10_9999.tpf
        SyncLock workLock
            work = True
        End SyncLock

        Try

            For Each bndfile In txtBNDfile.Lines
                bigEndian = True

                Dim DCX As Boolean = False
                Dim OnlyDCX = False
                Dim IsRegulation = False

                Dim currFileSize As UInteger = 0
                Dim currFileOffset As Long = 0
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
                Dim dcxBytes() As Byte

                Dim padding As UInteger = 0

                filepath = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, "\"))
                filename = Microsoft.VisualBasic.Right(bndfile, bndfile.Length - filepath.Length)
                DCX = (Microsoft.VisualBasic.Right(filename, 4).ToLower = ".dcx")

                If Microsoft.VisualBasic.Right(filename, 3) = "bhd" Then
                    bytes = File.ReadAllBytes(filepath & filename)
                    Dim firstBytes As UInteger = RUInt32(&H0)
                    If archiveDict.ContainsKey(firstBytes) Then
                        If archiveDict(firstBytes) = "Data0" Then
                            IsRegulation = True
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"
                            MsgBox("Using a modified Data0.bdt (regulation file) online might get you banned. Proceed at your own risk.", MessageBoxIcon.Warning)
                        End If
                    End If
                End If

                If DCX = True Then
                    dcxBytes = File.ReadAllBytes(filepath & filename)
                    filename = filename.Substring(0, filename.Length - 4)
                End If
                Try
                    output(TimeOfDay & " - Processing filelist.txt..." & Environment.NewLine)
                    'If Not DCX Then
                    fileList = File.ReadAllLines(filepath & filename & ".extract\" & "fileList.txt")
                    'Else
                    'fileList = File.ReadAllLines(filepath & filename & ".info.txt")
                    'End If
                Catch ex As DirectoryNotFoundException
                    OnlyDCX = True
                Catch ex As Exception
                    MsgBox(ex.Message, MessageBoxIcon.Error)
                    SyncLock workLock
                        work = False
                    End SyncLock
                    Return
                End Try

                If OnlyDCX = False Then
                    If IsRegulation Then
                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                    End If
                    If Not File.Exists(filepath & filename & ".bak") Then
                        bytes = File.ReadAllBytes(filepath & filename)
                        File.WriteAllBytes(filepath & filename & ".bak", bytes)
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                        output(TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine)
                    Else
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                        output(TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine)
                    End If

                    Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
                        Case "BHD5"

                            BinderID = fileList(0).Split(",")(1)

                            If fileList(1).Split(",").Length < 2 Then
                                MsgBox("filelist.txt incompatible. Please extract once more before rebuilding.", MessageBoxIcon.Error)
                                SyncLock workLock
                                    work = False
                                End SyncLock
                                Return
                            End If
                            output(TimeOfDay & " - Beginning BHD5 rebuild." & Environment.NewLine)

                            Dim IsSwitch As Boolean = fileList(1).Split(",")(1)

                            flags = fileList(1).Split(",")(0)
                            numFiles = fileList.Length - 2
                            If flags = 0 Then
                                bigEndian = True
                            Else
                                bigEndian = False
                            End If

                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, InStrRev(bndfile, ".")) & "bdt"

                            Dim IsDS3 As Boolean = False

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            Dim bdtoffset As ULong = 0

                            Dim bins(fileList.Length - 2) As UInteger
                            Dim currBin As UInteger = 0
                            Dim totBin As UInteger = 0

                            Dim bucketEntryLength = &H10
                            Dim bucketLength = &H8

                            For i = 0 To fileList.Length - 3
                                currBin = fileList(i + 2).Split(",")(0)
                                bins(currBin) += 1
                            Next
                            totBin = Val(fileList(numFiles + 1).Split(",")(0)) + 1

                            Dim idxOffset As UInteger = 0
                            Dim startOffset As UInteger

                            Select Case flags
                                Case &H1FF
                                    IsDS3 = True
                                    startOffset = &H1C + BinderID.Length
                                    bucketEntryLength = &H28

                                Case Else
                                    BDTStream.Position = 0
                                    WBytes(BDTStream, StrToBytes(BinderID))
                                    BDTStream.Position = &H10

                                    bdtoffset = &H10

                                    If IsSwitch Then
                                        startOffset = &H20
                                        bucketLength = &H10
                                    Else
                                        startOffset = &H18
                                    End If

                            End Select

                            ReDim bytes(startOffset - 1)

                            If IsDS3 Then
                                WUint32(BinderID.Length, &H18)
                                WAsciiStr(BinderID, &H1C)
                            End If

                            WAsciiStr("BHD5", 0)
                            WUint32(flags, &H4)
                            WUint32(1, &H8)
                            'total file size, &HC
                            WUint32(totBin, &H10)
                            If IsSwitch Then
                                WUint32(startOffset, &H18)
                            Else
                                WUint32(startOffset, &H14)
                            End If

                            idxOffset = startOffset + totBin * bucketLength

                            ReDim Preserve bytes((startOffset - 1) + totBin * bucketLength)

                            'output(TimeOfDay & " - Generating buckets..." & Environment.NewLine)
                            For i As UInteger = 0 To totBin - 1
                                WUint32(bins(i), startOffset + i * bucketLength)
                                If IsSwitch Then
                                    WUint32(&H1, startOffset + 4 + i * bucketLength)
                                    WUint32(idxOffset, startOffset + 8 + i * bucketLength)
                                Else
                                    WUint32(idxOffset, startOffset + 4 + i * bucketLength)
                                End If
                                idxOffset += bins(i) * bucketEntryLength
                            Next

                            ReDim Preserve bytes(bytes.Length + numFiles * bucketEntryLength - 1)
                            idxOffset = startOffset + totBin * bucketLength

                            For i = 0 To numFiles - 1
                                currFileName = fileList(i + 2).Split(",")(1)
                                If currFileName(0) = "\" Then
                                    WUint32(HashFileName(currFileName.Replace("\", "/")), idxOffset + i * bucketEntryLength)
                                Else
                                    WUint32(Convert.ToUInt32(currFileName.Split("-")(1), 16), idxOffset + i * bucketEntryLength)
                                    currFileName = "\" & currFileName
                                End If

                                Dim fStream As New IO.FileStream(filepath & filename & ".extract" & currFileName, IO.FileMode.Open)

                                WUint32(fStream.Length, idxOffset + &H4 + i * bucketEntryLength)
                                If bigEndian Then
                                    WUint32(bdtoffset, idxOffset + &HC + i * bucketEntryLength)
                                ElseIf IsDS3 Then
                                    WUInt64(bdtoffset, idxOffset + &H8 + i * bucketEntryLength)
                                Else
                                    WUint32(bdtoffset, idxOffset + &H8 + i * bucketEntryLength)
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

                                'fStream.Close()
                                fStream.Dispose()
                                output(TimeOfDay & " - Added " & currFileName & Environment.NewLine)
                            Next

                            WUint32(bytes.Length, &HC)

                            'BDTStream.Close()
                            BDTStream.Dispose()

                        'txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

                        Case "BHF3"
                            BinderID = fileList(0).Split(",")(1)
                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            Dim currNameOffset As UInteger = 0

                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, bndfile.Length - 3) & "bdt"

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            BDTStream.Position = 0
                            WBytes(BDTStream, StrToBytes("BDF3" & BinderID))
                            BDTStream.Position = &H10

                            ReDim bytes(&H1F)

                            Dim bdtoffset As ULong = &H10
                            Dim unk As UInteger = 0

                            WAsciiStr("BHF3" & BinderID, 0)

                            If flags = &H74 Or flags = &H54 Or flags = &H7C Or flags = &H5C Then
                                bigEndian = False
                                unk = &H40
                            Else
                                unk = &H2000000
                            End If

                            WUint32(flags, &HC)
                            WUint32(numFiles, &H10)

                            Dim elemLength As UInteger = &H18

                            If flags = &H7C Or flags = &H5C Then elemLength = &H1C

                            ReDim Preserve bytes(&H1F + numFiles * elemLength)


                            Dim idxOffset As UInteger
                            idxOffset = &H20


                            For i = 0 To numFiles - 1
                                currFileID = fileList(i + 2).Split(",")(0)
                                currFileName = fileList(i + 2).Split(",")(1)
                                currNameOffset = bytes.Length

                                Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                                WUint32(unk, idxOffset + i * elemLength)
                                WUint32(fStream.Length, idxOffset + &H4 + i * elemLength)
                                If flags = &H7C Or flags = &H5C Then
                                    WUInt64(bdtoffset, idxOffset + &H8 + i * elemLength)
                                    WUint32(currFileID, idxOffset + &H10 + i * elemLength)
                                    WUint32(currNameOffset, idxOffset + &H14 + i * elemLength)
                                    WUint32(fStream.Length, idxOffset + &H18 + i * elemLength)
                                Else
                                    WUint32(bdtoffset, idxOffset + &H8 + i * elemLength)
                                    WUint32(currFileID, idxOffset + &HC + i * elemLength)
                                    WUint32(currNameOffset, idxOffset + &H10 + i * elemLength)
                                    WUint32(fStream.Length, idxOffset + &H14 + i * elemLength)
                                End If

                                ReDim Preserve bytes(bytes.Length + currFileName.Length)

                                EncodeFileName(currFileName, currNameOffset)

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

                                'fStream.Close()
                                fStream.Dispose()

                            Next

                            'BDTStream.Close()
                            BDTStream.Dispose()

                            output(TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine)

                            'txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

                        Case "BHF4"

                            Dim type As Byte
                            Dim unicode As Byte
                            Dim extendedHeader As Byte
                            ReDim bytes(&H3F)
                            BinderID = fileList(0).Split(",")(1)
                            WAsciiStr(fileList(0).Substring(0, 4), 0)
                            WAsciiStr(BinderID, &H18)

                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            unicode = flags And &HFF
                            type = (flags And &HFF00) >> 8
                            extendedHeader = flags >> 16
                            For i = 2 To fileList.Length - 1
                                namesEndLoc += EncodeFileNameBND4(fileList(i)).Length - InStr(fileList(i), ",") * 2 + 2
                            Next

                            Select Case type
                                Case &H74, &H54
                                    currFileNameOffset = &H40 + &H24 * numFiles
                                    namesEndLoc += &H40 + &H24 * numFiles
                                    bigEndian = False
                            End Select


                            WUint32(&H10000, &H8)
                            WUint32(&H40, &H10)
                            WUint32(&H24, &H20)
                            WUint32(flags, &H30)
                            WUint32(numFiles, &HC)
                            WUint32(namesEndLoc, &H38)


                            Dim groupCount As UInteger


                            For i As UInteger = numFiles \ 7 To 100000
                                Dim noPrime = False
                                For j As UInteger = 2 To i - 1
                                    If i Mod j = 0 Or i = 2 Then
                                        noPrime = True
                                        Exit For
                                    End If
                                Next
                                If noPrime = False And i > 1 Then
                                    groupCount = i
                                    Exit For
                                End If
                            Next

                            Dim hashLists(groupCount) As List(Of pathHash)


                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i) = New List(Of pathHash)
                            Next

                            Dim hashGroups As New List(Of hashGroup)
                            Dim pathHashes As New List(Of pathHash)

                            If extendedHeader = 4 Then
                                For i As UInteger = 0 To numFiles - 1
                                    Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                    Dim pathHash As pathHash = New pathHash()
                                    If internalFileName(0) <> "\" Then
                                        internalFileName = "\" & internalFileName
                                    End If
                                    Dim hash As UInteger = HashFileName(internalFileName.Replace("\", "/"))

                                    pathHash.hash = hash
                                    pathHash.idx = i
                                    Dim group As UInteger = hash Mod groupCount
                                    hashLists(group).Add(pathHash)
                                Next

                                For i As UInteger = 0 To groupCount - 1
                                    hashLists(i).Sort(Function(x, y) x.hash.CompareTo(y.hash))
                                Next


                                Dim count As UInteger = 0
                                For i As UInteger = 0 To groupCount - 1
                                    Dim index As UInteger = count
                                    For Each pathHash As pathHash In hashLists(i)
                                        pathHashes.Add(pathHash)
                                        count += 1
                                    Next
                                    Dim hashGroup As hashGroup
                                    hashGroup.idx = index
                                    hashGroup.length = count - index
                                    hashGroups.Add(hashGroup)
                                Next

                                Dim extendedPadding As UInteger = namesEndLoc Mod 8
                                If extendedPadding = 0 Then
                                Else
                                    namesEndLoc += 8 - extendedPadding
                                End If

                                ReDim Preserve bytes((namesEndLoc - 1) + &H10 + groupCount * 8 + numFiles * 8)

                                WUint32(namesEndLoc, &H38)
                                WUint32(namesEndLoc + &H10 + groupCount * 8, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(0, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(groupCount, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(&H80810, namesEndLoc)
                                namesEndLoc += 4

                                For i As UInteger = 0 To groupCount - 1
                                    WUint32(hashGroups(i).length, namesEndLoc)
                                    namesEndLoc += 4
                                    WUint32(hashGroups(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next

                                For i As UInteger = 0 To numFiles - 1
                                    WUint32(pathHashes(i).hash, namesEndLoc)
                                    namesEndLoc += 4
                                    WUint32(pathHashes(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next


                            End If


                            Dim BDTFilename As String
                            BDTFilename = Microsoft.VisualBasic.Left(bndfile, bndfile.Length - 3) & "bdt"

                            File.Delete(BDTFilename)

                            Dim BDTStream As New IO.FileStream(BDTFilename, IO.FileMode.CreateNew)

                            BDTStream.Position = 0
                            WBytes(BDTStream, StrToBytes("BDF4"))
                            BDTStream.Position = &HA
                            BDTStream.WriteByte(1)
                            BDTStream.Position = &H10
                            BDTStream.WriteByte(&H30)
                            BDTStream.Position = &H18
                            WBytes(BDTStream, StrToBytes(BinderID))

                            Dim bdtoffset As UInteger = &H30

                            BDTStream.Position = bdtoffset

                            For i = 0 To numFiles - 1

                                Select Case type
                                    Case &H74, &H54
                                        currFileID = fileList(i + 2).Split(",")(0)
                                        currFileName = fileList(i + 2).Split(",")(1)

                                        Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                                        WUint32(&H40, &H40 + i * &H24)
                                        WUint32(&HFFFFFFFF, &H44 + i * &H24)
                                        WUInt64(fStream.Length, &H48 + i * &H24)
                                        WUInt64(fStream.Length, &H50 + i * &H24)
                                        WUint32(bdtoffset, &H58 + i * &H24)
                                        WUint32(currFileID, &H5C + i * &H24)
                                        WUint32(currFileNameOffset, &H60 + i * &H24)

                                        EncodeFileNameBND4(currFileName, currFileNameOffset)
                                        currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2

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





                                End Select

                            Next

                            'BDTStream.Close()
                            BDTStream.Dispose()

                            output(TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine)

                            'txtInfo.Text += TimeOfDay & " - " & BDTFilename & " rebuilt." & Environment.NewLine

                        Case "BND3"
                            ReDim bytes(&H1F)
                            WAsciiStr(fileList(0), 0)

                            flags = fileList(1)
                            numFiles = fileList.Length - 2


                            For i = 2 To fileList.Length - 1
                                namesEndLoc += EncodeFileName(fileList(i)).Length - InStr(fileList(i), ",") + 1
                            Next

                            Select Case flags
                                Case &H74000000, &H78000000, &H54000000
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
                                Case &H7C000000, &H5C000000
                                    currFileNameOffset = &H20 + &H1C * numFiles
                                    namesEndLoc += &H20 + &H1C * numFiles
                            End Select

                            WUint32(flags, &HC)
                            If flags = &H74000000 Or flags = &H78000000 Or flags = &H54000000 Or flags = &H7C000000 Or flags = &H5C000000 Then bigEndian = False

                            WUint32(numFiles, &H10)
                            WUint32(namesEndLoc, &H14)

                            If namesEndLoc Mod &H10 > 0 Then
                                padding = &H10 - (namesEndLoc Mod &H10)
                            Else
                                padding = 0
                            End If

                            ReDim Preserve bytes(namesEndLoc + padding - 1)

                            currFileOffset = namesEndLoc + padding

                            For i As UInteger = 0 To numFiles - 1
                                Select Case flags
                                    Case &H74000000, &H54000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H40, &H20 + i * &H18)
                                        WUint32(tmpbytes.Length, &H24 + i * &H18)
                                        WUint32(currFileOffset, &H28 + i * &H18)
                                        WUint32(currFileID, &H2C + i * &H18)
                                        WUint32(currFileNameOffset, &H30 + i * &H18)
                                        WUint32(tmpbytes.Length, &H34 + i * &H18)

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

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H78000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H40, &H20 + i * &H18)
                                        WUint32(tmpbytes.Length, &H24 + i * &H18)
                                        WUInt64(currFileOffset, &H28 + i * &H18)
                                        WUint32(currFileID, &H30 + i * &H18)
                                        WUint32(currFileNameOffset, &H34 + i * &H18)

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

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H7C000000, &H5C000000
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H40, &H20 + i * &H1C)
                                        WUint32(tmpbytes.Length, &H24 + i * &H1C)
                                        WUInt64(currFileOffset, &H28 + i * &H1C)
                                        WUint32(currFileID, &H30 + i * &H1C)
                                        WUint32(currFileNameOffset, &H34 + i * &H1C)
                                        WUint32(tmpbytes.Length, &H38 + i * &H1C)

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

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
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

                                        WUint32(&H2000000, &H20 + i * &HC)
                                        WUint32(currFileSize, &H24 + i * &HC)
                                        WUint32(currFileOffset, &H28 + i * &HC)

                                        ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                                        InsBytes(tmpbytes, currFileOffset)

                                        currFileOffset += tmpbytes.Length + padding

                                    Case &HE010100
                                        currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H2000000, &H20 + i * &H14)
                                        WUint32(tmpbytes.Length, &H24 + i * &H14)
                                        WUint32(currFileOffset, &H28 + i * &H14)
                                        WUint32(currFileID, &H2C + i * &H14)
                                        WUint32(currFileNameOffset, &H30 + i * &H14)

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

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                    Case &H2E010100
                                        currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H2000000, &H20 + i * &H18)
                                        WUint32(tmpbytes.Length, &H24 + i * &H18)
                                        WUint32(currFileOffset, &H28 + i * &H18)
                                        WUint32(currFileID, &H2C + i * &H18)
                                        WUint32(currFileNameOffset, &H30 + i * &H18)
                                        WUint32(tmpbytes.Length, &H34 + i * &H18)

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

                                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ","))), currFileNameOffset)
                                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))).Length + 1
                                End Select
                            Next

                            output(TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine)
                        Case "BND4"

                            'Reversing and hash grouping code by TKGP
                            'https://github.com/JKAnderson/SoulsFormats

                            Dim type As Byte
                            Dim unicode As Byte
                            Dim extendedHeader As Byte
                            ReDim bytes(&H3F)
                            WAsciiStr(fileList(0).Substring(0, 4), 0)
                            WAsciiStr(fileList(0).Substring(4), &H18)
                            filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bnd"

                            flags = fileList(1)
                            numFiles = fileList.Length - 2

                            unicode = flags And &HFF
                            type = (flags And &HFF00) >> 8
                            extendedHeader = flags >> 16
                            For i = 2 To fileList.Length - 1
                                namesEndLoc += EncodeFileNameBND4(fileList(i)).Length - InStr(fileList(i), ",") * 2 + 2
                            Next

                            Select Case type
                                Case &H74, &H54
                                    currFileNameOffset = &H40 + &H24 * numFiles
                                    namesEndLoc += &H40 + &H24 * numFiles
                                    bigEndian = False
                            End Select


                            WUint32(&H10000, &H8)
                            WUint32(&H40, &H10)
                            WUint32(&H24, &H20)
                            WUint32(flags, &H30)
                            WUint32(numFiles, &HC)
                            WUint32(namesEndLoc, &H38)


                            Dim groupCount As UInteger


                            For i As UInteger = numFiles \ 7 To 100000
                                Dim noPrime = False
                                For j As UInteger = 2 To i - 1
                                    If i Mod j = 0 Or i = 2 Then
                                        noPrime = True
                                        Exit For
                                    End If
                                Next
                                If noPrime = False And i > 1 Then
                                    groupCount = i
                                    Exit For
                                End If
                            Next

                            Dim hashLists(groupCount) As List(Of pathHash)


                            For i As UInteger = 0 To groupCount - 1
                                hashLists(i) = New List(Of pathHash)
                            Next

                            Dim hashGroups As New List(Of hashGroup)
                            Dim pathHashes As New List(Of pathHash)

                            If extendedHeader = 4 Then
                                For i As UInteger = 0 To numFiles - 1
                                    Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                    Dim pathHash As pathHash = New pathHash()
                                    If internalFileName(0) <> "\" Then
                                        internalFileName = "\" & internalFileName
                                    End If
                                    Dim hash As UInteger = HashFileName(internalFileName.Replace("\", "/"))

                                    pathHash.hash = hash
                                    pathHash.idx = i
                                    Dim group As UInteger = hash Mod groupCount
                                    hashLists(group).Add(pathHash)
                                Next

                                For i As UInteger = 0 To groupCount - 1
                                    hashLists(i).Sort(Function(x, y) x.hash.CompareTo(y.hash))
                                Next


                                Dim count As UInteger = 0
                                For i As UInteger = 0 To groupCount - 1
                                    Dim index As UInteger = count
                                    For Each pathHash As pathHash In hashLists(i)
                                        pathHashes.Add(pathHash)
                                        count += 1
                                    Next
                                    Dim hashGroup As hashGroup
                                    hashGroup.idx = index
                                    hashGroup.length = count - index
                                    hashGroups.Add(hashGroup)
                                Next

                                Dim extendedPadding As UInteger = namesEndLoc Mod 8
                                If extendedPadding = 0 Then
                                Else
                                    namesEndLoc += 8 - extendedPadding
                                End If


                                ReDim Preserve bytes((namesEndLoc - 1) + &H10 + groupCount * 8 + numFiles * 8)

                                WUint32(namesEndLoc, &H38)
                                WUint32(namesEndLoc + &H10 + groupCount * 8, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(0, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(groupCount, namesEndLoc)
                                namesEndLoc += 4
                                WUint32(&H80810, namesEndLoc)
                                namesEndLoc += 4

                                For i As UInteger = 0 To groupCount - 1
                                    WUint32(hashGroups(i).length, namesEndLoc)
                                    namesEndLoc += 4
                                    WUint32(hashGroups(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next

                                For i As UInteger = 0 To numFiles - 1
                                    WUint32(pathHashes(i).hash, namesEndLoc)
                                    namesEndLoc += 4
                                    WUint32(pathHashes(i).idx, namesEndLoc)
                                    namesEndLoc += 4
                                Next


                            End If

                            If namesEndLoc Mod &H10 > 0 Then
                                padding = &H10 - (namesEndLoc Mod &H10)
                            Else
                                padding = 0
                            End If

                            ReDim Preserve bytes(namesEndLoc + padding - 1)

                            currFileOffset = namesEndLoc + padding

                            WUint32(namesEndLoc, &H28)

                            For i As UInteger = 0 To numFiles - 1
                                Select Case type
                                    Case &H74, &H54
                                        currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                                        currFileName = currFileName.Replace("N:\", "")
                                        currFileName = currFileName.Replace("n:\", "")
                                        currFileName = filepath & filename & ".extract\" & currFileName

                                        tmpbytes = File.ReadAllBytes(currFileName)
                                        currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                                        WUint32(&H40, &H40 + i * &H24)
                                        WUint32(&HFFFFFFFF, &H44 + i * &H24)
                                        WUint32(tmpbytes.Length, &H48 + i * &H24)
                                        WUint32(0, &H4C + i * &H24)
                                        WUint32(tmpbytes.Length, &H50 + i * &H24)
                                        WUint32(0, &H54 + i * &H24)
                                        WUint32(currFileOffset, &H58 + i * &H24)
                                        WUint32(currFileID, &H5C + i * &H24)
                                        WUint32(currFileNameOffset, &H60 + i * &H24)

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

                                        Dim internalFileName As String = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))

                                        EncodeFileNameBND4(internalFileName, currFileNameOffset)
                                        currFileNameOffset += EncodeFileNameBND4(internalFileName).Length + 2


                                End Select
                            Next


                        Case "TPF"
                            'TODO:  Handle m10_9999 (PC) format
                            Dim currFileFlags1
                            Dim currFileFlags2
                            Dim texWidth
                            Dim texHeight
                            Dim totalFileSize = 0
                            ReDim bytes(&HF)
                            WAsciiStr(fileList(0), 0)

                            flags = fileList(1)

                            If flags = &H2010200 Or flags = &H201000 Then
                                ' Demon's Souls (headerless DDS)
                                'TODO:  Differentiate flag format differences

                                bigEndian = True

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H20

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileName(currFileName).Length + 1
                                Next

                                WUint32(numFiles, &H8)
                                WUint32(flags, &HC)

                                If namesEndLoc Mod &H10 > 0 Then
                                    padding = &H10 - (namesEndLoc Mod &H10)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                WUint32(currFileOffset, &H10)

                                currFileNameOffset = &H10 + &H20 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))
                                    tmpbytes = File.ReadAllBytes(currFileName)


                                    REM Remove DDS Header
                                    Dim tmpbytes2(tmpbytes.Length - &H81) As Byte
                                    Array.Copy(tmpbytes, &H80, tmpbytes2, 0, tmpbytes.Length - &H80)

                                    tmpbytes = tmpbytes2

                                    currFileSize = tmpbytes.Length
                                    If currFileSize Mod &H10 > 0 Then
                                        padding = &H10 - (currFileSize Mod &H10)
                                    Else
                                        padding = 0
                                    End If

                                    currFileFlags1 = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)
                                    texWidth = fileList(i + 2).Split(",")(1)
                                    texHeight = fileList(i + 2).Split(",")(2)

                                    WUint32(currFileOffset, &H10 + i * &H20)
                                    WUint32(currFileSize, &H14 + i * &H20)
                                    WUint32(currFileFlags1, &H18 + i * &H20)
                                    WUInt16(texWidth, &H1C + i * &H20)
                                    WUInt16(texHeight, &H1E + i * &H20)
                                    WUint32(currFileNameOffset, &H28 + i * &H20)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileName(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileName(currFileName).Length + 1
                                Next

                                WUint32(totalFileSize, &H4)
                            ElseIf flags = &H20300 Or flags = &H20304 Then
                                ' Dark Souls
                                'TODO:  Fix this endian check in particular.

                                bigEndian = False

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H14

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileName(currFileName).Length + 1
                                Next

                                WUint32(numFiles, &H8)
                                WUint32(flags, &HC)

                                If namesEndLoc Mod &H10 > 0 Then
                                    padding = &H10 - (namesEndLoc Mod &H10)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                currFileNameOffset = &H10 + &H14 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = fileList(i + 2)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFilePath = filepath & filename & ".extract\"
                                    currFileName = currFileName

                                    tmpbytes = File.ReadAllBytes(currFilePath & currFileName)

                                    REM Remove DDS Header
                                    Dim tmpbytes2(tmpbytes.Length - &H81) As Byte
                                    Array.Copy(tmpbytes, &H80, tmpbytes2, 0, tmpbytes.Length - &H80)

                                    tmpbytes = tmpbytes2

                                    currFileSize = tmpbytes.Length
                                    If currFileSize Mod &H10 > 0 Then
                                        padding = &H10 - (currFileSize Mod &H10)
                                    Else
                                        padding = 0
                                    End If

                                    Dim words() As String = fileList(i + 2).Split(",")
                                    currFileFlags1 = words(0)
                                    currFileFlags2 = words(1)

                                    WUint32(currFileOffset, &H10 + i * &H14)
                                    WUint32(currFileSize, &H14 + i * &H14)
                                    WUint32(currFileFlags1, &H18 + i * &H14)
                                    WUint32(currFileNameOffset, &H1C + i * &H14)
                                    WUint32(currFileFlags2, &H20 + i * &H14)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileName(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileName(currFileName).Length + 1
                                Next

                                WUint32(totalFileSize, &H4)
                            ElseIf flags = &H2030200 Then
                                ' Dark Souls/Demon's Souls (headerless DDS)

                                bigEndian = True

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H20

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileName(currFileName).Length + 1
                                Next

                                WUint32(numFiles, &H8)
                                WUint32(flags, &HC)

                                If namesEndLoc Mod &H100 > 0 Then
                                    padding = &H100 - (namesEndLoc Mod &H100)
                                Else
                                    padding = 0
                                End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                WUint32(currFileOffset, &H10)

                                currFileNameOffset = &H10 + &H20 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))
                                    tmpbytes = File.ReadAllBytes(currFileName)

                                    REM Remove DDS Header
                                    Dim tmpbytes2(tmpbytes.Length - &H81) As Byte
                                    Array.Copy(tmpbytes, &H80, tmpbytes2, 0, tmpbytes.Length - &H80)

                                    tmpbytes = tmpbytes2
                                    currFileSize = tmpbytes.Length
                                    If currFileSize Mod &H20 > 0 Then
                                        padding = &H20 - (currFileSize Mod &H20)
                                    Else
                                        padding = 0
                                    End If

                                    currFileFlags1 = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)
                                    texWidth = fileList(i + 2).Split(",")(1)
                                    texHeight = fileList(i + 2).Split(",")(2)

                                    REM texWidth = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1), Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1).Length - InStr(Microsoft.VisualBasic.Left(fileList(i + 2), InStrRev(fileList(i + 2), ",") - 1), ","))


                                    WUint32(currFileOffset, &H10 + i * &H20)
                                    WUint32(currFileSize, &H14 + i * &H20)
                                    WUint32(currFileFlags1, &H18 + i * &H20)
                                    WUInt16(texWidth, &H1C + i * &H20)
                                    WUInt16(texHeight, &H1E + i * &H20)
                                    WUint32(currFileNameOffset, &H28 + i * &H20)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileName(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileName(currFileName).Length + 1
                                    REM currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))).Length + 1
                                Next

                                WUint32(totalFileSize, &H4)













                            ElseIf flags = &H10300 Then
                                ' Dark Souls III

                                bigEndian = False

                                numFiles = fileList.Length - 2

                                namesEndLoc = &H10 + numFiles * &H14

                                For i = 2 To fileList.Length - 1
                                    currFileName = fileList(i)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    namesEndLoc += EncodeFileNameBND4(currFileName).Length + 2
                                Next

                                WUint32(numFiles, &H8)
                                WUint32(flags, &HC)

                                'If namesEndLoc Mod &H10 > 0 Then
                                'padding = &H10 - (namesEndLoc Mod &H10)
                                'Else
                                'padding = 0
                                'End If

                                ReDim Preserve bytes(namesEndLoc + padding - 1)
                                currFileOffset = namesEndLoc + padding

                                currFileNameOffset = &H10 + &H14 * numFiles

                                For i = 0 To numFiles - 1
                                    currFileName = fileList(i + 2)
                                    currFileName = currFileName.Substring(InStrRev(currFileName, ","))
                                    currFilePath = filepath & filename & ".extract\"
                                    currFileName = currFileName

                                    tmpbytes = File.ReadAllBytes(currFilePath & currFileName)

                                    currFileSize = tmpbytes.Length
                                    'If currFileSize Mod &H10 > 0 Then
                                    'padding = &H10 - (currFileSize Mod &H10)
                                    'Else
                                    'padding = 0
                                    'End If

                                    Dim words() As String = fileList(i + 2).Split(",")
                                    currFileFlags1 = words(0)
                                    currFileFlags2 = words(1)

                                    WUint32(currFileOffset, &H10 + i * &H14)
                                    WUint32(currFileSize, &H14 + i * &H14)
                                    WUint32(currFileFlags1, &H18 + i * &H14)
                                    WUint32(currFileNameOffset, &H1C + i * &H14)
                                    WUint32(currFileFlags2, &H20 + i * &H14)

                                    ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                                    InsBytes(tmpbytes, currFileOffset)

                                    currFileOffset += currFileSize + padding
                                    totalFileSize += currFileSize

                                    currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                                    EncodeFileNameBND4(currFileName, currFileNameOffset)
                                    currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2
                                Next

                                WUint32(totalFileSize, &H4)
                            End If


                    End Select

                    If Not IsRegulation Then
                        File.WriteAllBytes(filepath & filename, bytes)
                        output(TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine)
                    End If


                End If


                If DCX Or IsRegulation Then
                    bigEndian = True
                    If IsRegulation Then
                        filename = Microsoft.VisualBasic.Left(filename, filename.Length - 4) & ".bdt"
                    Else
                        filename = filename & ".dcx"

                        If Not File.Exists(filepath & filename & ".bak") Then
                            File.WriteAllBytes(filepath & filename & ".bak", dcxBytes)
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                            output(TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine)
                        Else
                            'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                            output(TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine)
                        End If
                    End If

                    fileList = File.ReadAllLines(filepath & filename & ".info.txt")

                    Select Case Microsoft.VisualBasic.Left(fileList(0), 4)
                        Case "EDGE"
                            Dim chunkBytes(&H10000) As Byte
                            Dim cmpChunkBytes() As Byte
                            Dim zipBytes() As Byte = {}

                            If fileList.Length > 2 Then
                                currFileName = filepath + fileList(2)
                            Else
                                currFileName = filepath + fileList(1)
                            End If
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

                                WUint32(lastchunk, &H64 + chunks * &H10)
                                WUint32(cmpChunkBytes.Length, &H68 + chunks * &H10)
                                WUint32(&H1, &H6C + chunks * &H10)

                            End While
                            ReDim Preserve bytes(bytes.Length + zipBytes.Length)

                            WAsciiStr("DCX", &H0)
                            WUint32(&H10000, &H4)
                            WUint32(&H18, &H8)
                            WUint32(&H24, &HC)
                            WUint32(&H24, &H10)
                            WUint32(&H50 + chunks * &H10, &H14)
                            WAsciiStr("DCS", &H18)
                            WUint32(currFileSize, &H1C)
                            WUint32(bytes.Length - (&H70 + chunks * &H10), &H20)
                            WAsciiStr("DCP", &H24)
                            WAsciiStr("EDGE", &H28)
                            WUint32(&H20, &H2C)
                            WUint32(&H9000000, &H30)
                            WUint32(&H10000, &H34)

                            WUint32(&H100100, &H40)
                            WAsciiStr("DCA", &H44)
                            WUint32(chunks * &H10 + &H2C, &H48)
                            WAsciiStr("EgdT", &H4C)
                            WUint32(&H10100, &H50)
                            WUint32(&H24, &H54)
                            WUint32(&H10, &H58)
                            WUint32(&H10000, &H5C)
                            WUint32(tmpbytes.Length Mod &H10000, &H60)
                            WUint32(&H24 + chunks * &H10, &H64)
                            WUint32(chunks, &H68)
                            WUint32(&H100000, &H6C)

                            Array.Copy(zipBytes, 0, bytes, &H70 + chunks * &H10, zipBytes.Length)
                        Case "DFLT"
                            Dim cmpBytes() As Byte
                            Dim zipBytes() As Byte = {}

                            If fileList.Length > 2 Then
                                currFileName = filepath + fileList(2)
                            Else
                                currFileName = filepath + fileList(1)
                            End If
                            If IsRegulation Then
                                tmpbytes = bytes
                            Else
                                tmpbytes = File.ReadAllBytes(currFileName)
                            End If

                            currFileSize = tmpbytes.Length

                            ReDim bytes(&H4C)


                            cmpBytes = Compress(tmpbytes)

                            ReDim Preserve bytes(bytes.Length + cmpBytes.Length)

                            WAsciiStr("DCX", &H0)
                            WUint32(&H10000, &H4)
                            WUint32(&H18, &H8)
                            WUint32(&H24, &HC)
                            'UIntToBytes(&H24, &H10)
                            'UIntToBytes(&H2C, &H14)
                            If fileList.Length > 2 Then
                                WUInt64(fileList(1), &H10)
                            Else
                                WUint32(&H24, &H10)
                                WUint32(&H2C, &H14)
                            End If
                            WAsciiStr("DCS", &H18)
                            WUint32(currFileSize, &H1C)
                            WUint32(cmpBytes.Length + 2, &H20)
                            WAsciiStr("DCP", &H24)
                            WAsciiStr("DFLT", &H28)
                            WUint32(&H20, &H2C)
                            WUint32(&H9000000, &H30)

                            WUint32(&H10100, &H40)
                            WAsciiStr("DCA", &H44)
                            WUint32(&H8, &H48)
                            WUint32(&H78DA0000, &H4C)


                            Array.Copy(cmpBytes, 0, bytes, &H4E, cmpBytes.Length)
                    End Select

                    If IsRegulation Then
                        output(TimeOfDay & " - Beginning encryption of regulation file." & Environment.NewLine)
                        bytes = EncryptRegulationFile(bytes)
                        output(TimeOfDay & " - Finished encryption of regulation file." & Environment.NewLine)
                    End If

                    File.WriteAllBytes(filepath & filename, bytes)

                    output(TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine)

                    'txtInfo.Text += TimeOfDay & " - " & filename & " rebuilt." & Environment.NewLine
                End If


            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            output(TimeOfDay & " - Unhandled exception - " & ex.Message & ex.StackTrace & Environment.NewLine)
        End Try

        SyncLock workLock
            work = False
        End SyncLock

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

    Private Function Adler32(ByRef input) As UInteger
        Dim a As UInteger = 1
        Dim b As UInteger = 0

        For Each elem As Byte In input
            a = (a + elem) Mod 65521
            b = (b + a) Mod 65521
        Next

        Return (b << 16) Or a
    End Function

    Public Function Compress(ByVal cmpBytes() As Byte) As Byte()
        Dim ms As New MemoryStream()
        Dim zipStream As Stream = Nothing

        zipStream = New DeflateStream(ms, CompressionMode.Compress, True)
        zipStream.Write(cmpBytes, 0, cmpBytes.Length)
        zipStream.Close()

        ms.Position = 0

        Dim outBytes(ms.Length + 3) As Byte

        Dim adlerBytes As Byte() = BitConverter.GetBytes(Adler32(cmpBytes))
        Array.Reverse(adlerBytes)
        Array.Copy(adlerBytes, 0, outBytes, ms.Length, 4)

        ms.Read(outBytes, 0, ms.Length)


        Return outBytes
    End Function

    Private Sub txt_Drop(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragDrop
        Dim file() As String = e.Data.GetData(DataFormats.FileDrop)
        'sender.Text = file(0)
        sender.Lines = file
    End Sub
    Private Sub txt_DragEnter(sender As Object, e As System.Windows.Forms.DragEventArgs) Handles txtBNDfile.DragEnter
        e.Effect = DragDropEffects.Copy
    End Sub

    Private Sub Des_BNDBuild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        updateUITimer.Interval = 200
        updateUITimer.Start()
    End Sub

    Private Sub updateUI() Handles updateUITimer.Tick
        SyncLock workLock
            If work Then
                btnExtract.Enabled = False
                btnRebuild.Enabled = False
            Else
                btnExtract.Enabled = True
                btnRebuild.Enabled = True
            End If
        End SyncLock

        SyncLock outputLock
            While outputList.Count > 0
                txtInfo.AppendText(outputList(0))
                outputList.RemoveAt(0)
            End While
        End SyncLock


        If txtInfo.Lines.Count > 10000 Then
            Dim newList As List(Of String) = txtInfo.Lines.ToList
            While newList.Count > 1000
                newList.RemoveAt(0)
            End While
            txtInfo.Lines = newList.ToArray
        End If
    End Sub

    Public Class DDS_PIXELFORMAT
        Public Property size As UInt32 = &H20
        Public Property Flags As UInt32 = &H4
        REM 0x1 Alphapixels
        REM 0x2 Old alpha info
        REM 0x4 Texture contains compressed RGB data; dwFourCC has valid data
        REM 0x40 Texture contains uncompressed RGB data; dwRGBBitCount and RGB masks contain valid data
        REM 0x200 older, YUV uncompressed data
        REM 0x20000 older, dwRGBBitCount contains luminance channel bit count
        REM Public Property FourCC As UInt32 = &H44585433
        Public Property FourCC As UInt32 = &H31545844
        Public Property RGBBitCount As UInt32 = 0
        Public Property RBitMask As UInt32 = 0
        Public Property GBitMask As UInt32 = 0
        Public Property BBitMask As UInt32 = 0
        Public Property ABitMask As UInt32 = 0
        Public Function ToBytes() As Byte()
            Dim bytes(0) As Byte
            bytes = BitConverter.GetBytes(size)
            bytes = bytes.Concat(BitConverter.GetBytes(Flags)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(FourCC)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(RGBBitCount)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(RBitMask)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(GBitMask)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(BBitMask)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(ABitMask)).ToArray()
            Return bytes
        End Function

    End Class
    Public Class DDSHeader
        REM Public Property Magic As UInt32 = &H44445320
        Public Property Magic As UInt32 = &H20534444
        Public Property size As UInt32 = &H7C
        Public Property flags As UInt32 = &H81007
        Public Property Height As UInt32 = 0
        Public Property Width As UInt32 = 0
        Public Property PitchOrLinearSize As UInt32 = 0
        Public Property Depth As UInt32 = 0
        Public Property MipMapCount As UInt32 = 0
        REM DWORD 0x11 reserved
        Public Property pixelFormat As New DDS_PIXELFORMAT
        Public Property Caps As UInt32 = &H1000
        Public Property Caps2 As UInt32 = 0
        Public Property Caps3 As UInt32 = 0
        Public Property Caps4 As UInt32 = 0
        REM DWORD reserved

        Public Function ToBytes() As Byte()
            Dim bytes(0) As Byte
            Dim reserved(43) As Byte
            bytes = BitConverter.GetBytes(Magic)
            bytes = bytes.Concat(BitConverter.GetBytes(size)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(flags)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Height)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Width)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(PitchOrLinearSize)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Depth)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(MipMapCount)).ToArray()
            bytes = bytes.Concat(reserved).ToArray()
            bytes = bytes.Concat(pixelFormat.ToBytes()).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Caps)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Caps2)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Caps3)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(Caps4)).ToArray()
            bytes = bytes.Concat(BitConverter.GetBytes(0)).ToArray()
            Return bytes
        End Function
    End Class
End Class
