﻿Imports System.IO
Imports System.IO.Compression
Imports System.Threading



Public Class Des_BNDBuild
    Public Shared bytes() As Byte
    Public Shared filename As String
    Public Shared filepath As String
    Public Shared extractPath As String

    Public Shared bigEndian As Boolean = True

    Private trdExtract As Thread

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
                loc += 2
            Else
                Exit While
            End If
        End While

        Return ShiftJISEncoding.GetString(b.ToArray())
    End Function

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
        Dim bArr(3) As Byte

        Array.Copy(bytes, loc, bArr, 0, 4)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt32(bArr, 0)

        Return tmpUint
    End Function

    Private Function UInt64FromBytes(ByVal loc As UInteger) As ULong
        Dim tmpUint As ULong = 0
        Dim bArr(7) As Byte

        Array.Copy(bytes, loc, bArr, 0, 8)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        tmpUint = BitConverter.ToUInt64(bArr, 0)

        Return tmpUint
    End Function

    Private Sub StrToBytes(ByVal str As String, ByVal loc As UInteger)
        'Insert string directly to main byte array
        Dim BArr() As Byte

        BArr = System.Text.Encoding.ASCII.GetBytes(str)

        Array.Copy(BArr, 0, bytes, loc, BArr.Length)
    End Sub
    Private Function StrToBytes(ByVal str As String) As Byte()
        'Return bytes of string, do not insert
        Return System.Text.Encoding.ASCII.GetBytes(str)
    End Function

    Private Sub InsBytes(ByVal bytes2() As Byte, ByVal loc As UInteger)
        Array.Copy(bytes2, 0, bytes, loc, bytes2.Length)
    End Sub
    Private Sub UIntToBytes(ByVal val As UInteger, loc As UInteger)

        Dim bArr(3) As Byte

        bArr = BitConverter.GetBytes(val)
        If bigEndian Then
            Array.Reverse(bArr)
        End If

        Array.Copy(bArr, 0, bytes, loc, 4)
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
        'Write to stream at present location
        For i = 0 To byt.Length - 1
            fs.WriteByte(byt(i))
        Next
    End Sub


    Private Sub btnExtract_Click(sender As Object, e As EventArgs) Handles btnExtract.Click
        trdExtract = New Thread(AddressOf extract)
        trdExtract.IsBackground = True
        trdExtract.Start()
    End Sub
    Private Sub extract()
        SyncLock workLock
            work = True
        End SyncLock


        Try



            'TODO:  Confirm endian correctness for all DeS/DaS PC/PS3 formats
            'TODO:  Bitch about the massive job that is the above

            'TODO:  Do it anyway.

            'TODO:  In the endian checks, look into why you check if it equals 0
            '       Since that can't matter, since a non-zero value will still
            '       be non-zero in either endian.

            '       Seriously, what the hell were you thinking?


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

            Try
                bytes = File.ReadAllBytes(filepath & filename)
            Catch ex As Exception
                MsgBox(ex.Message, MessageBoxIcon.Error)
                SyncLock workLock
                    work = False
                End SyncLock
                Return
            End Try


            If Not File.Exists(filepath & filename & ".bak") Then
                bytes = File.ReadAllBytes(filepath & filename)
                File.WriteAllBytes(filepath & filename & ".bak", bytes)
                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine
                output(TimeOfDay & " - " & filename & ".bak created." & Environment.NewLine)
            Else
                'txtInfo.Text += TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine
                output(TimeOfDay & " - " & filename & ".bak already exists." & Environment.NewLine)
            End If

            output(TimeOfDay & " - Beginning extraction." & Environment.NewLine)

            Select Case Microsoft.VisualBasic.Left(StrFromBytes(0), 4)
                Case "BHD5"
                    bigEndian = False
                    If Not (UIntFromBytes(&H4) And &HFF) = &HFF Then
                        bigEndian = True
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
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                        'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                    Else
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                        'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
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

                    Dim IsSwitch As Boolean = False
                    If (UIntFromBytes(&H14) = 0 And UIntFromBytes(&H18) = 32) Then
                        IsSwitch = True
                    End If

                    For i As UInteger = 0 To numFiles - 1

                        If (IsSwitch) Then
                            count = UIntFromBytes(&H20 + i * &H10)
                            bhdOffSet = UIntFromBytes(&H28 + i * 16)
                        Else
                            count = UIntFromBytes(&H18 + i * &H8)
                            bhdOffSet = UIntFromBytes(&H1C + i * 8)
                        End If

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
                            output(TimeOfDay & " - Extracted " & currFileName & Environment.NewLine)

                            bhdOffSet += &H10
                        Next


                    Next
                    filename = filename & ".bhd5"
                    BDTStream.Close()
                    BDTStream.Dispose()

                Case "BHF3"
                    fileList = "BHF3,"


                    REM this assumes we'll always have between 1 and 16777215 files 
                    bigEndian = False
                    If UIntFromBytes(&H10) >= &H1000000 Then
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
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine
                        'output(TimeOfDay & " - " & filename & ".bdt.bak created." & Environment.NewLine)
                    Else
                        'txtInfo.Text += TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine
                        'output(TimeOfDay & " - " & filename & ".bdt.bak already exists." & Environment.NewLine)
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


                        currFileName = DecodeFileName(UIntFromBytes(bhdOffSet + &H10))
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
                    'TODO:  DeS, c0300.anibnd, no files found?
                    Dim currFileSize As UInteger = 0
                    Dim currFileOffset As UInteger = 0
                    Dim currFileID As UInteger = 0
                    Dim currFileNameOffset As UInteger = 0
                    Dim currFileBytes() As Byte = {}

                    BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 12)
                    flags = UIntFromBytes(&HC)

                    If flags = &H74000000 Or flags = &H54000000 Or flags = &H70000000 Or flags = &H7C000000 Then bigEndian = False

                    numFiles = UIntFromBytes(&H10)
                    namesEndLoc = UIntFromBytes(&H14)

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
                                currFileSize = UIntFromBytes(&H24 + i * &H14)
                                currFileOffset = UIntFromBytes(&H28 + i * &H14)
                                currFileID = UIntFromBytes(&H2C + i * &H14)
                                currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                                currFileName = DecodeFileName(currFileNameOffset)
                                fileList += currFileID & "," & currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                            Case &H74000000, &H54000000
                                currFileSize = UIntFromBytes(&H24 + i * &H18)
                                currFileOffset = UIntFromBytes(&H28 + i * &H18)
                                currFileID = UIntFromBytes(&H2C + i * &H18)
                                currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                                currFileName = DecodeFileName(currFileNameOffset)
                                fileList += currFileID & "," & currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                            Case &H10100
                                currFileSize = UIntFromBytes(&H24 + i * &HC)
                                currFileOffset = UIntFromBytes(&H28 + i * &HC)
                                currFileID = i
                                currFileName = i & "." & Microsoft.VisualBasic.Left(DecodeFileName(currFileOffset), 4)
                                fileList += currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                            Case &HE010100
                                currFileSize = UIntFromBytes(&H24 + i * &H14)
                                currFileOffset = UIntFromBytes(&H28 + i * &H14)
                                currFileID = UIntFromBytes(&H2C + i * &H14)
                                currFileNameOffset = UIntFromBytes(&H30 + i * &H14)
                                currFileName = DecodeFileName(currFileNameOffset)
                                fileList += currFileID & "," & currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                            Case &H2E010100
                                currFileSize = UIntFromBytes(&H24 + i * &H18)
                                currFileOffset = UIntFromBytes(&H28 + i * &H18)
                                currFileID = UIntFromBytes(&H2C + i * &H18)
                                currFileNameOffset = UIntFromBytes(&H30 + i * &H18)
                                currFileName = DecodeFileName(currFileNameOffset)
                                fileList += currFileID & "," & currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
                                currFilePath = Microsoft.VisualBasic.Left(currFileName, InStrRev(currFileName, "\"))
                                currFileName = Microsoft.VisualBasic.Right(currFileName, currFileName.Length - currFilePath.Length)
                            Case &H7C000000
                                currFileSize = UIntFromBytes(&H24 + i * &H1C)
                                currFileOffset = UIntFromBytes(&H28 + i * &H1C)
                                currFileID = UIntFromBytes(&H30 + i * &H1C)
                                currFileNameOffset = UIntFromBytes(&H34 + i * &H1C)
                                currFileName = DecodeFileName(currFileNameOffset)
                                fileList += currFileID & "," & currFileName & Environment.NewLine
                                currFileName = currFileName.Replace("N:\", "")
                                currFileName = currFileName.Replace("n:\", "")
                                currFileName = filepath & filename & ".extract\" & currFileName
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

                    BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 4) + Microsoft.VisualBasic.Left(StrFromBytes(&H18), 8)

                    numFiles = UIntFromBytes(&HC)
                    flags = UIntFromBytes(&H30)
                    unicode = flags And &HFF
                    type = (flags And &HFF00) >> 8
                    extendedHeader = flags >> 16
                    namesEndLoc = UIntFromBytes(&H38)


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
                            Case &H74

                                currFileSize = UInt64FromBytes(&H48 + i * &H24)
                                currFileOffset = UIntFromBytes(&H58 + i * &H24)
                                currFileID = UIntFromBytes(&H5C + i * &H24)
                                currFileNameOffset = UIntFromBytes(&H60 + i * &H24)
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

                    bigEndian = False
                    If UIntFromBytes(&H8) >= &H1000000 Then
                        bigEndian = True
                    Else
                        bigEndian = False
                    End If

                    flags = UIntFromBytes(&HC)

                    If flags = &H2010200 Or flags = &H2010000 Then
                        ' Demon's Souls

                        BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                        numFiles = UIntFromBytes(&H8)
                        currFileNameOffset = UIntFromBytes(&H10)

                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        For i As UInteger = 0 To numFiles - 1
                            currFileOffset = UIntFromBytes(&H10 + i * &H20)
                            currFileSize = UIntFromBytes(&H14 + i * &H20)
                            currFileFlags1 = UIntFromBytes(&H18 + i * &H20)
                            currFileFlags2 = UIntFromBytes(&H1C + i * &H20)
                            currFileNameOffset = UIntFromBytes(&H28 + i * &H20)
                            currFileName = DecodeFileName(currFileNameOffset)
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
                    ElseIf flags = &H20300 Or flags = &H20304 Then
                        ' Dark Souls

                        BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                        numFiles = UIntFromBytes(&H8)

                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        For i As UInteger = 0 To numFiles - 1
                            currFileOffset = UIntFromBytes(&H10 + i * &H14)
                            currFileSize = UIntFromBytes(&H14 + i * &H14)
                            currFileFlags1 = UIntFromBytes(&H18 + i * &H14)
                            currFileNameOffset = UIntFromBytes(&H1C + i * &H14)
                            currFileFlags2 = UIntFromBytes(&H20 + i * &H14)

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
                    ElseIf flags = &H10300 Then
                        ' Dark Souls III

                        BinderID = Microsoft.VisualBasic.Left(StrFromBytes(&H0), 3)
                        numFiles = UIntFromBytes(&H8)
                        fileList = BinderID & Environment.NewLine & flags & Environment.NewLine

                        For i As UInteger = 0 To numFiles - 1
                            currFileOffset = UIntFromBytes(&H10 + i * &H14)
                            currFileSize = UIntFromBytes(&H14 + i * &H14)
                            currFileFlags1 = UIntFromBytes(&H18 + i * &H14)
                            currFileNameOffset = UIntFromBytes(&H1C + i * &H14)
                            currFileFlags2 = UIntFromBytes(&H20 + i * &H14)

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
                    Else
                        output(TimeOfDay & " - Unknown TPF format" & Environment.NewLine)
                    End If

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

                            fileList = DecodeFileName(&H28) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

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
                            Dim startOffset As UInteger

                            If UIntFromBytes(&H14) = 76 Then
                                startOffset = UIntFromBytes(&H14) + 2
                            Else
                                startOffset = UIntFromBytes(&H14) + &H22
                            End If

                            Dim newbytes(UIntFromBytes(&H20) - 1) As Byte
                            Dim decbytes(UIntFromBytes(&H1C)) As Byte

                            fileList = DecodeFileName(&H28) & Environment.NewLine & Microsoft.VisualBasic.Left(filename, filename.Length - &H4) & Environment.NewLine

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

            output(TimeOfDay & " - " & filename & " extracted." & Environment.NewLine)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            'MessageBox.Show("Stack Trace: " & vbCrLf & ex.StackTrace)
            output(TimeOfDay & " - Unhandled exception - " & ex.Message & Environment.NewLine)
        End Try

        SyncLock workLock
            work = False
        End SyncLock
        'txtInfo.Text += TimeOfDay & " - " & filename & " extracted." & Environment.NewLine
    End Sub
    Private Sub btnRebuild_Click(sender As Object, e As EventArgs) Handles btnRebuild.Click
        'TODO:  Confirm endian before each rebuild.

        'TODO:  List of non-DCXs that don't rebuild byte-perfect
        '   DeS, facegen.tpf
        '   DeS, i7006.tpf
        '   DeS, m07_9990.tpf
        '   DaS, m10_9999.tpf



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

        Try
            If Not DCX Then
                fileList = File.ReadAllLines(filepath & filename & ".extract\" & "fileList.txt")
            Else
                fileList = File.ReadAllLines(filepath & filename & ".info.txt")
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Error)
            SyncLock workLock
                work = False
            End SyncLock
            Return
        End Try


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
                WriteBytes(BDTStream, StrToBytes(BinderID))
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
                UIntToBytes(flags, &H4)
                UIntToBytes(1, &H8)
                'total file size, &HC
                UIntToBytes(totBin, &H10)
                UIntToBytes(&H18, &H14)


                ReDim Preserve bytes(&H17 + totBin * &H8)
                Dim idxOffset As UInteger
                idxOffset = &H18 + totBin * &H8


                For i As UInteger = 0 To totBin - 1
                    UIntToBytes(bins(i), &H18 + i * &H8)
                    UIntToBytes(idxOffset, &H1C + i * &H8)
                    idxOffset += bins(i) * &H10
                Next

                ReDim Preserve bytes(bytes.Length + numFiles * &H10 - 1)
                idxOffset = &H18 + totBin * &H8

                For i = 0 To numFiles - 1
                    currFileName = fileList(i + 2).Split(",")(1)
                    If currFileName(0) = "\" Then
                        UIntToBytes(HashFileName(currFileName.Replace("\", "/")), idxOffset + i * &H10)
                    Else
                        UIntToBytes(Convert.ToUInt32(currFileName.Split("-")(1), 16), idxOffset + i * &H10)
                        currFileName = "\" & currFileName
                    End If

                    Dim fStream As New IO.FileStream(filepath & filename & ".extract" & currFileName, IO.FileMode.Open)

                    UIntToBytes(fStream.Length, idxOffset + &H4 + i * &H10)
                    If bigEndian Then
                        UIntToBytes(bdtoffset, idxOffset + &HC + i * &H10)
                    Else
                        UIntToBytes(bdtoffset, idxOffset + &H8 + i * &H10)
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

                UIntToBytes(bytes.Length, &HC)

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
                WriteBytes(BDTStream, StrToBytes("BDF3" & BinderID))
                BDTStream.Position = &H10

                ReDim bytes(&H1F)

                Dim bdtoffset As UInteger = &H10

                StrToBytes("BHF3" & BinderID, 0)

                UIntToBytes(flags, &HC)
                UIntToBytes(numFiles, &H10)


                ReDim Preserve bytes(&H1F + numFiles * &H18)
                Dim idxOffset As UInteger
                idxOffset = &H20


                For i = 0 To numFiles - 1
                    currFileID = fileList(i + 2).Split(",")(0)
                    currFileName = fileList(i + 2).Split(",")(1)
                    currNameOffset = bytes.Length

                    Dim fStream As New IO.FileStream(filepath & filename & ".extract\" & currFileName, IO.FileMode.Open)

                    UIntToBytes(&H2000000, idxOffset + i * &H18)
                    UIntToBytes(fStream.Length, idxOffset + &H4 + i * &H18)
                    UIntToBytes(bdtoffset, idxOffset + &H8 + i * &H18)
                    UIntToBytes(currFileID, idxOffset + &HC + i * &H18)
                    UIntToBytes(currNameOffset, idxOffset + &H10 + i * &H18)
                    UIntToBytes(fStream.Length, idxOffset + &H14 + i * &H18)

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
                    namesEndLoc += EncodeFileName(fileList(i)).Length - InStr(fileList(i), ",") + 1
                Next

                Select Case flags
                    Case &H74000000, &H54000000
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
                    Case &H7C000000
                        currFileNameOffset = &H20 + &H1C * numFiles
                        namesEndLoc += &H20 + &H1C * numFiles
                End Select

                UIntToBytes(flags, &HC)
                If flags = &H74000000 Or flags = &H54000000 Or flags = &H7C000000 Then bigEndian = False

                UIntToBytes(numFiles, &H10)
                UIntToBytes(namesEndLoc, &H14)

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


                            UIntToBytes(&H40, &H20 + i * &H18)
                            UIntToBytes(tmpbytes.Length, &H24 + i * &H18)
                            UIntToBytes(currFileOffset, &H28 + i * &H18)
                            UIntToBytes(currFileID, &H2C + i * &H18)
                            UIntToBytes(currFileNameOffset, &H30 + i * &H18)
                            UIntToBytes(tmpbytes.Length, &H34 + i * &H18)

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
                        Case &H7C000000
                            currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                            currFileName = currFileName.Replace("N:\", "")
                            currFileName = currFileName.Replace("n:\", "")
                            currFileName = filepath & filename & ".extract\" & currFileName

                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UIntToBytes(&H40, &H20 + i * &H1C)
                            UIntToBytes(tmpbytes.Length, &H24 + i * &H1C)
                            UIntToBytes(currFileOffset, &H28 + i * &H1C)
                            UIntToBytes(currFileID, &H30 + i * &H1C)
                            UIntToBytes(currFileNameOffset, &H34 + i * &H1C)
                            UIntToBytes(tmpbytes.Length, &H38 + i * &H1C)

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

                            UIntToBytes(&H2000000, &H20 + i * &HC)
                            UIntToBytes(currFileSize, &H24 + i * &HC)
                            UIntToBytes(currFileOffset, &H28 + i * &HC)

                            ReDim Preserve bytes(bytes.Length + tmpbytes.Length + padding - 1)

                            InsBytes(tmpbytes, currFileOffset)

                            currFileOffset += tmpbytes.Length + padding

                        Case &HE010100
                            currFileName = filepath & filename & ".extract\" & Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",") + 3))
                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UIntToBytes(&H2000000, &H20 + i * &H14)
                            UIntToBytes(tmpbytes.Length, &H24 + i * &H14)
                            UIntToBytes(currFileOffset, &H28 + i * &H14)
                            UIntToBytes(currFileID, &H2C + i * &H14)
                            UIntToBytes(currFileNameOffset, &H30 + i * &H14)

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


                            UIntToBytes(&H2000000, &H20 + i * &H18)
                            UIntToBytes(tmpbytes.Length, &H24 + i * &H18)
                            UIntToBytes(currFileOffset, &H28 + i * &H18)
                            UIntToBytes(currFileID, &H2C + i * &H18)
                            UIntToBytes(currFileNameOffset, &H30 + i * &H18)
                            UIntToBytes(tmpbytes.Length, &H34 + i * &H18)

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
            Case "BND4"

                'Reversing and hash grouping code by TKGP
                'https://github.com/JKAnderson/SoulsFormats

                Dim type As Byte
                Dim unicode As Byte
                Dim extendedHeader As Byte
                ReDim bytes(&H3F)
                StrToBytes(fileList(0).Substring(0, 4), 0)
                StrToBytes(fileList(0).Substring(4), &H18)

                flags = fileList(1)
                numFiles = fileList.Length - 2

                unicode = flags And &HFF
                type = (flags And &HFF00) >> 8
                extendedHeader = flags >> 16
                For i = 2 To fileList.Length - 1
                    namesEndLoc += EncodeFileNameBND4(fileList(i)).Length - InStr(fileList(i), ",") * 2 + 2
                Next

                Select Case type
                    Case &H74
                        currFileNameOffset = &H40 + &H24 * numFiles
                        namesEndLoc += &H40 + &H24 * numFiles
                End Select

                If type = &H74 Then bigEndian = False

                UIntToBytes(&H40, &H10)
                UIntToBytes(flags, &H30)
                UIntToBytes(numFiles, &HC)
                UIntToBytes(namesEndLoc, &H38)


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


                    ReDim Preserve bytes(namesEndLoc + &H10 + groupCount * 8 + numFiles * 8)

                    UIntToBytes(namesEndLoc, &H38)
                    UIntToBytes(namesEndLoc + &H10 + groupCount * 8, namesEndLoc)
                    namesEndLoc += 4
                    UIntToBytes(0, namesEndLoc)
                    namesEndLoc += 4
                    UIntToBytes(groupCount, namesEndLoc)
                    namesEndLoc += 4
                    UIntToBytes(&H80810, namesEndLoc)
                    namesEndLoc += 4

                    For i As UInteger = 0 To groupCount - 1
                        UIntToBytes(hashGroups(i).length, namesEndLoc)
                        namesEndLoc += 4
                        UIntToBytes(hashGroups(i).idx, namesEndLoc)
                        namesEndLoc += 4
                    Next

                    For i As UInteger = 0 To numFiles - 1
                        UIntToBytes(pathHashes(i).hash, namesEndLoc)
                        namesEndLoc += 4
                        UIntToBytes(pathHashes(i).idx, namesEndLoc)
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

                UIntToBytes(currFileOffset, &H28)

                For i As UInteger = 0 To numFiles - 1
                    Select Case type
                        Case &H74
                            currFileName = Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStr(fileList(i + 2), ",")))
                            currFileName = currFileName.Replace("N:\", "")
                            currFileName = currFileName.Replace("n:\", "")
                            currFileName = filepath & filename & ".extract\" & currFileName

                            tmpbytes = File.ReadAllBytes(currFileName)
                            currFileID = Microsoft.VisualBasic.Left(fileList(i + 2), InStr(fileList(i + 2), ",") - 1)


                            UIntToBytes(&H40, &H40 + i * &H24)
                            UIntToBytes(&HFFFFFFFF, &H44 + i * &H24)
                            UIntToBytes(tmpbytes.Length, &H48 + i * &H24)
                            UIntToBytes(0, &H4C + i * &H24)
                            UIntToBytes(tmpbytes.Length, &H50 + i * &H24)
                            UIntToBytes(0, &H54 + i * &H24)
                            UIntToBytes(currFileOffset, &H58 + i * &H24)
                            UIntToBytes(currFileID, &H5C + i * &H24)
                            UIntToBytes(currFileNameOffset, &H60 + i * &H24)

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
                Dim totalFileSize = 0
                ReDim bytes(&HF)
                StrToBytes(fileList(0), 0)

                flags = fileList(1)

                If flags = &H2010200 Or flags = &H201000 Then
                    ' Demon's Souls
                    'TODO:  Differentiate flag format differences

                    bigEndian = True

                    numFiles = fileList.Length - 2

                    namesEndLoc = &H10 + numFiles * &H20

                    For i = 2 To fileList.Length - 1
                        namesEndLoc += EncodeFileName(fileList(i)).Length - InStrRev(fileList(i), ",") + 1
                    Next

                    UIntToBytes(numFiles, &H8)
                    UIntToBytes(flags, &HC)

                    If namesEndLoc Mod &H10 > 0 Then
                        padding = &H10 - (namesEndLoc Mod &H10)
                    Else
                        padding = 0
                    End If

                    ReDim Preserve bytes(namesEndLoc + padding - 1)
                    currFileOffset = namesEndLoc + padding

                    UIntToBytes(currFileOffset, &H10)

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

                        UIntToBytes(currFileOffset, &H10 + i * &H20)
                        UIntToBytes(currFileSize, &H14 + i * &H20)
                        UIntToBytes(currFileFlags1, &H18 + i * &H20)
                        UIntToBytes(currFileFlags2, &H1C + i * &H20)
                        UIntToBytes(currFileNameOffset, &H28 + i * &H20)

                        ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                        InsBytes(tmpbytes, currFileOffset)

                        currFileOffset += currFileSize + padding
                        totalFileSize += currFileSize

                        EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ","))), currFileNameOffset)
                        currFileNameOffset += EncodeFileName(Microsoft.VisualBasic.Right(fileList(i + 2), fileList(i + 2).Length - (InStrRev(fileList(i + 2), ",")))).Length + 1
                    Next

                    UIntToBytes(totalFileSize, &H4)
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

                    UIntToBytes(numFiles, &H8)
                    UIntToBytes(flags, &HC)

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

                        currFileSize = tmpbytes.Length
                        If currFileSize Mod &H10 > 0 Then
                            padding = &H10 - (currFileSize Mod &H10)
                        Else
                            padding = 0
                        End If

                        Dim words() As String = fileList(i + 2).Split(",")
                        currFileFlags1 = words(0)
                        currFileFlags2 = words(1)

                        UIntToBytes(currFileOffset, &H10 + i * &H14)
                        UIntToBytes(currFileSize, &H14 + i * &H14)
                        UIntToBytes(currFileFlags1, &H18 + i * &H14)
                        UIntToBytes(currFileNameOffset, &H1C + i * &H14)
                        UIntToBytes(currFileFlags2, &H20 + i * &H14)

                        ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                        InsBytes(tmpbytes, currFileOffset)

                        currFileOffset += currFileSize + padding
                        totalFileSize += currFileSize

                        currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                        EncodeFileName(currFileName, currFileNameOffset)
                        currFileNameOffset += EncodeFileName(currFileName).Length + 1
                    Next

                    UIntToBytes(totalFileSize, &H4)
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

                    UIntToBytes(numFiles, &H8)
                    UIntToBytes(flags, &HC)

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

                        UIntToBytes(currFileOffset, &H10 + i * &H14)
                        UIntToBytes(currFileSize, &H14 + i * &H14)
                        UIntToBytes(currFileFlags1, &H18 + i * &H14)
                        UIntToBytes(currFileNameOffset, &H1C + i * &H14)
                        UIntToBytes(currFileFlags2, &H20 + i * &H14)

                        ReDim Preserve bytes(bytes.Length + currFileSize + padding - 1)

                        InsBytes(tmpbytes, currFileOffset)

                        currFileOffset += currFileSize + padding
                        totalFileSize += currFileSize

                        currFileName = currFileName.Substring(0, currFileName.Length - ".dds".Length)
                        EncodeFileNameBND4(currFileName, currFileNameOffset)
                        currFileNameOffset += EncodeFileNameBND4(currFileName).Length + 2
                    Next

                    UIntToBytes(totalFileSize, &H4)
                End If

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

                    UIntToBytes(lastchunk, &H64 + chunks * &H10)
                    UIntToBytes(cmpChunkBytes.Length, &H68 + chunks * &H10)
                    UIntToBytes(&H1, &H6C + chunks * &H10)

                End While
                ReDim Preserve bytes(bytes.Length + zipBytes.Length)

                StrToBytes("DCX", &H0)
                UIntToBytes(&H10000, &H4)
                UIntToBytes(&H18, &H8)
                UIntToBytes(&H24, &HC)
                UIntToBytes(&H24, &H10)
                UIntToBytes(&H50 + chunks * &H10, &H14)
                StrToBytes("DCS", &H18)
                UIntToBytes(currFileSize, &H1C)
                UIntToBytes(bytes.Length - (&H70 + chunks * &H10), &H20)
                StrToBytes("DCP", &H24)
                StrToBytes("EDGE", &H28)
                UIntToBytes(&H20, &H2C)
                UIntToBytes(&H9000000, &H30)
                UIntToBytes(&H10000, &H34)

                UIntToBytes(&H100100, &H40)
                StrToBytes("DCA", &H44)
                UIntToBytes(chunks * &H10 + &H2C, &H48)
                StrToBytes("EgdT", &H4C)
                UIntToBytes(&H10100, &H50)
                UIntToBytes(&H24, &H54)
                UIntToBytes(&H10, &H58)
                UIntToBytes(&H10000, &H5C)
                UIntToBytes(tmpbytes.Length Mod &H10000, &H60)
                UIntToBytes(&H24 + chunks * &H10, &H64)
                UIntToBytes(chunks, &H68)
                UIntToBytes(&H100000, &H6C)

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
                UIntToBytes(&H10000, &H4)
                UIntToBytes(&H18, &H8)
                UIntToBytes(&H24, &HC)
                UIntToBytes(&H24, &H10)
                UIntToBytes(&H2C, &H14)
                StrToBytes("DCS", &H18)
                UIntToBytes(currFileSize, &H1C)
                UIntToBytes(cmpBytes.Length + 4, &H20)
                StrToBytes("DCP", &H24)
                StrToBytes("DFLT", &H28)
                UIntToBytes(&H20, &H2C)
                UIntToBytes(&H9000000, &H30)

                UIntToBytes(&H10100, &H40)
                StrToBytes("DCA", &H44)
                UIntToBytes(&H8, &H48)
                UIntToBytes(&H78DA0000, &H4C)


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
End Class
