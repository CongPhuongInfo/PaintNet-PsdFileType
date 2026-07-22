'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/ZipPredict32Image.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: the original reinterprets a byte buffer directly as an
' Int32* to read/write native-endian 32-bit words. VB.NET has no
' unsafe/pointer support, so each 32-bit word is read/written through
' BitConverter at the matching byte offset instead, which reproduces the
' same native-endian in-place reinterpretation without pointers. The
' byte-level packing/delta logic (which explicitly shifts individual bytes)
' is unchanged, since it never relied on pointers for anything but iteration.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing

Namespace PhotoshopFile.Compression

    Public Class ZipPredict32Image
        Inherits ImageData

        Private zipImage As ZipImage

        ' Prediction will pack the data into a temporary buffer, so the
        ' original data will remain unchanged.
        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return False
            End Get
        End Property

        Public Sub New(zipData As Byte(), size As Size)
            MyBase.New(size, 32)
            zipImage = New ZipImage(zipData, size, 32)
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            If buffer.Length = 0 Then
                Return
            End If

            Dim predictedData(buffer.Length - 1) As Byte
            zipImage.Read(predictedData)

            Unpredict(predictedData, buffer)
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            Return zipImage.ReadCompressed()
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            If array.Length = 0 Then
                Return
            End If

            Dim predictedData(array.Length - 1) As Byte
            Predict(array, predictedData)

            zipImage.WriteInternal(predictedData)
        End Sub

        Private Sub Predict(data As Byte(), output As Byte())
            Dim offset1 = Size.Width
            Dim offset2 = 2 * offset1
            Dim offset3 = 3 * offset1

            For i As Integer = 0 To Size.Height - 1
                Dim rowWordStart = i * Size.Width
                Dim rowOutputStart = i * BytesPerRow

                ' Pack together the individual bytes of the 32-bit words, high-order
                ' bytes before low-order bytes.
                For col As Integer = 0 To Size.Width - 1
                    Dim value = BitConverter.ToInt32(data, (rowWordStart + col) * 4)
                    Dim outIdx = rowOutputStart + col
                    output(outIdx) = CByte((value >> 24) And &HFF)
                    output(outIdx + offset1) = CByte((value >> 16) And &HFF)
                    output(outIdx + offset2) = CByte((value >> 8) And &HFF)
                    output(outIdx + offset3) = CByte(value And &HFF)
                Next

                ' Delta-encode the row
                Dim rowOutputEnd = rowOutputStart + BytesPerRow
                For idx As Integer = rowOutputEnd - 1 To rowOutputStart + 1 Step -1
                    output(idx) = CByte((CInt(output(idx)) - CInt(output(idx - 1)) + 256) Mod 256)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Unpredicts the raw decompressed image data into a 32-bpp bitmap with
        ''' native endianness.
        ''' </summary>
        Private Sub Unpredict(data As Byte(), output As Byte())
            Dim offset1 = Size.Width
            Dim offset2 = 2 * offset1
            Dim offset3 = 3 * offset1

            For i As Integer = 0 To Size.Height - 1
                Dim rowDataStart = i * BytesPerRow
                Dim rowDataEnd = rowDataStart + BytesPerRow

                ' Delta-decode each row
                For idx As Integer = rowDataStart + 1 To rowDataEnd - 1
                    data(idx) = CByte((CInt(data(idx)) + CInt(data(idx - 1))) Mod 256)
                Next

                ' Within each row, the individual bytes of the 32-bit words are
                ' packed together, high-order bytes before low-order bytes.
                ' We now unpack them into words.
                Dim rowWordStart = i * Size.Width
                For col As Integer = 0 To Size.Width - 1
                    Dim srcIdx = rowDataStart + col
                    Dim value As Integer =
                        (CInt(data(srcIdx)) << 24) Or
                        (CInt(data(srcIdx + offset1)) << 16) Or
                        (CInt(data(srcIdx + offset2)) << 8) Or
                        CInt(data(srcIdx + offset3))

                    Dim bytes = BitConverter.GetBytes(value)
                    Dim outByteIdx = (rowWordStart + col) * 4
                    output(outByteIdx) = bytes(0)
                    output(outByteIdx + 1) = bytes(1)
                    output(outByteIdx + 2) = bytes(2)
                    output(outByteIdx + 3) = bytes(3)
                Next
            Next
        End Sub

    End Class

End Namespace
