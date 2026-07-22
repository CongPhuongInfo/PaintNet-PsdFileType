'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/ZipPredict16Image.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: the original reinterprets the byte buffer directly as a
' UInt16* and walks it word-by-word with pointer arithmetic - only valid
' because it never changes the machine's native byte order (it deliberately
' operates on the value as stored, not on a big/little-endian-specific
' value). VB.NET has no unsafe/pointer support, so each 16-bit word is
' instead read/written through BitConverter.ToUInt16/GetBytes at the
' matching byte offset, which reproduces the exact same native-endian
' in-place reinterpretation without pointers.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Drawing

Namespace PhotoshopFile.Compression

    Public Class ZipPredict16Image
        Inherits ImageData

        Private zipImage As ImageData

        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return True
            End Get
        End Property

        Public Sub New(zipData As Byte(), size As Size)
            MyBase.New(size, 16)

            ' 16-bitdepth images are delta-encoded word-by-word. The deltas
            ' are thus big-endian and must be reversed for further processing.
            Dim zipRawImage = New ZipImage(zipData, size, 16)
            zipImage = New EndianReverser(zipRawImage)
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            If buffer.Length = 0 Then
                Return
            End If

            zipImage.Read(buffer)
            Unpredict(buffer)
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            Return zipImage.ReadCompressed()
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            If array.Length = 0 Then
                Return
            End If

            Predict(array)
            zipImage.WriteInternal(array)
        End Sub

        Private Sub Predict(data As Byte())
            ' Delta-encode each row. Each "word" is 2 bytes at offset word*2.
            For i As Integer = 0 To Size.Height - 1
                Dim rowWordStart = i * Size.Width
                Dim rowWordEnd = rowWordStart + Size.Width

                ' Start with the last column in the row, walking backwards.
                For w As Integer = rowWordEnd - 1 To rowWordStart + 1 Step -1
                    Dim current = BitConverter.ToUInt16(data, w * 2)
                    Dim previous = BitConverter.ToUInt16(data, (w - 1) * 2)
                    ' 16-bit unsigned subtraction with wraparound (matches C#'s
                    ' unchecked ushort arithmetic in the original pointer code).
                    Dim delta = CUShort((CInt(current) - CInt(previous) + 65536) Mod 65536)
                    WriteUInt16(data, w * 2, delta)
                Next
            Next
        End Sub

        ''' <summary>
        ''' Unpredicts the decompressed, native-endian image data.
        ''' </summary>
        Private Sub Unpredict(data As Byte())
            ' Delta-decode each row.
            For i As Integer = 0 To Size.Height - 1
                Dim rowWordStart = i * Size.Width
                Dim rowWordEnd = rowWordStart + Size.Width

                ' Start with column index 1 on each row.
                For w As Integer = rowWordStart + 1 To rowWordEnd - 1
                    Dim current = BitConverter.ToUInt16(data, w * 2)
                    Dim previous = BitConverter.ToUInt16(data, (w - 1) * 2)
                    WriteUInt16(data, w * 2, CUShort((CInt(current) + CInt(previous)) And &HFFFF))
                Next
            Next
        End Sub

        Private Shared Sub WriteUInt16(buffer As Byte(), byteOffset As Integer, value As UInt16)
            Dim bytes = BitConverter.GetBytes(value)
            buffer(byteOffset) = bytes(0)
            buffer(byteOffset + 1) = bytes(1)
        End Sub

    End Class

End Namespace
