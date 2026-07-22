'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/ZipImage.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2024 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.IO.Compression

Namespace PhotoshopFile.Compression

    Public Class ZipImage
        Inherits ImageData

        Private zipDataStream As MemoryStream
        Private zipStream As DeflateStream

        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return False
            End Get
        End Property

        Public Sub New(data As Byte(), size As Size, bitDepth As Integer)
            MyBase.New(size, bitDepth)
            If data Is Nothing Then
                InitCompress()
            Else
                InitDecompress(data)
            End If
        End Sub

        Private Sub InitCompress()
            zipDataStream = New MemoryStream()

            ' Write 2-byte zlib (RFC 1950) header
            '
            ' CMF Compression Method and flags:
            '   CM     0:3 = 8 = deflate
            '   CINFO  4:7 = 4 = undefined, RFC 1950 only defines CINFO = 8
            '
            ' FLG Flags:
            '   FCHECK  0:4 = 9 = check bits for CMF and FLG
            '   FDICT     5 = 0 = no preset dictionary
            '   FLEVEL  6:7 = 2 = default compression level

            zipDataStream.WriteByte(&H48)
            zipDataStream.WriteByte(&H89)
            zipStream = New DeflateStream(zipDataStream, CompressionMode.Compress, True)
        End Sub

        Private Sub InitDecompress(data As Byte())
            zipDataStream = New MemoryStream(data)

            ' .NET implements Deflate (RFC 1951) but not zlib (RFC 1950),
            ' so we have to skip the first two bytes.
            zipDataStream.ReadByte()
            zipDataStream.ReadByte()
            zipStream = New DeflateStream(zipDataStream, CompressionMode.Decompress, True)
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            Dim longTotalBytesToRead As Long = CLng(Size.Height) * BytesPerRow
            Util.CheckByteArrayLength(longTotalBytesToRead)

            Dim totalBytesToRead As Integer = CInt(longTotalBytesToRead)
            Dim totalBytesRead As Integer = 0
            While totalBytesRead < totalBytesToRead
                Dim bytesRead = zipStream.Read(buffer, totalBytesRead, totalBytesToRead - totalBytesRead)
                If bytesRead = 0 Then
                    Exit While
                End If
                totalBytesRead += bytesRead
            End While

            If totalBytesRead <> totalBytesToRead Then
                Throw New Exception("ZIP stream was not fully decompressed.")
            End If
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            ' Write out the last block. (Flush leaves the last block open.)
            zipStream.Close()

            ' Do not write the zlib header when the image data is empty
            Dim result = If(zipDataStream.Length = 2, New Byte() {}, zipDataStream.ToArray())
            Return result
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            zipStream.Write(array, 0, array.Length)
        End Sub

    End Class

End Namespace
