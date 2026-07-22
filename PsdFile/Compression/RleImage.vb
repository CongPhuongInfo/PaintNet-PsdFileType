'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/RleImage.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO

Namespace PhotoshopFile.Compression

    Friend Class RleImage
        Inherits ImageData

        Private rleData As Byte()
        Private rleRowLengths As RleRowLengths

        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return False
            End Get
        End Property

        Public Sub New(rleData As Byte(), rleRowLengths As RleRowLengths,
                        size As Size, bitDepth As Integer)
            MyBase.New(size, bitDepth)
            Me.rleData = rleData
            Me.rleRowLengths = rleRowLengths
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            Dim rleStream = New MemoryStream(rleData)
            Dim rleReader = New RleReader(rleStream)
            Dim bufferIndex = 0
            For i As Integer = 0 To Size.Height - 1
                Dim bytesRead = rleReader.Read(buffer, bufferIndex, BytesPerRow)
                If bytesRead <> BytesPerRow Then
                    Throw New Exception("RLE row decompressed to unexpected length.")
                End If
                bufferIndex += bytesRead
            Next
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            Return rleData
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            If rleData IsNot Nothing Then
                Throw New Exception("Cannot write to RLE image in Decompress mode.")
            End If

            Using dataStream As New MemoryStream()
                Dim rleWriter = New RleWriter(dataStream)
                For row As Integer = 0 To Size.Height - 1
                    Dim rowIndex = row * BytesPerRow
                    rleRowLengths(row) = rleWriter.Write(array, rowIndex, BytesPerRow)
                Next

                ' Save compressed data
                dataStream.Flush()
                rleData = dataStream.ToArray()
                Debug.Assert(rleRowLengths.Total = rleData.Length,
                    "RLE row lengths do not sum to the compressed data length.")
            End Using
        End Sub

    End Class

End Namespace
