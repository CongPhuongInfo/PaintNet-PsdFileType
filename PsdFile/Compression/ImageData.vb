'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/ImageData.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2016 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Drawing

Namespace PhotoshopFile.Compression

    Public MustInherit Class ImageData

        Public Property BitDepth As Integer
        Public Property BytesPerRow As Integer
        Public Property Size As Size

        Protected MustOverride ReadOnly Property AltersWrittenData As Boolean

        Protected Sub New(size As Size, bitDepth As Integer)
            Me.Size = size
            Me.BitDepth = bitDepth
            Me.BytesPerRow = Util.BytesPerRow(size, bitDepth)
        End Sub

        ''' <summary>
        ''' Reads decompressed image data.
        ''' </summary>
        Public Overridable Function Read() As Byte()
            Dim imageLongLength As Long = CLng(BytesPerRow) * Size.Height
            Util.CheckByteArrayLength(imageLongLength)

            Dim buffer(CInt(imageLongLength) - 1) As Byte
            Read(buffer)
            Return buffer
        End Function

        Friend MustOverride Sub Read(buffer As Byte())

        ''' <summary>
        ''' Reads compressed image data.
        ''' </summary>
        Public MustOverride Function ReadCompressed() As Byte()

        ''' <summary>
        ''' Writes rows of image data into compressed format.
        ''' </summary>
        ''' <param name="array">An array containing the data to be compressed.</param>
        Public Sub Write(array As Byte())
            Dim imageLength As Long = CLng(BytesPerRow) * Size.Height
            Dim toWrite = array
            If toWrite.Length <> imageLength Then
                Throw New ArgumentException(
                    "Array length is not equal to image length.", NameOf(array))
            End If

            If AltersWrittenData Then
                toWrite = CType(toWrite.Clone(), Byte())
            End If

            WriteInternal(toWrite)
        End Sub

        Friend MustOverride Sub WriteInternal(array As Byte())

    End Class

End Namespace
