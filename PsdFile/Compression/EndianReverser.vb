'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/EndianReverser.cs
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

Namespace PhotoshopFile.Compression

    Public Class EndianReverser
        Inherits ImageData

        Private imageData As ImageData

        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return True
            End Get
        End Property

        Public Sub New(imageData As ImageData)
            MyBase.New(imageData.Size, imageData.BitDepth)
            Me.imageData = imageData
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            imageData.Read(buffer)

            Dim numPixels = Size.Width * Size.Height
            If numPixels = 0 Then
                Return
            End If
            Util.SwapByteArray(BitDepth, buffer, 0, numPixels)
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            Return imageData.ReadCompressed()
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            ' Reverse endianness before passing on to underlying compressor
            If array.Length > 0 Then
                Dim numPixels = array.Length \ BytesPerRow * Size.Width
                Util.SwapByteArray(BitDepth, array, 0, numPixels)
            End If

            imageData.Write(array)
        End Sub

    End Class

End Namespace
