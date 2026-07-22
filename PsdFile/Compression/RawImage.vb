'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/RawImage.cs
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
Imports System.Drawing

Namespace PhotoshopFile.Compression

    Friend Class RawImage
        Inherits ImageData

        Private data As Byte()

        Protected Overrides ReadOnly Property AltersWrittenData As Boolean
            Get
                Return False
            End Get
        End Property

        Public Sub New(data As Byte(), size As Size, bitDepth As Integer)
            MyBase.New(size, bitDepth)
            Me.data = data
        End Sub

        Friend Overrides Sub Read(buffer As Byte())
            Array.Copy(data, buffer, data.Length)
        End Sub

        Public Overrides Function ReadCompressed() As Byte()
            Return data
        End Function

        Friend Overrides Sub WriteInternal(array As Byte())
            data = array
        End Sub

    End Class

End Namespace
