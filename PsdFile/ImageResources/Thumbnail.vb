'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/Thumbnail.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2013 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging

Namespace PhotoshopFile

    ''' <summary>
    ''' Summary description for Thumbnail.
    ''' </summary>
    Public Class Thumbnail
        Inherits RawImageResource

        Private _image As Bitmap
        Public Property Image As Bitmap
            Get
                Return _image
            End Get
            Private Set(value As Bitmap)
                _image = value
            End Set
        End Property

        Public Sub New(id As ResourceID, name As String)
            MyBase.New(id, name)
        End Sub

        Public Sub New(psdReader As PsdBinaryReader, id As ResourceID, name As String, numBytes As Integer)
            MyBase.New(psdReader, "8BIM", id, name, numBytes)

            Using memoryStream As New MemoryStream(Data)
                Using reader As New PsdBinaryReader(memoryStream, psdReader)
                    Const HEADER_LENGTH As Integer = 28
                    Dim format = reader.ReadUInt32()
                    Dim width = reader.ReadUInt32()
                    Dim height = reader.ReadUInt32()
                    Dim widthBytes = reader.ReadUInt32()
                    Dim size = reader.ReadUInt32()
                    Dim compressedSize = reader.ReadUInt32()
                    Dim bitPerPixel = reader.ReadUInt16()
                    Dim planes = reader.ReadUInt16()

                    ' Raw RGB bitmap
                    If format = 0 Then
                        Image = New Bitmap(CInt(width), CInt(height), PixelFormat.Format24bppRgb)

                        ' JPEG bitmap
                    ElseIf format = 1 Then
                        Dim imgData = reader.ReadBytes(numBytes - HEADER_LENGTH)
                        Using stream As New MemoryStream(imgData)
                            Dim bitmap = New Bitmap(stream)
                            Image = CType(bitmap.Clone(), Bitmap)
                        End Using

                        ' Reverse BGR pixels from old thumbnail format
                        If id = ResourceID.ThumbnailBgr Then
                            'For y As Integer = 0 To Image.Height - 1
                            '  For x As Integer = 0 To Image.Width - 1
                            '    Dim c = Image.GetPixel(x, y)
                            '    Dim c2 = Color.FromArgb(c.B, c.G, c.R)
                            '    Image.SetPixel(x, y, c)
                            '  Next
                            'Next
                        End If
                    Else
                        Throw New PsdInvalidException("Unknown thumbnail format.")
                    End If
                End Using
            End Using
        End Sub

    End Class

End Namespace
