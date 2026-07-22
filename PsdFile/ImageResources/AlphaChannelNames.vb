'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/AlphaChannelNames.cs
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
Imports System.Collections.Generic

Namespace PhotoshopFile

    ''' <summary>
    ''' The names of the alpha channels
    ''' </summary>
    Public Class AlphaChannelNames
        Inherits ImageResource

        Public Overrides ReadOnly Property ID As ResourceID
            Get
                Return ResourceID.AlphaChannelNames
            End Get
        End Property

        Private _channelNames As New List(Of String)()
        Public ReadOnly Property ChannelNames As List(Of String)
            Get
                Return _channelNames
            End Get
        End Property

        Public Sub New()
            MyBase.New(String.Empty)
        End Sub

        Public Sub New(reader As PsdBinaryReader, name As String, resourceDataLength As Integer)
            MyBase.New(name)

            Dim endPosition = reader.BaseStream.Position + resourceDataLength

            ' Alpha channel names are Pascal strings, with no padding in-between.
            While reader.BaseStream.Position < endPosition
                Dim channelName = reader.ReadPascalString(1)
                ChannelNames.Add(channelName)
            End While
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            For Each channelName In ChannelNames
                writer.WritePascalString(channelName, 1)
            Next
        End Sub

    End Class

End Namespace
