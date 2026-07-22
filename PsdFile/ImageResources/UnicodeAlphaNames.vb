'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/UnicodeAlphaNames.cs
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
    ''' The names of the alpha channels.
    ''' </summary>
    Public Class UnicodeAlphaNames
        Inherits ImageResource

        Public Overrides ReadOnly Property ID As ResourceID
            Get
                Return ResourceID.UnicodeAlphaNames
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

            While reader.BaseStream.Position < endPosition
                Dim channelName = reader.ReadUnicodeString()

                ' Photoshop writes out a null terminator for Unicode alpha names.
                ' There is no null terminator on other Unicode strings in PSD files.
                If channelName.EndsWith(vbNullChar) Then
                    channelName = channelName.Substring(0, channelName.Length - 1)
                End If
                ChannelNames.Add(channelName)
            End While
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            For Each channelName In ChannelNames
                ' We must add a null terminator because Photoshop always strips the
                ' last character of a Unicode alpha name, even if it is not null.
                writer.WriteUnicodeString(channelName & vbNullChar)
            Next
        End Sub

    End Class

End Namespace
