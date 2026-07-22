'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/VersionInfo.cs
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

Namespace PhotoshopFile

    Public Class VersionInfo
        Inherits ImageResource

        Public Overrides ReadOnly Property ID As ResourceID
            Get
                Return ResourceID.VersionInfo
            End Get
        End Property

        Public Property Version As UInt32

        Public Property HasRealMergedData As Boolean

        Public Property ReaderName As String

        Public Property WriterName As String

        Public Property FileVersion As UInt32

        Public Sub New()
            MyBase.New(String.Empty)
        End Sub

        Public Sub New(reader As PsdBinaryReader, name As String)
            MyBase.New(name)

            Version = reader.ReadUInt32()
            HasRealMergedData = reader.ReadBoolean()
            ReaderName = reader.ReadUnicodeString()
            WriterName = reader.ReadUnicodeString()
            FileVersion = reader.ReadUInt32()
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            writer.Write(Version)
            writer.Write(HasRealMergedData)
            writer.WriteUnicodeString(ReaderName)
            writer.WriteUnicodeString(WriterName)
            writer.Write(FileVersion)
        End Sub

    End Class

End Namespace
