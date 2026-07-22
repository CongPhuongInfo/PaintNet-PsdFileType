'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/LayerInfo/LayerSectionInfo.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System

Namespace PhotoshopFile

    Public Enum LayerSectionType
        Layer = 0
        OpenFolder = 1
        ClosedFolder = 2
        SectionDivider = 3
    End Enum

    Public Enum LayerSectionSubtype
        Normal = 0
        SceneGroup = 1
    End Enum

    ''' <summary>
    ''' Layer sections are known as Groups in the Photoshop UI.
    ''' </summary>
    Public Class LayerSectionInfo
        Inherits LayerInfo

        Public Overrides ReadOnly Property Signature As String
            Get
                Return "8BIM"
            End Get
        End Property

        Private _key As String
        Public Overrides ReadOnly Property Key As String
            Get
                Return _key
            End Get
        End Property

        Public Property SectionType As LayerSectionType

        Private subtypeField As LayerSectionSubtype?
        Public Property Subtype As LayerSectionSubtype
            Get
                Return If(subtypeField, LayerSectionSubtype.Normal)
            End Get
            Set(value As LayerSectionSubtype)
                subtypeField = value
            End Set
        End Property

        Private _blendModeKey As String
        Public Property BlendModeKey As String
            Get
                Return _blendModeKey
            End Get
            Set(value As String)
                If value.Length <> 4 Then
                    Throw New ArgumentException($"{NameOf(BlendModeKey)} must be 4 characters in length.")
                End If
                _blendModeKey = value
            End Set
        End Property

        Public Sub New(reader As PsdBinaryReader, key As String, dataLength As Integer)
            ' The key for layer section info is documented to be "lsct". However,
            ' some Photoshop files use the undocumented key "lsdk", with apparently
            ' the same data format.
            Me._key = key

            SectionType = CType(reader.ReadInt32(), LayerSectionType)
            If dataLength >= 12 Then
                Dim signature = reader.ReadAsciiChars(4)
                If signature <> "8BIM" Then
                    Throw New PsdInvalidException("Invalid section divider signature.")
                End If

                BlendModeKey = reader.ReadAsciiChars(4)
                If dataLength >= 16 Then
                    Subtype = CType(reader.ReadInt32(), LayerSectionSubtype)
                End If
            End If
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            writer.Write(CInt(SectionType))
            If BlendModeKey IsNot Nothing Then
                writer.WriteAsciiChars("8BIM")
                writer.WriteAsciiChars(BlendModeKey)
                If subtypeField IsNot Nothing Then
                    writer.Write(CInt(Subtype))
                End If
            End If
        End Sub

    End Class

End Namespace
