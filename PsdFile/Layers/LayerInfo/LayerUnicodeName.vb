'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/LayerInfo/LayerUnicodeName.cs
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

    Public Class LayerUnicodeName
        Inherits LayerInfo

        Public Overrides ReadOnly Property Signature As String
            Get
                Return "8BIM"
            End Get
        End Property

        Public Overrides ReadOnly Property Key As String
            Get
                Return "luni"
            End Get
        End Property

        Public Property Name As String

        Public Sub New(name As String)
            Me.Name = name
        End Sub

        Public Sub New(reader As PsdBinaryReader)
            Name = reader.ReadUnicodeString()
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            Dim startPosition = writer.BaseStream.Position

            writer.WriteUnicodeString(Name)
            writer.WritePadding(startPosition, 4)
        End Sub

    End Class

End Namespace
