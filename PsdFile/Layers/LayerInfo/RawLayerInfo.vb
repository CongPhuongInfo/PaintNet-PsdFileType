'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/LayerInfo/RawLayerInfo.cs
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

Imports System.Diagnostics

Namespace PhotoshopFile

    <DebuggerDisplay("Layer Info: {key}")>
    Public Class RawLayerInfo
        Inherits LayerInfo

        Private _signature As String
        Public Overrides ReadOnly Property Signature As String
            Get
                Return _signature
            End Get
        End Property

        Private _key As String
        Public Overrides ReadOnly Property Key As String
            Get
                Return _key
            End Get
        End Property

        Private _data As Byte()
        Public Property Data As Byte()
            Get
                Return _data
            End Get
            Private Set(value As Byte())
                _data = value
            End Set
        End Property

        Public Sub New(key As String, Optional signature As String = "8BIM")
            Me._signature = signature
            Me._key = key
        End Sub

        Public Sub New(reader As PsdBinaryReader, signature As String, key As String, dataLength As Long)
            Me._signature = signature
            Me._key = key

            Util.CheckByteArrayLength(dataLength)
            Data = reader.ReadBytes(CInt(dataLength))
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            writer.Write(Data)
        End Sub

    End Class

End Namespace
