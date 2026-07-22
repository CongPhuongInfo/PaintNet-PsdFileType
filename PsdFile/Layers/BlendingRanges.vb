'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/BlendingRanges.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Namespace PhotoshopFile

    Public Class BlendingRanges

        ''' <summary>
        ''' The layer to which this channel belongs
        ''' </summary>
        Private _layer As Layer
        Public Property Layer As Layer
            Get
                Return _layer
            End Get
            Private Set(value As Layer)
                _layer = value
            End Set
        End Property

        Public Property Data As Byte()

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub New(layer As Layer)
            Me.Layer = layer
            Data = New Byte() {}
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub New(reader As PsdBinaryReader, layer As Layer)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, BlendingRanges")

            Me.Layer = layer
            Dim dataLength = reader.ReadInt32()
            If dataLength <= 0 Then
                Return
            End If

            Data = reader.ReadBytes(dataLength)

            Util.DebugMessage(reader.BaseStream, "Load, End, BlendingRanges")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub Save(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, BlendingRanges")

            If Data Is Nothing Then
                writer.Write(CUInt(0))
                Return
            End If

            writer.Write(CUInt(Data.Length))
            writer.Write(Data)

            Util.DebugMessage(writer.BaseStream, "Save, End, BlendingRanges")
        End Sub

    End Class

End Namespace
