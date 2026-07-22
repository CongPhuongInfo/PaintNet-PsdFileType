'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/Mask.cs
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

Imports System
Imports System.Collections.Specialized
Imports System.Drawing

Namespace PhotoshopFile

    Public Class Mask

        ''' <summary>
        ''' The layer to which this mask belongs
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

        ''' <summary>
        ''' The rectangle enclosing the mask.
        ''' </summary>
        Public Property Rect As Rectangle

        Private _backgroundColor As Byte
        Public Property BackgroundColor As Byte
            Get
                Return _backgroundColor
            End Get
            Set(value As Byte)
                If (value <> 0) AndAlso (value <> 255) Then
                    Throw New PsdInvalidException("Mask background must be fully-opaque or fully-transparent.")
                End If
                _backgroundColor = value
            End Set
        End Property

        Private Shared positionVsLayerBit As Integer = BitVector32.CreateMask()
        Private Shared disabledBit As Integer = BitVector32.CreateMask(positionVsLayerBit)
        Private Shared invertOnBlendBit As Integer = BitVector32.CreateMask(disabledBit)

        Private flagsField As BitVector32
        Public ReadOnly Property Flags As BitVector32
            Get
                Return flagsField
            End Get
        End Property

        ''' <summary>
        ''' If true, the position of the mask is relative to the layer.
        ''' </summary>
        Public Property PositionVsLayer As Boolean
            Get
                Return flagsField(positionVsLayerBit)
            End Get
            Set(value As Boolean)
                flagsField(positionVsLayerBit) = value
            End Set
        End Property

        Public Property Disabled As Boolean
            Get
                Return flagsField(disabledBit)
            End Get
            Set(value As Boolean)
                flagsField(disabledBit) = value
            End Set
        End Property

        ''' <summary>
        ''' if true, invert the mask when blending.
        ''' </summary>
        Public Property InvertOnBlend As Boolean
            Get
                Return flagsField(invertOnBlendBit)
            End Get
            Set(value As Boolean)
                flagsField(invertOnBlendBit) = value
            End Set
        End Property

        ''' <summary>
        ''' Mask image data.
        ''' </summary>
        Public Property ImageData As Byte()

        Public Sub New(layer As Layer)
            Me.Layer = layer
            Me.flagsField = New BitVector32()
        End Sub

        Public Sub New(layer As Layer, rect As Rectangle, color As Byte, flags As BitVector32)
            Me.Layer = layer
            Me.Rect = rect
            Me.BackgroundColor = color
            Me.flagsField = flags
        End Sub

    End Class

    ''' <summary>
    ''' Mask info for a layer. Contains both the layer and user masks.
    ''' </summary>
    Public Class MaskInfo

        Public Property LayerMask As Mask

        Public Property UserMask As Mask

        ''' <summary>
        ''' Construct MaskInfo with null masks.
        ''' </summary>
        Public Sub New()
        End Sub

        Public Sub New(reader As PsdBinaryReader, layer As Layer)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, MaskInfo")

            Dim maskLength = reader.ReadUInt32()
            If maskLength <= 0 Then
                Return
            End If

            Dim startPosition = reader.BaseStream.Position
            Dim endPosition = startPosition + maskLength

            ' Read layer mask
            Dim rectangle = reader.ReadRectangle()
            Dim backgroundColor = reader.ReadByte()
            Dim flagsByte = reader.ReadByte()
            LayerMask = New Mask(layer, rectangle, backgroundColor, New BitVector32(CInt(flagsByte)))

            ' User mask is supplied separately when there is also a vector mask.
            If maskLength = 36 Then
                Dim userFlagsByte = reader.ReadByte()
                Dim userBackgroundColor = reader.ReadByte()
                Dim userRectangle = reader.ReadRectangle()
                UserMask = New Mask(layer, userRectangle, userBackgroundColor,
                    New BitVector32(CInt(userFlagsByte)))
            End If

            ' 20-byte mask data will end with padding.
            reader.BaseStream.Position = endPosition

            Util.DebugMessage(reader.BaseStream, "Load, End, MaskInfo")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub Save(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, MaskInfo")

            If LayerMask Is Nothing Then
                writer.Write(CUInt(0))
                Return
            End If

            Using New PsdBlockLengthWriter(writer)
                writer.Write(LayerMask.Rect)
                writer.Write(LayerMask.BackgroundColor)
                writer.Write(CByte(LayerMask.Flags.Data))

                If UserMask Is Nothing Then
                    ' Pad by 2 bytes to make the block length 20
                    writer.Write(CUShort(0))
                Else
                    writer.Write(CByte(UserMask.Flags.Data))
                    writer.Write(UserMask.BackgroundColor)
                    writer.Write(UserMask.Rect)
                End If
            End Using

            Util.DebugMessage(writer.BaseStream, "Save, End, MaskInfo")
        End Sub

    End Class

End Namespace
