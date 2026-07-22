'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/Layer.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2021 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: CreateMissingChannels used unsafe pointer code
' (fixed/Util.Fill(byte*, byte*, byte)) to fill a channel's image buffer with
' 255. Our Util.vb port exposes a safe array-based Util.Fill(buffer, start,
' count, value) instead - used directly below, no unsafe block needed.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Collections.Specialized
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq

Namespace PhotoshopFile

    <DebuggerDisplay("Name = {Name}")>
    Public Class Layer

        Private _psdFile As PsdFile
        Friend Property PsdFile As PsdFile
            Get
                Return _psdFile
            End Get
            Private Set(value As PsdFile)
                _psdFile = value
            End Set
        End Property

        ''' <summary>
        ''' The rectangle containing the contents of the layer.
        ''' </summary>
        Public Property Rect As Rectangle

        ''' <summary>
        ''' Image channels.
        ''' </summary>
        Private _channels As IChannelList
        Public Property Channels As IChannelList
            Get
                Return _channels
            End Get
            Private Set(value As IChannelList)
                _channels = value
            End Set
        End Property

        ''' <summary>
        ''' Returns alpha channel if it exists, otherwise Nothing.
        ''' </summary>
        Public ReadOnly Property AlphaChannel As Channel
            Get
                Return Channels.SingleOrDefault(Function(x) x.ID = -1)
            End Get
        End Property

        Private _blendModeKey As String
        ''' <summary>
        ''' Photoshop blend mode key for the layer
        ''' </summary>
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

        ''' <summary>
        ''' 0 = transparent ... 255 = opaque
        ''' </summary>
        Public Property Opacity As Byte

        ''' <summary>
        ''' false = base, true = non-base
        ''' </summary>
        Public Property Clipping As Boolean

        Private Shared protectTransBit As Integer = BitVector32.CreateMask()
        Private Shared visibleBit As Integer = BitVector32.CreateMask(protectTransBit)
        Private flags As New BitVector32()

        ''' <summary>
        ''' If true, the layer is visible.
        ''' </summary>
        Public Property Visible As Boolean
            Get
                Return Not flags(visibleBit)
            End Get
            Set(value As Boolean)
                flags(visibleBit) = Not value
            End Set
        End Property

        ''' <summary>
        ''' Protect the transparency
        ''' </summary>
        Public Property ProtectTrans As Boolean
            Get
                Return flags(protectTransBit)
            End Get
            Set(value As Boolean)
                flags(protectTransBit) = value
            End Set
        End Property

        ''' <summary>
        ''' The descriptive layer name
        ''' </summary>
        Public Property Name As String

        Public Property BlendingRangesData As BlendingRanges

        Public Property Masks As MaskInfo

        Public Property AdditionalInfo As List(Of LayerInfo)

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub New(psdFile As PsdFile)
            Me.PsdFile = psdFile
            Rect = Rectangle.Empty
            Channels = New ChannelList()
            BlendModeKey = PsdBlendMode.Normal
            AdditionalInfo = New List(Of LayerInfo)()
        End Sub

        Friend Sub New(reader As PsdBinaryReader, psdFile As PsdFile)
            Me.New(psdFile)

            Util.DebugMessage(reader.BaseStream, "Load, Begin, Layer")

            Rect = reader.ReadRectangle()

            ' Channel headers
            Dim numberOfChannels As Integer = reader.ReadUInt16()
            For channel As Integer = 0 To numberOfChannels - 1
                Dim ch = New Channel(reader, Me)
                Channels.Add(ch)
            Next

            ' Layer blending
            Dim signature = reader.ReadAsciiChars(4)
            If signature <> "8BIM" Then
                Throw New PsdInvalidException("Invalid signature in layer header.")
            End If
            BlendModeKey = reader.ReadAsciiChars(4)
            Opacity = reader.ReadByte()
            Clipping = reader.ReadBoolean()

            Dim flagsByte = reader.ReadByte()
            flags = New BitVector32(CInt(flagsByte))
            reader.ReadByte() ' Padding

            ' Variable-length data
            Dim extraDataSize = reader.ReadUInt32()
            Dim extraDataStartPosition = reader.BaseStream.Position
            Dim extraDataEndPosition As Long = extraDataStartPosition + extraDataSize

            Masks = New MaskInfo(reader, Me)
            BlendingRangesData = New BlendingRanges(reader, Me)
            Name = reader.ReadPascalString(4)
            LayerInfoFactory.LoadAll(reader, PsdFile, AdditionalInfo, extraDataEndPosition, False)

            For Each adjustmentInfo In AdditionalInfo
                Select Case adjustmentInfo.Key
                    Case "luni"
                        Name = CType(adjustmentInfo, LayerUnicodeName).Name
                End Select
            Next

            Util.DebugMessage(reader.BaseStream, $"Load, End, Layer, {Name}")

            PsdFile.LoadContext.OnLoadLayerHeader(Me)
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Create ImageData for any missing channels.
        ''' </summary>
        Public Sub CreateMissingChannels()
            Dim channelCount = Me.PsdFile.ColorMode.MinChannelCount()
            For id As Short = 0 To CShort(channelCount - 1)
                If Not Me.Channels.ContainsId(id) Then
                    Dim size = Me.Rect.Height * Me.Rect.Width

                    Dim ch = New Channel(id, Me)
                    ch.ImageData = New Byte(If(size > 0, size - 1, -1)) {}

                    If size > 0 Then
                        Util.Fill(ch.ImageData, 0, size, CByte(255))
                    End If

                    Me.Channels.Add(ch)
                End If
            Next
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub PrepareSave()
            For Each ch In Channels
                ch.CompressImageData()
            Next

            ' Create or update the Unicode layer name to be consistent with the
            ' ANSI layer name.
            Dim layerUnicodeNames = AdditionalInfo.Where(Function(x) TypeOf x Is LayerUnicodeName)
            If layerUnicodeNames.Count() > 1 Then
                Throw New PsdInvalidException($"{NameOf(Layer)} can only have one {NameOf(LayerUnicodeName)}.")
            End If

            Dim unicodeNameInfo = CType(layerUnicodeNames.FirstOrDefault(), LayerUnicodeName)
            If unicodeNameInfo Is Nothing Then
                unicodeNameInfo = New LayerUnicodeName(Name)
                AdditionalInfo.Add(unicodeNameInfo)
            ElseIf unicodeNameInfo.Name <> Name Then
                unicodeNameInfo.Name = Name
            End If
        End Sub

        Public Sub Save(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Layer")

            writer.Write(Rect)

            '-----------------------------------------------------------------------

            writer.Write(CShort(Channels.Count))
            For Each ch In Channels
                ch.Save(writer)
            Next

            '-----------------------------------------------------------------------

            writer.WriteAsciiChars("8BIM")
            writer.WriteAsciiChars(BlendModeKey)
            writer.Write(Opacity)
            writer.Write(Clipping)

            writer.Write(CByte(flags.Data))
            writer.Write(CByte(0))

            '-----------------------------------------------------------------------

            Using New PsdBlockLengthWriter(writer)
                Masks.Save(writer)
                BlendingRangesData.Save(writer)

                Dim namePosition = writer.BaseStream.Position

                ' Legacy layer name is limited to 31 bytes. Unicode layer name
                ' can be much longer.
                writer.WritePascalString(Name, 4, 31)

                For Each info As LayerInfo In AdditionalInfo
                    info.Save(writer, globalLayerInfo:=False, isLargeDocument:=PsdFile.IsLargeDocument)
                Next
            End Using

            Util.DebugMessage(writer.BaseStream, $"Save, End, Layer, {Name}")
        End Sub

    End Class

End Namespace
