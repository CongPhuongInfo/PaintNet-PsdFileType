'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/Channel.cs
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq

Imports PhotoshopFile.Compression

Namespace PhotoshopFile

    ' ChannelList is exposed as an ICollection so that it cannot be accidentally
    ' indexed by ID.
    Public Interface IChannelList
        Inherits ICollection(Of Channel)

        Function ToIdArray() As Channel()

        Function GetId(id As Integer) As Channel

        Function ContainsId(id As Integer) As Boolean

    End Interface

    Friend Class ChannelList
        Inherits List(Of Channel)
        Implements IChannelList

        Public Sub New()
            MyBase.New()
        End Sub

        ''' <summary>
        ''' Returns channels with nonnegative IDs as an array, so that accessing
        ''' a channel by Id can be optimized into direct array access rather than
        ''' being implemented as a List scan.
        ''' </summary>
        ''' <remarks>
        ''' This optimization is crucial for blitting lots of pixels back and
        ''' forth between Photoshop's per-channel representation, and Paint.NET's
        ''' per-pixel BGRA representation.
        ''' </remarks>
        Public Function ToIdArray() As Channel() Implements IChannelList.ToIdArray
            Dim maxId = Me.Max(Function(x) x.ID)
            Dim idArray(CInt(maxId)) As Channel
            For Each channel In Me
                If channel.ID >= 0 Then
                    idArray(channel.ID) = channel
                End If
            Next
            Return idArray
        End Function

        Public Function GetId(id As Integer) As Channel Implements IChannelList.GetId
            Return Me.Single(Function(x) x.ID = id)
        End Function

        Public Function ContainsId(id As Integer) As Boolean Implements IChannelList.ContainsId
            Return Me.Exists(Function(x) x.ID = id)
        End Function

    End Class

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    <DebuggerDisplay("ID = {ID}")>
    Public Class Channel

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

        ''' <summary>
        ''' Channel ID.
        ''' <list type="bullet">
        ''' <item>-1 = transparency mask</item>
        ''' <item>-2 = user-supplied layer mask, or vector mask</item>
        ''' <item>-3 = user-supplied layer mask, if channel -2 contains a vector mask</item>
        ''' <item>
        ''' Nonnegative channel IDs give the actual image channels, in the
        ''' order defined by the colormode. For example, 0, 1, 2 = R, G, B.
        ''' </item>
        ''' </list>
        ''' </summary>
        Public Property ID As Short

        Public ReadOnly Property Rect As Rectangle
            Get
                Select Case ID
                    Case -2
                        Return Layer.Masks.LayerMask.Rect
                    Case -3
                        Return Layer.Masks.UserMask.Rect
                    Case Else
                        Return Layer.Rect
                End Select
            End Get
        End Property

        ''' <summary>
        ''' Total length of the channel data, including compression headers.
        ''' </summary>
        Public Property Length As Long

        ''' <summary>
        ''' Raw image data for this color channel, in compressed on-disk format.
        ''' </summary>
        ''' <remarks>
        ''' If Nothing, the ImageData will be automatically compressed during save.
        ''' </remarks>
        Public Property ImageDataRaw As Byte()

        ''' <summary>
        ''' Decompressed image data for this color channel.
        ''' </summary>
        ''' <remarks>
        ''' When making changes to the ImageData, set ImageDataRaw to Nothing so
        ''' that the correct data will be compressed during save.
        ''' </remarks>
        Public Property ImageData As Byte()

        ''' <summary>
        ''' Image compression method used.
        ''' </summary>
        Public Property ImageCompression As ImageCompression

        ''' <summary>
        ''' RLE-compressed length of each row.
        ''' </summary>
        Public Property RleRowLengths As RleRowLengths

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub New(id As Short, layer As Layer)
            Me.ID = id
            Me.Layer = layer
        End Sub

        Friend Sub New(reader As PsdBinaryReader, layer As Layer)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, Channel")

            ID = reader.ReadInt16()
            Length = If(layer.PsdFile.IsLargeDocument, reader.ReadInt64(), CLng(reader.ReadInt32()))
            Me.Layer = layer

            Util.DebugMessage(reader.BaseStream, $"Load, End, Channel, {ID}")
        End Sub

        Friend Sub Save(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Channel")

            writer.Write(ID)
            If Layer.PsdFile.IsLargeDocument Then
                writer.Write(Length)
            Else
                writer.Write(CInt(Length))
            End If

            Util.DebugMessage(writer.BaseStream, $"Save, End, Channel, {ID}")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Friend Sub LoadPixelData(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, Channel image")

            If Length = 0 Then
                ImageCompression = ImageCompression.Raw
                ImageDataRaw = New Byte() {}
                Return
            End If

            Dim endPosition = reader.BaseStream.Position + Me.Length
            ImageCompression = CType(reader.ReadInt16(), ImageCompression)
            Dim longDataLength = Me.Length - 2
            Util.CheckByteArrayLength(longDataLength)
            Dim dataLength = CInt(longDataLength)

            Select Case ImageCompression
                Case PhotoshopFile.ImageCompression.Raw
                    ImageDataRaw = reader.ReadBytes(dataLength)

                Case PhotoshopFile.ImageCompression.Rle
                    ' RLE row lengths
                    RleRowLengths = New RleRowLengths(reader, Rect.Height, Layer.PsdFile.IsLargeDocument)
                    Dim rleDataLength = CInt(endPosition - reader.BaseStream.Position)
                    Debug.Assert(rleDataLength = RleRowLengths.Total,
                        "RLE row lengths do not sum to length of channel image data.")

                    ' The PSD specification states that rows are padded to even sizes.
                    ' However, Photoshop doesn't actually do this. RLE rows can have
                    ' odd lengths in the header, and there is no padding between rows.
                    ImageDataRaw = reader.ReadBytes(rleDataLength)

                Case PhotoshopFile.ImageCompression.Zip, PhotoshopFile.ImageCompression.ZipPrediction
                    ImageDataRaw = reader.ReadBytes(dataLength)
            End Select

            Util.DebugMessage(reader.BaseStream, $"Load, End, Channel image, {ID}")
            Debug.Assert(reader.BaseStream.Position = endPosition, "Pixel data was not fully read in.")
        End Sub

        ''' <summary>
        ''' Decodes the raw image data from the compressed on-disk format into
        ''' an uncompressed bitmap, in native byte order.
        ''' </summary>
        Public Sub DecodeImageData()
            If (ImageCompression = PhotoshopFile.ImageCompression.Raw) AndAlso (Layer.PsdFile.BitDepth <= 8) Then
                ImageData = ImageDataRaw
                Return
            End If

            Dim image = ImageDataFactory.Create(Me, ImageDataRaw)
            Dim longLength = CLng(image.BytesPerRow) * Rect.Height
            Util.CheckByteArrayLength(longLength)
            ImageData = New Byte(CInt(longLength) - 1) {}
            image.Read(ImageData)
        End Sub

        ''' <summary>
        ''' Compresses the image data.
        ''' </summary>
        Public Sub CompressImageData()
            ' Do not recompress if compressed data is already present.
            If ImageDataRaw IsNot Nothing Then
                Return
            End If

            If ImageData Is Nothing Then
                Return
            End If

            If ImageCompression = PhotoshopFile.ImageCompression.Rle Then
                RleRowLengths = New RleRowLengths(Rect.Height)
            End If

            Dim compressor = ImageDataFactory.Create(Me, Nothing)
            compressor.Write(ImageData)
            ImageDataRaw = compressor.ReadCompressed()

            Length = 2 + ImageDataRaw.Length
            If ImageCompression = PhotoshopFile.ImageCompression.Rle Then
                Dim rowLengthSize = If(Layer.PsdFile.IsLargeDocument, 4, 2)
                Length += rowLengthSize * Rect.Height
            End If
        End Sub

        Friend Sub SavePixelData(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Channel image")

            writer.Write(CShort(ImageCompression))
            If ImageDataRaw Is Nothing Then
                Return
            End If

            If ImageCompression = PhotoshopFile.ImageCompression.Rle Then
                RleRowLengths.Write(writer, Layer.PsdFile.IsLargeDocument)
            End If
            writer.Write(ImageDataRaw)

            Util.DebugMessage(writer.BaseStream, $"Save, End, Channel image, {ID}")
        End Sub

    End Class

End Namespace
