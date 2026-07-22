'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/PsdFile.cs
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
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Namespace PhotoshopFile

    Public Enum PsdColorMode
        Bitmap = 0
        Grayscale = 1
        Indexed = 2
        RGB = 3
        CMYK = 4
        Multichannel = 7
        Duotone = 8
        Lab = 9
    End Enum

    Public Enum PsdFileVersion As Short
        Psd = 1
        PsbLargeDocument = 2
    End Enum

    Public Class PsdFile

#Region "Constructors"

        Public Sub New(Optional version As PsdFileVersion = PsdFileVersion.Psd)
            Me.Version = version

            BaseLayer = New Layer(Me)
            ImageResources = New ImageResources()
            Layers = New List(Of Layer)()
            AdditionalInfo = New List(Of LayerInfo)()
        End Sub

        Public Sub New(filename As String, loadContext As LoadContext)
            Me.New()
            Using stream As New FileStream(filename, FileMode.Open)
                Load(stream, loadContext)
            End Using
        End Sub

        Public Sub New(stream As Stream, loadContext As LoadContext)
            Me.New()
            Load(stream, loadContext)
        End Sub

#End Region

#Region "Load and save"

        Private _loadContext As LoadContext
        Friend Property LoadContext As LoadContext
            Get
                Return _loadContext
            End Get
            Private Set(value As LoadContext)
                _loadContext = value
            End Set
        End Property

        Private Sub Load(stream As Stream, loadContext As LoadContext)
            Me.LoadContext = loadContext
            Dim reader = New PsdBinaryReader(stream, loadContext.Encoding)

            LoadHeader(reader)
            LoadColorModeData(reader)
            LoadImageResources(reader)
            LoadLayerAndMaskInfo(reader)

            LoadImage(reader)
            DecompressImages()
        End Sub

        Public Sub Save(fileName As String, encoding As Encoding)
            Using stream As New FileStream(fileName, FileMode.Create)
                Save(stream, encoding)
            End Using
        End Sub

        Public Sub Save(stream As Stream, encoding As Encoding)
            PrepareSave()

            Using writer As New PsdBinaryWriter(stream, encoding)
                SaveHeader(writer)
                SaveColorModeData(writer)
                SaveImageResources(writer)
                SaveLayerAndMaskInfo(writer)
                SaveImage(writer)
            End Using
        End Sub

#End Region

#Region "Header"

        ''' <summary>
        ''' Photoshop file format version.
        ''' </summary>
        Private _version As PsdFileVersion
        Public Property Version As PsdFileVersion
            Get
                Return _version
            End Get
            Private Set(value As PsdFileVersion)
                _version = value
            End Set
        End Property

        Public ReadOnly Property IsLargeDocument As Boolean
            Get
                Return Version = PsdFileVersion.PsbLargeDocument
            End Get
        End Property

        Private channelCountField As Int16
        ''' <summary>
        ''' The number of channels in the image, including any alpha channels.
        ''' </summary>
        Public Property ChannelCount As Int16
            Get
                Return channelCountField
            End Get
            Set(value As Int16)
                If value < 1 OrElse value > 56 Then
                    Throw New ArgumentException("Number of channels must be from 1 to 56.")
                End If
                channelCountField = value
            End Set
        End Property

        Private Sub CheckDimension(dimension As Integer)
            If dimension < 1 Then
                Throw New ArgumentException("Image dimension must be at least 1.")
            End If
            If (Version = PsdFileVersion.Psd) AndAlso (dimension > 30000) Then
                Throw New ArgumentException("PSD image dimension cannot exceed 30000.")
            End If
            If (Version = PsdFileVersion.PsbLargeDocument) AndAlso (dimension > 300000) Then
                Throw New ArgumentException("PSB image dimension cannot exceed 300000.")
            End If
        End Sub

        ''' <summary>
        ''' The height of the image in pixels.
        ''' </summary>
        Public Property RowCount As Integer
            Get
                Return Me.BaseLayer.Rect.Height
            End Get
            Set(value As Integer)
                CheckDimension(value)
                BaseLayer.Rect = New Rectangle(0, 0, BaseLayer.Rect.Width, value)
            End Set
        End Property

        ''' <summary>
        ''' The width of the image in pixels.
        ''' </summary>
        Public Property ColumnCount As Integer
            Get
                Return Me.BaseLayer.Rect.Width
            End Get
            Set(value As Integer)
                CheckDimension(value)
                BaseLayer.Rect = New Rectangle(0, 0, value, BaseLayer.Rect.Height)
            End Set
        End Property

        Private bitDepthField As Integer
        ''' <summary>
        ''' The number of bits per channel. Supported values are 1, 8, 16, and 32.
        ''' </summary>
        Public Property BitDepth As Integer
            Get
                Return bitDepthField
            End Get
            Set(value As Integer)
                Select Case value
                    Case 1, 8, 16, 32
                        bitDepthField = value
                    Case Else
                        Throw New NotImplementedException("Invalid bit depth.")
                End Select
            End Set
        End Property

        ''' <summary>
        ''' The color mode of the file.
        ''' </summary>
        Public Property ColorMode As PsdColorMode

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub LoadHeader(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, File header")

            Dim signature = reader.ReadAsciiChars(4)
            If signature <> "8BPS" Then
                Throw New PsdInvalidException("The given stream is not a valid PSD file")
            End If

            Version = CType(reader.ReadInt16(), PsdFileVersion)
            Util.DebugMessage(reader.BaseStream, $"Load, Info, Version {CInt(Version)}")
            If (Version <> PsdFileVersion.Psd) AndAlso (Version <> PsdFileVersion.PsbLargeDocument) Then
                Throw New PsdInvalidException("The PSD file has an unknown version")
            End If

            ' 6 bytes reserved
            reader.BaseStream.Position += 6

            Me.ChannelCount = reader.ReadInt16()
            Me.RowCount = reader.ReadInt32()
            Me.ColumnCount = reader.ReadInt32()
            BitDepth = reader.ReadInt16()
            ColorMode = CType(reader.ReadInt16(), PsdColorMode)

            Util.DebugMessage(reader.BaseStream, "Load, End, File header")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SaveHeader(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, File header")

            Dim signature As String = "8BPS"
            writer.WriteAsciiChars(signature)
            writer.Write(CShort(Version))
            writer.Write(New Byte() {&H0, &H0, &H0, &H0, &H0, &H0})
            writer.Write(ChannelCount)
            writer.Write(RowCount)
            writer.Write(ColumnCount)
            writer.Write(CShort(BitDepth))
            writer.Write(CShort(ColorMode))

            Util.DebugMessage(writer.BaseStream, "Save, End, File header")
        End Sub

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "ColorModeData"

        ''' <summary>
        ''' If ColorMode is ColorModes.Indexed, the following 768 bytes will contain
        ''' a 256-color palette. If the ColorMode is ColorModes.Duotone, the data
        ''' following presumably consists of screen parameters and other related
        ''' information. Unfortunately, it is intentionally not documented by Adobe,
        ''' and non-Photoshop readers are advised to treat duotone images as
        ''' gray-scale images.
        ''' </summary>
        Public ColorModeData As Byte() = New Byte() {}

        Private Sub LoadColorModeData(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, ColorModeData")

            Dim paletteLength = reader.ReadUInt32()
            If paletteLength > 0 Then
                ColorModeData = reader.ReadBytes(CInt(paletteLength))
            End If

            Util.DebugMessage(reader.BaseStream, "Load, End, ColorModeData")
        End Sub

        Private Sub SaveColorModeData(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, ColorModeData")

            writer.Write(CUInt(ColorModeData.Length))
            writer.Write(ColorModeData)

            Util.DebugMessage(writer.BaseStream, "Save, End, ColorModeData")
        End Sub

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "ImageResources"

        ''' <summary>
        ''' The Image resource blocks for the file
        ''' </summary>
        Public Property ImageResources As ImageResources

        Public Property Resolution As ResolutionInfo
            Get
                Return CType(ImageResources.Get(ResourceID.ResolutionInfo), ResolutionInfo)
            End Get
            Set(value As ResolutionInfo)
                ImageResources.Set(value)
            End Set
        End Property

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub LoadImageResources(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, ImageResources")

            Dim imageResourcesLength = reader.ReadUInt32()
            If imageResourcesLength <= 0 Then
                Return
            End If

            Dim startPosition = reader.BaseStream.Position
            Dim endPosition = startPosition + imageResourcesLength
            While reader.BaseStream.Position < endPosition
                Dim imageResource = ImageResourceFactory.CreateImageResource(reader)
                ImageResources.Add(imageResource)
            End While

            Util.DebugMessage(reader.BaseStream, "Load, End, ImageResources")

            '-----------------------------------------------------------------------
            ' make sure we are not on a wrong offset, so set the stream position
            ' manually
            reader.BaseStream.Position = startPosition + imageResourcesLength
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SaveImageResources(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, ImageResources")

            Using New PsdBlockLengthWriter(writer)
                For Each imgRes In ImageResources
                    imgRes.Save(writer)
                Next
            End Using

            Util.DebugMessage(writer.BaseStream, "Save, End, ImageResources")
        End Sub

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "LayerAndMaskInfo"

        Private _layers As List(Of Layer)
        Public Property Layers As List(Of Layer)
            Get
                Return _layers
            End Get
            Private Set(value As List(Of Layer))
                _layers = value
            End Set
        End Property

        Private _additionalInfo As List(Of LayerInfo)
        Public Property AdditionalInfo As List(Of LayerInfo)
            Get
                Return _additionalInfo
            End Get
            Private Set(value As List(Of LayerInfo))
                _additionalInfo = value
            End Set
        End Property

        Public Property AbsoluteAlpha As Boolean

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub LoadLayerAndMaskInfo(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, Layer and mask info")

            Dim layersAndMaskLength As Long = If(IsLargeDocument, reader.ReadInt64(), CLng(reader.ReadUInt32()))
            If layersAndMaskLength <= 0 Then
                Return
            End If

            Dim startPosition = reader.BaseStream.Position
            Dim endPosition = startPosition + layersAndMaskLength

            LoadLayers(reader, True)
            LoadGlobalLayerMask(reader, endPosition)
            LayerInfoFactory.LoadAll(reader, Me, AdditionalInfo, endPosition, True)

            For Each layerInfo In AdditionalInfo
                Select Case layerInfo.Key
                    Case "LMsk"
                        GlobalLayerMaskData = CType(layerInfo, RawLayerInfo).Data
                End Select
            Next

            Util.DebugMessage(reader.BaseStream, "Load, End, Layer and mask info")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SaveLayerAndMaskInfo(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Layer and mask info")

            Using New PsdBlockLengthWriter(writer, IsLargeDocument)
                Dim startPosition = writer.BaseStream.Position

                SaveLayers(writer)
                SaveGlobalLayerMask(writer)

                For Each info In AdditionalInfo
                    info.Save(writer, globalLayerInfo:=True, isLargeDocument:=IsLargeDocument)
                Next

                writer.WritePadding(startPosition, 2)
            End Using

            Util.DebugMessage(writer.BaseStream, "Save, End, Layer and mask info")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Load Layers Info section, including image data.
        ''' </summary>
        ''' <param name="reader">PSD reader.</param>
        ''' <param name="hasHeader">Whether the Layers Info section has a length header.</param>
        Friend Sub LoadLayers(reader As PsdBinaryReader, hasHeader As Boolean)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, Layers Info section")

            Dim sectionLength As Long = 0
            If hasHeader Then
                sectionLength = If(IsLargeDocument, reader.ReadInt64(), CLng(reader.ReadUInt32()))

                If sectionLength <= 0 Then
                    ' The callback may take action when there are 0 layers, so it must
                    ' be called even though the Layers Info section is empty.
                    LoadContext.OnLoadLayersHeader(Me)
                    Util.DebugMessage(reader.BaseStream, "Load, End, Layers Info section")
                    Return
                End If
            End If

            Dim startPosition = reader.BaseStream.Position
            Dim numLayers As Integer = reader.ReadInt16()

            ' If numLayers < 0, then number of layers is absolute value,
            ' and the first alpha channel contains the transparency data for
            ' the merged result.
            If numLayers < 0 Then
                AbsoluteAlpha = True
                numLayers = Math.Abs(numLayers)
            End If

            For i As Integer = 0 To numLayers - 1
                Dim layer = New Layer(reader, Me)
                Layers.Add(layer)
            Next

            ' Header is complete just before loading pixel data
            LoadContext.OnLoadLayersHeader(Me)

            '-----------------------------------------------------------------------

            ' Load image data for all channels.
            For Each layer In Layers
                Util.DebugMessage(reader.BaseStream, $"Load, Begin, Layer image, {layer.Name}")
                For Each channel In layer.Channels
                    channel.LoadPixelData(reader)
                Next
                Util.DebugMessage(reader.BaseStream, $"Load, End, Layer image, {layer.Name}")
            Next

            ' Length is set to 0 when called on higher bitdepth layers.
            If sectionLength > 0 Then
                ' Layers Info section is documented to be even-padded, but Photoshop
                ' actually pads to 4 bytes.
                Dim endPosition = startPosition + sectionLength
                Dim positionOffset = reader.BaseStream.Position - endPosition
                Debug.Assert(positionOffset > -4,
                    "LoadLayers did not read the full length of the Layers Info section.")
                Debug.Assert(positionOffset <= 0,
                    "LoadLayers read past the end of the Layers Info section.")

                If reader.BaseStream.Position < endPosition Then
                    reader.BaseStream.Position = endPosition
                End If
            End If

            Util.DebugMessage(reader.BaseStream, "Load, End, Layers")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Decompress the document image data and all the layers' image data, in parallel.
        ''' </summary>
        Private Sub DecompressImages()
            Dim layersAndComposite = Layers.Concat(New Layer() {BaseLayer})
            Dim channels = layersAndComposite.SelectMany(Function(x) x.Channels)
            Parallel.ForEach(channels, Sub(channel) channel.DecodeImageData())

            For Each layer In Layers
                For Each channel In layer.Channels
                    If channel.ID = -2 Then
                        layer.Masks.LayerMask.ImageData = channel.ImageData
                    ElseIf channel.ID = -3 Then
                        layer.Masks.UserMask.ImageData = channel.ImageData
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' Check the validity of the PSD file and generate necessary data.
        ''' </summary>
        Public Sub PrepareSave()
            CheckDimension(ColumnCount)
            CheckDimension(RowCount)
            VerifyInfoLayers()
            VerifyLayerSections()

            Dim imageLayers = Layers.Concat(New List(Of Layer) From {Me.BaseLayer}).ToList()

            For Each layer In imageLayers
                layer.PrepareSave()
            Next

            SetVersionInfo()
        End Sub

        ''' <summary>
        ''' Verifies that any Additional Info layers are consistent.
        ''' </summary>
        Private Sub VerifyInfoLayers()
            Dim infoLayersCount = AdditionalInfo.Where(Function(x) TypeOf x Is InfoLayers).Count()
            If infoLayersCount > 1 Then
                Throw New PsdInvalidException($"Cannot have more than one {NameOf(InfoLayers)} in a PSD file.")
            End If
            If (infoLayersCount > 0) AndAlso (Layers.Count = 0) Then
                Throw New PsdInvalidException($"{NameOf(InfoLayers)} cannot exist when there are 0 layers.")
            End If
        End Sub

        ''' <summary>
        ''' Verify validity of layer sections. Each start marker should have a
        ''' matching end marker.
        ''' </summary>
        Friend Sub VerifyLayerSections()
            Dim depth As Integer = 0
            For Each layer In Enumerable.Reverse(Layers)
                Dim layerSectionInfo = layer.AdditionalInfo.SingleOrDefault(Function(x) TypeOf x Is LayerSectionInfo)
                If layerSectionInfo Is Nothing Then
                    Continue For
                End If

                Dim sectionInfo = CType(layerSectionInfo, LayerSectionInfo)
                Select Case sectionInfo.SectionType
                    Case LayerSectionType.Layer
                        ' No-op

                    Case LayerSectionType.OpenFolder, LayerSectionType.ClosedFolder
                        depth += 1

                    Case LayerSectionType.SectionDivider
                        depth -= 1
                        If depth < 0 Then
                            Throw New PsdInvalidException("Layer section ended without matching start marker.")
                        End If

                    Case Else
                        Throw New PsdInvalidException("Unrecognized layer section type.")
                End Select
            Next

            If depth <> 0 Then
                Throw New PsdInvalidException("Layer section not closed by end marker.")
            End If
        End Sub

        ''' <summary>
        ''' Set the VersionInfo resource on the file.
        ''' </summary>
        Public Sub SetVersionInfo()
            Dim versionInfo = CType(ImageResources.Get(ResourceID.VersionInfo), VersionInfo)
            If versionInfo Is Nothing Then
                versionInfo = New VersionInfo()
                ImageResources.Set(versionInfo)

                ' Get the version string. We don't use the fourth part (revision).
                Dim assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Dim version = assembly.GetName().Version
                Dim versionString = version.Major & "." & version.Minor & "." & version.Build

                ' Strings are not localized since they are not shown to the user.
                versionInfo.Version = 1
                versionInfo.HasRealMergedData = True
                versionInfo.ReaderName = "Paint.NET PSD Plugin"
                versionInfo.WriterName = "Paint.NET PSD Plugin " & versionString
                versionInfo.FileVersion = 1
            End If
        End Sub

        ''' <summary>
        ''' Saves the Layers Info section, including headers and padding.
        ''' </summary>
        ''' <param name="writer">The PSD writer.</param>
        Friend Sub SaveLayers(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Layers Info section")

            Using New PsdBlockLengthWriter(writer, IsLargeDocument)
                Dim startPosition = writer.BaseStream.Position

                ' Only one set of Layers can exist in the file. If layers will be
                ' written to the Additional Info section, then the Layers section
                ' must be empty to avoid conflict.
                Dim hasInfoLayers = AdditionalInfo.Exists(Function(x) TypeOf x Is InfoLayers)
                If Not hasInfoLayers Then
                    SaveLayersData(writer)
                End If

                ' Documentation states that the Layers Info section is even-padded,
                ' but it is actually padded to a multiple of 4.
                writer.WritePadding(startPosition, 4)
            End Using

            Util.DebugMessage(writer.BaseStream, "Save, End, Layers Info section")
        End Sub

        ''' <summary>
        ''' Saves the layer data, excluding headers and padding.
        ''' </summary>
        ''' <param name="writer">The PSD writer.</param>
        Friend Sub SaveLayersData(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Layers")

            Dim numLayers As Int16 = CShort(Layers.Count)
            If AbsoluteAlpha Then
                numLayers = CShort(-numLayers)
            End If

            ' Photoshop will not load files that have a layer count of 0 in the
            ' compatible Layers section. Instead, the Layers section must be
            ' entirely empty.
            If numLayers = 0 Then
                Return
            End If

            writer.Write(numLayers)

            For Each layer In Layers
                layer.Save(writer)
            Next

            For Each layer In Layers
                Util.DebugMessage(writer.BaseStream, $"Save, Begin, Layer image, {layer.Name}")
                For Each channel In layer.Channels
                    channel.SavePixelData(writer)
                Next
                Util.DebugMessage(writer.BaseStream, $"Save, End, Layer image, {layer.Name}")
            Next

            ' The caller is responsible for padding. Photoshop writes padded
            ' lengths for compatible layers, but unpadded lengths for Additional
            ' Info layers.

            Util.DebugMessage(writer.BaseStream, "Save, End, Layers")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private GlobalLayerMaskData As Byte() = New Byte() {}

        Private Sub LoadGlobalLayerMask(reader As PsdBinaryReader, endPosition As Long)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, GlobalLayerMask")

            If endPosition - reader.BaseStream.Position >= 4 Then
                Dim maskLength = reader.ReadUInt32()

                If maskLength > 0 Then
                    GlobalLayerMaskData = reader.ReadBytes(CInt(maskLength))
                End If
            End If

            Util.DebugMessage(reader.BaseStream, "Load, End, GlobalLayerMask")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SaveGlobalLayerMask(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, GlobalLayerMask")

            If AdditionalInfo.Exists(Function(x) x.Key = "LMsk") Then
                writer.Write(CUInt(0))
            Else
                writer.Write(CUInt(GlobalLayerMaskData.Length))
                writer.Write(GlobalLayerMaskData)
            End If

            Util.DebugMessage(writer.BaseStream, "Save, End, GlobalLayerMask")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "Composite image"

        ''' <summary>
        ''' Represents the composite image.
        ''' </summary>
        Public Property BaseLayer As Layer

        Public Property ImageCompression As ImageCompression

        Private Sub LoadImage(reader As PsdBinaryReader)
            Util.DebugMessage(reader.BaseStream, "Load, Begin, Composite image")

            ImageCompression = CType(reader.ReadInt16(), ImageCompression)

            ' Create channels
            For i As Int16 = 0 To CShort(ChannelCount - 1)
                Util.DebugMessage(reader.BaseStream, "Load, Begin, Channel image data")

                Dim channel = New Channel(i, Me.BaseLayer)
                channel.ImageCompression = ImageCompression
                channel.Length = CLng(Me.RowCount) * Util.BytesPerRow(BaseLayer.Rect.Size, BitDepth)

                ' The composite image stores all RLE headers up-front, rather than
                ' with each channel.
                If ImageCompression = PhotoshopFile.ImageCompression.Rle Then
                    channel.RleRowLengths = New RleRowLengths(reader, RowCount, IsLargeDocument)
                    channel.Length = channel.RleRowLengths.Total
                End If

                BaseLayer.Channels.Add(channel)
                Util.DebugMessage(reader.BaseStream, "Load, End, Channel image data")
            Next

            For Each channel In Me.BaseLayer.Channels
                Util.DebugMessage(reader.BaseStream, "Load, Begin, Channel image data")
                Util.CheckByteArrayLength(channel.Length)
                channel.ImageDataRaw = reader.ReadBytes(CInt(channel.Length))
                Util.DebugMessage(reader.BaseStream, "Load, End, Channel image data")
            Next

            ' If there is exactly one more channel than we need, then it is the
            ' alpha channel.
            If (ColorMode <> PsdColorMode.Multichannel) AndAlso (ChannelCount = ColorMode.MinChannelCount() + 1) Then
                Dim alphaChannel = BaseLayer.Channels.Last()
                alphaChannel.ID = -1
            End If

            Util.DebugMessage(reader.BaseStream, "Load, End, Composite image")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SaveImage(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, Composite image")

            writer.Write(CShort(Me.ImageCompression))
            If Me.ImageCompression = PhotoshopFile.ImageCompression.Rle Then
                For Each channel In Me.BaseLayer.Channels
                    Util.DebugMessage(writer.BaseStream, "Save, Begin, RLE header")
                    channel.RleRowLengths.Write(writer, IsLargeDocument)
                    Util.DebugMessage(writer.BaseStream, "Save, End, RLE header")
                Next
            End If
            For Each channel In Me.BaseLayer.Channels
                Util.DebugMessage(writer.BaseStream, "Save, Begin, Channel image data")
                writer.Write(channel.ImageDataRaw)
                Util.DebugMessage(writer.BaseStream, "Save, End, Channel image data")
            Next

            Util.DebugMessage(writer.BaseStream, "Save, End, Composite image")
        End Sub

#End Region

    End Class

    ''' <summary>
    ''' The possible Compression methods.
    ''' </summary>
    Public Enum ImageCompression
        ''' <summary>
        ''' Raw data
        ''' </summary>
        Raw = 0
        ''' <summary>
        ''' RLE compressed
        ''' </summary>
        Rle = 1
        ''' <summary>
        ''' ZIP without prediction.
        ''' </summary>
        Zip = 2
        ''' <summary>
        ''' ZIP with prediction.
        ''' </summary>
        ZipPrediction = 3
    End Enum

End Namespace
