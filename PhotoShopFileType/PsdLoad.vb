'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PsdLoad.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2021 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Runtime.CompilerServices

Imports PhotoshopFile

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Module PsdLoad

        Public Function Load(input As Stream) As Document
            ' Load and decompress Photoshop file structures
            Dim loadContext = New DocumentLoadContext()
            Dim psdFile = New PsdFile(input, loadContext)

            ' Multichannel images are loaded by processing each channel as a
            ' grayscale layer.
            If psdFile.ColorMode = PsdColorMode.Multichannel Then
                CreateLayersFromChannels(psdFile)
                psdFile.ColorMode = PsdColorMode.Grayscale
            End If

            ' Convert into Paint.NET internal representation
            Dim document = New Document(psdFile.ColumnCount, psdFile.RowCount)

            If psdFile.Layers.Count = 0 Then
                psdFile.BaseLayer.CreateMissingChannels()
                Dim layer = PaintDotNet.Layer.CreateBackgroundLayer(psdFile.ColumnCount, psdFile.RowCount)
                ImageDecoderPdn.DecodeImage(layer, psdFile.BaseLayer)
                document.Layers.Add(layer)
            Else
                psdFile.VerifyLayerSections()
                ApplyLayerSections(psdFile.Layers)

                Dim pdnLayers = psdFile.Layers.AsParallel().AsOrdered().
                    Select(Function(psdLayer) psdLayer.DecodeToPdnLayer()).
                    ToList()
                document.Layers.AddRange(pdnLayers.Cast(Of PaintDotNet.Layer)())
            End If

            SetPdnResolutionInfo(psdFile, document)

            Return document
        End Function

        <Extension()>
        Friend Function DecodeToPdnLayer(psdLayer As PhotoshopFile.Layer) As BitmapLayer
            Dim psdFile = psdLayer.PsdFile
            psdLayer.CreateMissingChannels()

            Dim pdnLayer = New BitmapLayer(psdFile.ColumnCount, psdFile.RowCount)
            pdnLayer.Name = psdLayer.Name
            pdnLayer.Opacity = psdLayer.Opacity
            pdnLayer.Visible = psdLayer.Visible
            pdnLayer.BlendMode = BlendModeMapping.FromPsdBlendMode(psdLayer.BlendModeKey)
            ImageDecoderPdn.DecodeImage(pdnLayer, psdLayer)

            Return pdnLayer
        End Function

        ''' <summary>
        ''' Creates a layer for each channel in a multichannel image.
        ''' </summary>
        Private Sub CreateLayersFromChannels(psdFile As PsdFile)
            If psdFile.ColorMode <> PsdColorMode.Multichannel Then
                Throw New Exception("Not a multichannel image.")
            End If
            If psdFile.Layers.Count > 0 Then
                Throw New PsdInvalidException("Multichannel image should not have layers.")
            End If

            ' Get alpha channel names, preferably in Unicode.
            Dim alphaChannelNames = CType(psdFile.ImageResources.[Get](ResourceID.AlphaChannelNames), AlphaChannelNames)
            Dim unicodeAlphaNames = CType(psdFile.ImageResources.[Get](ResourceID.UnicodeAlphaNames), UnicodeAlphaNames)
            If (alphaChannelNames Is Nothing) AndAlso (unicodeAlphaNames Is Nothing) Then
                Throw New PsdInvalidException("No channel names found.")
            End If

            Dim channelNames = If(unicodeAlphaNames IsNot Nothing, unicodeAlphaNames.ChannelNames, alphaChannelNames.ChannelNames)
            Dim channels = psdFile.BaseLayer.Channels
            If channels.Count > channelNames.Count Then
                Throw New PsdInvalidException("More channels than channel names.")
            End If

            ' Channels are stored from top to bottom, but layers are stored from
            ' bottom to top.
            Dim channelsNamesReversed = channels.Zip(channelNames, Function(ch, name) Tuple.Create(ch, name)).Reverse()
            For Each pair In channelsNamesReversed
                Dim channel = pair.Item1
                Dim channelName = pair.Item2

                ' Copy metadata over from base layer
                Dim layer = New PhotoshopFile.Layer(psdFile)
                layer.Rect = psdFile.BaseLayer.Rect
                layer.Visible = True
                layer.Masks = New MaskInfo()
                layer.BlendingRangesData = New BlendingRanges(layer)

                ' We do not attempt to reconstruct the appearance of the image, but
                ' only to provide access to the channels image data.
                layer.Name = channelName
                layer.BlendModeKey = PsdBlendMode.Darken
                layer.Opacity = 255

                ' Copy channel image data into the new grayscale layer
                Dim layerChannel = New Channel(0, layer)
                layerChannel.ImageCompression = channel.ImageCompression
                layerChannel.ImageData = channel.ImageData
                layer.Channels.Add(layerChannel)

                psdFile.Layers.Add(layer)
            Next
        End Sub

        ''' <summary>
        ''' Transform Photoshop's layer tree to Paint.NET's flat layer list.
        ''' Indicate where layer sections begin and end, and hide all layers within
        ''' hidden layer sections.
        ''' </summary>
        Private Sub ApplyLayerSections(layers As List(Of PhotoshopFile.Layer))
            ' BUG: PsdPluginResources.GetString will always return English resource,
            ' because Paint.NET does not set the CurrentUICulture when OnLoad is
            ' called. This situation should be resolved with Paint.NET 4.0, which
            ' will provide an alternative mechanism to retrieve the UI language.

            ' Cache layer section strings
            Dim beginSectionWrapper = PsdPluginResources.GetString("LayersPalette_LayerGroupBegin")
            Dim endSectionWrapper = PsdPluginResources.GetString("LayersPalette_LayerGroupEnd")

            ' Track the depth of the topmost hidden section. Any nested sections
            ' will be hidden, whether or not they themselves have the flag set.
            Dim topHiddenSectionDepth As Integer = Integer.MaxValue
            Dim layerSectionNames As New Stack(Of String)()

            ' Layers are stored bottom-to-top, but layer sections are specified
            ' top-to-bottom.
            For Each layer In Enumerable.Reverse(layers)
                ' Apply to all layers within the layer section, as well as the
                ' closing layer.
                If layerSectionNames.Count > topHiddenSectionDepth Then
                    layer.Visible = False
                End If

                Dim sectionInfo = CType(layer.AdditionalInfo.SingleOrDefault(Function(x) TypeOf x Is LayerSectionInfo), LayerSectionInfo)
                If sectionInfo Is Nothing Then
                    Continue For
                End If

                Select Case sectionInfo.SectionType
                    Case LayerSectionType.OpenFolder, LayerSectionType.ClosedFolder
                        ' Start a new layer section
                        If (Not layer.Visible) AndAlso (topHiddenSectionDepth = Integer.MaxValue) Then
                            topHiddenSectionDepth = layerSectionNames.Count
                        End If
                        layerSectionNames.Push(layer.Name)
                        layer.Name = String.Format(beginSectionWrapper, layer.Name)

                    Case LayerSectionType.SectionDivider
                        ' End the current layer section
                        Dim layerSectionName = layerSectionNames.Pop()
                        If layerSectionNames.Count = topHiddenSectionDepth Then
                            topHiddenSectionDepth = Integer.MaxValue
                        End If
                        layer.Name = String.Format(endSectionWrapper, layerSectionName)
                End Select
            Next
        End Sub

        ''' <summary>
        ''' Set the resolution on the Paint.NET Document to match the PSD file.
        ''' </summary>
        Private Sub SetPdnResolutionInfo(psdFile As PsdFile, document As Document)
            If psdFile.Resolution IsNot Nothing Then
                ' PSD files always specify the resolution in DPI. When loading and
                ' saving cm, we will have to round-trip the conversion, but doubles
                ' have plenty of precision to spare vs. PSD's 16/16 fixed-point.

                If (psdFile.Resolution.HResDisplayUnit = ResolutionInfo.ResUnit.PxPerCm) AndAlso
                   (psdFile.Resolution.VResDisplayUnit = ResolutionInfo.ResUnit.PxPerCm) Then
                    document.DpuUnit = MeasurementUnit.Centimeter

                    ' HACK: Paint.NET truncates DpuX and DpuY to three decimal places,
                    ' so add 0.0005 to get a rounded value instead.
                    document.DpuX = CDbl(psdFile.Resolution.HDpi) / 2.54 + 0.0005
                    document.DpuY = CDbl(psdFile.Resolution.VDpi) / 2.54 + 0.0005
                Else
                    document.DpuUnit = MeasurementUnit.Inch
                    document.DpuX = CDbl(psdFile.Resolution.HDpi)
                    document.DpuY = CDbl(psdFile.Resolution.VDpi)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Verify that the PSD file will fit into physical memory once loaded
        ''' and converted to Paint.NET format.
        ''' </summary>
        ''' <remarks>
        ''' This check is necessary because layers in Paint.NET have the same
        ''' dimensions as the canvas. Thus, PSD files that contain lots of
        ''' tiny adjustment layers may blow up in size by several
        ''' orders of magnitude.
        ''' </remarks>
        Friend Sub CheckSufficientMemory(psdFile As PsdFile)
            ' Multichannel images have channels converted to layers
            Dim numLayers = If(psdFile.ColorMode = PsdColorMode.Multichannel,
                psdFile.BaseLayer.Channels.Count,
                Math.Max(psdFile.Layers.Count, 1))

            ' Paint.NET also requires a scratch layer and composite layer
            numLayers += 2

            Dim numPixels As Long = CLng(psdFile.ColumnCount) * psdFile.RowCount
            Dim bytesRequired As ULong = CULng(4L * numPixels * numLayers)

            ' Check that the file will fit entirely into physical memory, so that we
            ' do not thrash and make the Paint.NET UI nonresponsive. We also have
            ' to check against virtual memory address space because 32-bit processes
            ' cannot access all 4 GB.
            Dim computerInfo = New Microsoft.VisualBasic.Devices.ComputerInfo()
            Dim accessibleMemory = Math.Min(computerInfo.TotalPhysicalMemory, computerInfo.TotalVirtualMemory)
            If bytesRequired > accessibleMemory Then
                Throw New OutOfMemoryException()
            End If
        End Sub

    End Module

End Namespace
