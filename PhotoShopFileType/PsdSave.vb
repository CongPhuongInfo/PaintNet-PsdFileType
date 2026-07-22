'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PsdSave.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2024 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: the original walks the Paint.NET Surface with raw
' ColorBgra* pointers from Surface.GetRowPointer (unsafe). VB.NET supports
' neither unsafe code nor pointers, so this port instead uses the Surface's
' safe (x, y) indexer property everywhere a pixel needs to be read. Logic
' and arithmetic are otherwise unchanged.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks

Imports PhotoshopFile

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Module PsdSave

        Public Sub Save(input As Document, output As Stream, psdToken As PsdSaveConfigToken,
                         scratchSurface As Surface, progressCallback As ProgressEventHandler)
            Dim psdVersion = If((input.Height > 30000) OrElse (input.Width > 30000),
                PsdFileVersion.PsbLargeDocument,
                PsdFileVersion.Psd)
            Dim psdFile = New PsdFile(psdVersion)

            psdFile.RowCount = input.Height
            psdFile.ColumnCount = input.Width

            ' We only save in RGBA format, 8 bits per channel, which corresponds to
            ' Paint.NET's internal representation.

            psdFile.ChannelCount = 4
            psdFile.ColorMode = PsdColorMode.RGB
            psdFile.BitDepth = 8
            psdFile.Resolution = GetResolutionInfo(input)
            psdFile.ImageCompression = If(psdToken.RleCompress, ImageCompression.Rle, ImageCompression.Raw)

            ' Treat the composite image as another layer when reporting progress.
            Dim progress = New ProgressNotifier(progressCallback)
            Dim percentPerLayer = percentStoreImages / (input.Layers.Count + 1)

            ' Render the composite image. This operation is parallelized within
            ' Paint.NET using its own private thread pool.
            Using ra As New RenderArgs(scratchSurface)
                input.Flatten(scratchSurface)
                progress.Notify(percentRenderComposite)
            End Using

            ' Delegate to store the composite
            Dim storeCompositeAction As Action =
                Sub()
                    ' Allocate space for the composite image data
                    Dim imageSize As Integer = psdFile.RowCount * psdFile.ColumnCount
                    For i As Short = 0 To CShort(psdFile.ChannelCount - 1)
                        Dim channel = New Channel(i, psdFile.BaseLayer)
                        channel.ImageData = New Byte(imageSize - 1) {}
                        channel.ImageCompression = psdFile.ImageCompression
                        psdFile.BaseLayer.Channels.Add(channel)
                    Next

                    Dim channelsArray = psdFile.BaseLayer.Channels.ToIdArray()
                    StoreLayerImage(channelsArray, channelsArray(3), scratchSurface, psdFile.BaseLayer.Rect)

                    progress.Notify(percentPerLayer)
                End Sub

            ' Delegate to store the layers
            Dim storeLayersAction As Action =
                Sub()
                    ' LayerList is an ArrayList, so we have to cast to get a generic
                    ' IEnumerable that works with LINQ.
                    Dim pdnLayers = input.Layers.Cast(Of BitmapLayer)()
                    Dim psdLayers = pdnLayers.AsParallel().AsOrdered().Select(
                        Function(pdnLayer)
                            Dim psdLayer = New PhotoshopFile.Layer(psdFile)
                            StoreLayer(pdnLayer, psdLayer, psdToken)

                            progress.Notify(percentPerLayer)
                            Return psdLayer
                        End Function)
                    psdFile.Layers.AddRange(psdLayers)
                End Sub

            ' Process composite and layers in parallel
            Parallel.Invoke(storeCompositeAction, storeLayersAction)

            psdFile.Save(output, Encoding.Default)
        End Sub

        Private Function GetResolutionInfo(input As Document) As ResolutionInfo
            Dim resInfo = New ResolutionInfo()

            resInfo.HeightDisplayUnit = ResolutionInfo.Unit.Inches
            resInfo.WidthDisplayUnit = ResolutionInfo.Unit.Inches

            If input.DpuUnit = MeasurementUnit.Inch Then
                resInfo.HResDisplayUnit = ResolutionInfo.ResUnit.PxPerInch
                resInfo.VResDisplayUnit = ResolutionInfo.ResUnit.PxPerInch

                resInfo.HDpi = New UFixed16_16(input.DpuX)
                resInfo.VDpi = New UFixed16_16(input.DpuY)
            Else
                resInfo.HResDisplayUnit = ResolutionInfo.ResUnit.PxPerCm
                resInfo.VResDisplayUnit = ResolutionInfo.ResUnit.PxPerCm

                ' Always stored as pixels/inch even if the display unit is
                ' pixels/centimeter.
                resInfo.HDpi = New UFixed16_16(input.DpuX * 2.54)
                resInfo.VDpi = New UFixed16_16(input.DpuY * 2.54)
            End If

            Return resInfo
        End Function

        ''' <summary>
        ''' Determine the real size of the layer, i.e., the smallest rectangle
        ''' that includes all non-transparent pixels.
        ''' </summary>
        Private Function FindImageRectangle(surface As Surface) As Rectangle
            Dim rectPos As New Util.RectanglePosition With {
                .Left = surface.Width,
                .Top = surface.Height,
                .Right = 0,
                .Bottom = 0
            }

            ' Search for top non-transparent pixel
            Dim fPixelFound = False
            For y As Integer = 0 To surface.Height - 1
                If ExpandImageRectangle(surface, y, 0, surface.Width, rectPos) Then
                    fPixelFound = True
                    Exit For
                End If
            Next

            ' Narrow down the other dimensions of the image rectangle
            If fPixelFound Then
                ' Search for bottom non-transparent pixel
                For y As Integer = surface.Height - 1 To rectPos.Bottom + 1 Step -1
                    If ExpandImageRectangle(surface, y, 0, surface.Width, rectPos) Then
                        Exit For
                    End If
                Next

                ' Search for left and right non-transparent pixels. Because we
                ' scan horizontally, we can't just break, but we can examine fewer
                ' candidate pixels on the remaining rows.
                For y As Integer = rectPos.Top + 1 To rectPos.Bottom - 1
                    ExpandImageRectangle(surface, y, 0, rectPos.Left, rectPos)
                    ExpandImageRectangle(surface, y, rectPos.Right + 1, surface.Width, rectPos)
                Next
            Else
                rectPos.Left = 0
                rectPos.Top = 0
            End If

            Debug.Assert(rectPos.Left <= rectPos.Right)
            Debug.Assert(rectPos.Top <= rectPos.Bottom)

            Dim result = New Rectangle(rectPos.Left, rectPos.Top,
                rectPos.Right - rectPos.Left + 1, rectPos.Bottom - rectPos.Top + 1)
            Return result
        End Function

        ''' <summary>
        ''' Check for non-transparent pixels in a row, or portion of a row.
        ''' Expands the size of the image rectangle if any were found.
        ''' </summary>
        ''' <returns>True if non-transparent pixels were found, false otherwise.</returns>
        Private Function ExpandImageRectangle(surface As Surface, y As Integer,
                                               xStart As Integer, xEnd As Integer,
                                               ByRef rectPos As Util.RectanglePosition) As Boolean
            Dim fPixelFound = False

            For x As Integer = xStart To xEnd - 1
                If surface(x, y).A > 0 Then
                    ' Expand the rectangle to include the specified point.
                    If x < rectPos.Left Then
                        rectPos.Left = x
                    End If
                    If x > rectPos.Right Then
                        rectPos.Right = x
                    End If
                    If y < rectPos.Top Then
                        rectPos.Top = y
                    End If
                    If y > rectPos.Bottom Then
                        rectPos.Bottom = y
                    End If
                    fPixelFound = True
                End If
            Next

            Return fPixelFound
        End Function

        ''' <summary>
        ''' Store layer metadata and image data.
        ''' </summary>
        Public Sub StoreLayer(layer As BitmapLayer, psdLayer As PhotoshopFile.Layer, psdToken As PsdSaveConfigToken)
            ' Set layer metadata
            psdLayer.Name = layer.Name
            psdLayer.Rect = FindImageRectangle(layer.Surface)
            psdLayer.BlendModeKey = layer.BlendMode.ToPsdBlendMode()
            psdLayer.Opacity = layer.Opacity
            psdLayer.Visible = layer.Visible
            psdLayer.Masks = New MaskInfo()
            psdLayer.BlendingRangesData = New BlendingRanges(psdLayer)

            ' Store channel metadata
            Dim layerSize As Integer = psdLayer.Rect.Width * psdLayer.Rect.Height
            For i As Integer = -1 To 2
                Dim ch = New Channel(CShort(i), psdLayer)
                ch.ImageCompression = If(psdToken.RleCompress, ImageCompression.Rle, ImageCompression.Raw)
                ch.ImageData = New Byte(layerSize - 1) {}
                psdLayer.Channels.Add(ch)
            Next

            ' Store and compress channel image data
            Dim channelsArray = psdLayer.Channels.ToIdArray()
            StoreLayerImage(channelsArray, psdLayer.AlphaChannel, layer.Surface, psdLayer.Rect)
        End Sub

        ''' <summary>
        ''' Stores and compresses the image data for the layer.
        ''' </summary>
        ''' <param name="channels">Destination channels.</param>
        ''' <param name="alphaChannel">Destination alpha channel.</param>
        ''' <param name="surface">Source image from Paint.NET.</param>
        ''' <param name="rect">Image rectangle to store.</param>
        Private Sub StoreLayerImage(channels As Channel(), alphaChannel As Channel,
                                     surface As Surface, rect As Rectangle)
            For y As Integer = 0 To rect.Height - 1
                Dim destRowIndex = y * rect.Width

                For x As Integer = 0 To rect.Width - 1
                    Dim destIndex = destRowIndex + x
                    Dim srcPixel = surface(rect.Left + x, y + rect.Top)

                    channels(0).ImageData(destIndex) = srcPixel.R
                    channels(1).ImageData(destIndex) = srcPixel.G
                    channels(2).ImageData(destIndex) = srcPixel.B
                    alphaChannel.ImageData(destIndex) = srcPixel.A
                Next
            Next

            Parallel.ForEach(channels, Sub(channel) channel.CompressImageData())
        End Sub

#Region "Progress notification"

        ' We only report progress to 90%, reserving 10% for writing out to disk.
        Private percentRenderComposite As Double = 20.0
        Private percentStoreImages As Double = 70.0

        Private Class ProgressNotifier
            Private callback As ProgressEventHandler
            Private percent As Double

            Friend Sub New(progressCallback As ProgressEventHandler)
                callback = progressCallback
                percent = 0
            End Sub

            Friend Sub Notify(percentIncrement As Double)
                SyncLock Me
                    percent += percentIncrement
                    callback.Invoke(Nothing, New ProgressEventArgs(percent))
                End SyncLock
            End Sub
        End Class

#End Region

    End Module

End Namespace
