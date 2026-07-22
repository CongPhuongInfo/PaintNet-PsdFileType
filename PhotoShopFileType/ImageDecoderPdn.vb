'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/ImageDecoderPdn.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2024 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: the original decodes each row of pixels by walking raw
' ColorBgra*/byte* pointers (fixed/unsafe). VB.NET supports neither unsafe
' code nor pointers, so this port is restructured as follows:
'   - Each row is decoded into a plain managed ColorBgra() array (a struct
'     array, so element fields like destRow(i).R remain directly writable,
'     exactly like the original pDest->R writes).
'   - Once a row is fully decoded (color + alpha + masks), it is copied back
'     to the Paint.NET Surface once, using the Surface's safe (x, y) indexer
'     instead of GetRowPointer/pointer arithmetic.
'   - Anywhere the original reinterpreted a byte* as a float* (HDR channels),
'     this port uses BitConverter.ToSingle(byteArray, offset) instead.
' Decode logic and arithmetic are otherwise unchanged.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Drawing

Imports PhotoshopFile

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Module ImageDecoderPdn

        Private Class DecodeContext
            Private _layer As PhotoshopFile.Layer
            Public Property Layer As PhotoshopFile.Layer
                Get
                    Return _layer
                End Get
                Private Set(value As PhotoshopFile.Layer)
                    _layer = value
                End Set
            End Property

            Private _byteDepth As Integer
            Public Property ByteDepth As Integer
                Get
                    Return _byteDepth
                End Get
                Private Set(value As Integer)
                    _byteDepth = value
                End Set
            End Property

            Private _channels As Channel()
            Public Property Channels As Channel()
                Get
                    Return _channels
                End Get
                Private Set(value As Channel())
                    _channels = value
                End Set
            End Property

            Private _alphaChannel As Channel
            Public Property AlphaChannel As Channel
                Get
                    Return _alphaChannel
                End Get
                Private Set(value As Channel)
                    _alphaChannel = value
                End Set
            End Property

            Private _colorMode As PsdColorMode
            Public Property ColorMode As PsdColorMode
                Get
                    Return _colorMode
                End Get
                Private Set(value As PsdColorMode)
                    _colorMode = value
                End Set
            End Property

            Private _colorModeData As Byte()
            Public Property ColorModeData As Byte()
                Get
                    Return _colorModeData
                End Get
                Private Set(value As Byte())
                    _colorModeData = value
                End Set
            End Property

            Private _rectangle As Rectangle
            Public Property Rectangle As Rectangle
                Get
                    Return _rectangle
                End Get
                Private Set(value As Rectangle)
                    _rectangle = value
                End Set
            End Property

            Private _layerMaskContext As MaskDecodeContext
            Public Property LayerMaskContext As MaskDecodeContext
                Get
                    Return _layerMaskContext
                End Get
                Private Set(value As MaskDecodeContext)
                    _layerMaskContext = value
                End Set
            End Property

            Private _userMaskContext As MaskDecodeContext
            Public Property UserMaskContext As MaskDecodeContext
                Get
                    Return _userMaskContext
                End Get
                Private Set(value As MaskDecodeContext)
                    _userMaskContext = value
                End Set
            End Property

            Public Sub New(layer As PhotoshopFile.Layer, bounds As Rectangle)
                Me.Layer = layer
                ByteDepth = Util.BytesFromBitDepth(layer.PsdFile.BitDepth)
                Channels = layer.Channels.ToIdArray()
                AlphaChannel = layer.AlphaChannel
                ColorMode = layer.PsdFile.ColorMode
                ColorModeData = layer.PsdFile.ColorModeData

                ' Clip the layer to the specified bounds
                Rectangle = Layer.Rect.IntersectWith(bounds)

                If layer.Masks IsNot Nothing Then
                    LayerMaskContext = GetMaskContext(layer.Masks.LayerMask)
                    UserMaskContext = GetMaskContext(layer.Masks.UserMask)
                End If
            End Sub

            Private Function GetMaskContext(mask As Mask) As MaskDecodeContext
                If (mask Is Nothing) OrElse mask.Disabled Then
                    Return Nothing
                End If

                Return New MaskDecodeContext(mask, Me)
            End Function
        End Class

        Private Class MaskDecodeContext
            Private _mask As Mask
            Public Property Mask As Mask
                Get
                    Return _mask
                End Get
                Private Set(value As Mask)
                    _mask = value
                End Set
            End Property

            Private _rectangle As Rectangle
            Public Property Rectangle As Rectangle
                Get
                    Return _rectangle
                End Get
                Private Set(value As Rectangle)
                    _rectangle = value
                End Set
            End Property

            Private _alphaBuffer As Byte()
            Public Property AlphaBuffer As Byte()
                Get
                    Return _alphaBuffer
                End Get
                Private Set(value As Byte())
                    _alphaBuffer = value
                End Set
            End Property

            Public Sub New(mask As Mask, layerContext As DecodeContext)
                Me.Mask = mask

                ' The PositionVsLayer flag is documented to indicate a position
                ' relative to the layer, but Photoshop treats the position as
                ' absolute. So that's what we do, too.
                Rectangle = mask.Rect.IntersectWith(layerContext.Rectangle)
                AlphaBuffer = New Byte(layerContext.Rectangle.Width - 1) {}
            End Sub

            Public Function IsRowEmpty(row As Integer) As Boolean
                Return (Mask.ImageData Is Nothing) OrElse
                    (Mask.ImageData.Length = 0) OrElse
                    (Rectangle.Size.IsEmpty) OrElse
                    (row < Rectangle.Top) OrElse
                    (row >= Rectangle.Bottom)
            End Function
        End Class

        ''' <summary>
        ''' Decode image from Photoshop's channel-separated formats to BGRA.
        ''' </summary>
        Public Sub DecodeImage(pdnLayer As BitmapLayer, psdLayer As PhotoshopFile.Layer)
            Dim decodeContext = New DecodeContext(psdLayer, pdnLayer.Bounds)
            Dim decoder As DecodeDelegate

            If decodeContext.ByteDepth = 4 Then
                decoder = GetDecodeDelegate32(decodeContext.ColorMode)
            Else
                decoder = GetDecodeDelegate(decodeContext.ColorMode)
            End If

            DecodeImage(pdnLayer, decodeContext, decoder)
        End Sub

        Private Delegate Sub DecodeDelegate(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)

        Private Function GetDecodeDelegate(psdColorMode As PsdColorMode) As DecodeDelegate
            Select Case psdColorMode
                Case PsdColorMode.Bitmap
                    Return AddressOf SetPDNRowBitmap
                Case PsdColorMode.Grayscale, PsdColorMode.Duotone
                    Return AddressOf SetPDNRowGrayscale
                Case PsdColorMode.Indexed
                    Return AddressOf SetPDNRowIndexed
                Case PsdColorMode.RGB
                    Return AddressOf SetPDNRowRgb
                Case PsdColorMode.CMYK
                    Return AddressOf SetPDNRowCmyk
                Case PsdColorMode.Lab
                    Return AddressOf SetPDNRowLab
                Case PsdColorMode.Multichannel
                    Throw New Exception("Cannot decode multichannel.")
                Case Else
                    Throw New Exception("Unknown color mode.")
            End Select
        End Function

        Private Function GetDecodeDelegate32(psdColorMode As PsdColorMode) As DecodeDelegate
            Select Case psdColorMode
                Case PsdColorMode.Grayscale
                    Return AddressOf SetPDNRowGrayscale32
                Case PsdColorMode.RGB
                    Return AddressOf SetPDNRowRgb32
                Case Else
                    Throw New PsdInvalidException("32-bit HDR images must be either RGB or grayscale.")
            End Select
        End Function

        ''' <summary>
        ''' Decode image from Photoshop's channel-separated formats to BGRA,
        ''' using the specified decode delegate on each row.
        ''' </summary>
        Private Sub DecodeImage(pdnLayer As BitmapLayer, decodeContext As DecodeContext, decoder As DecodeDelegate)
            Dim psdLayer = decodeContext.Layer
            Dim surface = pdnLayer.Surface
            Dim rect = decodeContext.Rectangle
            Dim width = rect.Right - rect.Left

            ' Convert rows from the Photoshop representation, writing the
            ' resulting ARGB values to the Paint.NET Surface.

            For y As Integer = rect.Top To rect.Bottom - 1
                ' Calculate index into ImageData source from row and column.
                Dim idxSrcPixel = (y - psdLayer.Rect.Top) * psdLayer.Rect.Width + (rect.Left - psdLayer.Rect.Left)
                Dim idxSrcByte = idxSrcPixel * decodeContext.ByteDepth

                ' For 16-bit images, take the higher-order byte from the image
                ' data, which is now in little-endian order.
                If decodeContext.ByteDepth = 2 Then
                    idxSrcByte += 1
                End If

                ' Decode the color and alpha channels into a row buffer.
                Dim destRow(width - 1) As ColorBgra
                decoder(destRow, idxSrcByte, decodeContext)
                SetPDNAlphaRow(destRow, idxSrcByte, decodeContext.ByteDepth, decodeContext.AlphaChannel)

                ' Apply layer masks(s) to the alpha channel
                Dim layerMaskAlphaRow = GetMaskAlphaRow(y, decodeContext, decodeContext.LayerMaskContext)
                Dim userMaskAlphaRow = GetMaskAlphaRow(y, decodeContext, decodeContext.UserMaskContext)
                ApplyPDNMask(destRow, layerMaskAlphaRow, userMaskAlphaRow)

                ' Copy the decoded row back into the Surface.
                For x As Integer = 0 To width - 1
                    surface(rect.Left + x, y) = destRow(x)
                Next
            Next
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SetPDNAlphaRow(destRow As ColorBgra(), idxSrc As Integer, byteDepth As Integer, alphaChannel As Channel)
            ' Set alpha to fully-opaque if there is no alpha channel
            If alphaChannel Is Nothing Then
                For i As Integer = 0 To destRow.Length - 1
                    destRow(i).A = 255
                Next
            Else
                ' Set the alpha channel data
                Dim srcAlpha = idxSrc
                For i As Integer = 0 To destRow.Length - 1
                    destRow(i).A = If(byteDepth < 4,
                        alphaChannel.ImageData(srcAlpha),
                        RGBByteFromHDRFloat(alphaChannel.ImageData, srcAlpha))
                    srcAlpha += byteDepth
                Next
            End If
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Gets one row of alpha values from the mask.
        ''' </summary>
        ''' <param name="y">The y-coordinate of the row.</param>
        ''' <param name="layerContext">The decode context for the layer containing
        '''   the mask.</param>
        ''' <param name="maskContext">The decode context for the mask.</param>
        ''' <returns>An array of alpha values for the row, corresponding to the
        '''   width of the layer decode context.</returns>
        Private Function GetMaskAlphaRow(y As Integer, layerContext As DecodeContext, maskContext As MaskDecodeContext) As Byte()
            If maskContext Is Nothing Then
                Return Nothing
            End If
            Dim mask = maskContext.Mask

            ' Background color for areas not covered by the mask
            Dim backgroundColor As Byte = If(mask.InvertOnBlend, CByte(255 - mask.BackgroundColor), mask.BackgroundColor)
            Util.Fill(maskContext.AlphaBuffer, 0, maskContext.AlphaBuffer.Length, backgroundColor)

            If maskContext.IsRowEmpty(y) Then
                Return maskContext.AlphaBuffer
            End If

            '''''''''''''''''''''''''''''''''''''
            ' Transfer mask into the alpha array
            Dim alphaColumn = maskContext.Rectangle.X - layerContext.Rectangle.X
            Dim alphaStart = alphaColumn
            Dim alphaEnd = alphaColumn + maskContext.Rectangle.Width

            Dim maskRow = y - maskContext.Mask.Rect.Y
            Dim maskColumn = maskContext.Rectangle.X - maskContext.Mask.Rect.X
            Dim idxMaskPixel = (maskRow * mask.Rect.Width) + maskColumn
            Dim maskStart = idxMaskPixel * layerContext.ByteDepth

            ' Take the high-order byte if values are 16-bit (little-endian)
            If layerContext.ByteDepth = 2 Then
                maskStart += 1
            End If

            ' Decode mask into the alpha array.
            If layerContext.ByteDepth = 4 Then
                DecodeMaskAlphaRow32(maskContext.AlphaBuffer, alphaStart, alphaEnd, mask.ImageData, maskStart)
            Else
                DecodeMaskAlphaRow(maskContext.AlphaBuffer, alphaStart, alphaEnd, mask.ImageData, maskStart, layerContext.ByteDepth)
            End If

            ' Obsolete since Photoshop CS6, but retained for compatibility with
            ' older versions. Note that the background has already been inverted.
            If mask.InvertOnBlend Then
                Util.Invert(maskContext.AlphaBuffer, alphaStart, alphaEnd - alphaStart)
            End If

            Return maskContext.AlphaBuffer
        End Function

        Private Sub DecodeMaskAlphaRow32(alphaBuffer As Byte(), alphaStart As Integer, alphaEnd As Integer,
                                          maskData As Byte(), maskStart As Integer)
            Dim maskIdx = maskStart
            For i As Integer = alphaStart To alphaEnd - 1
                alphaBuffer(i) = RGBByteFromHDRFloat(maskData, maskIdx)
                maskIdx += 4
            Next
        End Sub

        Private Sub DecodeMaskAlphaRow(alphaBuffer As Byte(), alphaStart As Integer, alphaEnd As Integer,
                                        maskData As Byte(), maskStart As Integer, byteDepth As Integer)
            Dim maskIdx = maskStart
            For i As Integer = alphaStart To alphaEnd - 1
                alphaBuffer(i) = maskData(maskIdx)
                maskIdx += byteDepth
            Next
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub ApplyPDNMask(destRow As ColorBgra(), layerMaskAlpha As Byte(), userMaskAlpha As Byte())
            ' Do nothing if there are no masks
            If (layerMaskAlpha Is Nothing) AndAlso (userMaskAlpha Is Nothing) Then
                Return

                ' Apply one mask
            ElseIf (layerMaskAlpha Is Nothing) OrElse (userMaskAlpha Is Nothing) Then
                Dim maskAlpha = If(layerMaskAlpha, userMaskAlpha)
                For i As Integer = 0 To destRow.Length - 1
                    destRow(i).A = CByte(CInt(destRow(i).A) * CInt(maskAlpha(i)) \ 255)
                Next

                ' Apply both masks in one pass, to minimize rounding error
            Else
                For i As Integer = 0 To destRow.Length - 1
                    Dim alphaFactor = CInt(layerMaskAlpha(i)) * CInt(userMaskAlpha(i))
                    destRow(i).A = CByte(CInt(destRow(i).A) * alphaFactor \ 65025)
                Next
            End If
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "Decode 32-bit HDR channels"

        Private Sub SetPDNRowRgb32(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim redChannel = context.Channels(0).ImageData
            Dim greenChannel = context.Channels(1).ImageData
            Dim blueChannel = context.Channels(2).ImageData
            Dim idxSrc = idxSrcStart

            For i As Integer = 0 To destRow.Length - 1
                destRow(i).R = RGBByteFromHDRFloat(redChannel, idxSrc)
                destRow(i).G = RGBByteFromHDRFloat(greenChannel, idxSrc)
                destRow(i).B = RGBByteFromHDRFloat(blueChannel, idxSrc)
                idxSrc += 4
            Next
        End Sub

        Private Sub SetPDNRowGrayscale32(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim channelData = context.Channels(0).ImageData
            Dim idxSrc = idxSrcStart

            For i As Integer = 0 To destRow.Length - 1
                Dim rgbValue = RGBByteFromHDRFloat(channelData, idxSrc)
                destRow(i).R = rgbValue
                destRow(i).G = rgbValue
                destRow(i).B = rgbValue
                idxSrc += 4
            Next
        End Sub

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "Decode 8-bit and 16-bit channels"

        Private Sub SetPDNRowRgb(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                destRow(i).R = context.Channels(0).ImageData(idxSrc)
                destRow(i).G = context.Channels(1).ImageData(idxSrc)
                destRow(i).B = context.Channels(2).ImageData(idxSrc)
                idxSrc += context.ByteDepth
            Next
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '
        ' The color-conversion formulas come from the Colour Space Conversions FAQ:
        '     http://www.poynton.com/PDFs/coloureq.pdf
        '
        ' RGB --> CMYK                              CMYK --> RGB
        ' ---------------------------------------   --------------------------------------------
        ' Black   = minimum(1-Red,1-Green,1-Blue)   Red   = 1-minimum(1,Cyan*(1-Black)+Black)
        ' Cyan    = (1-Red-Black)/(1-Black)         Green = 1-minimum(1,Magenta*(1-Black)+Black)
        ' Magenta = (1-Green-Black)/(1-Black)       Blue  = 1-minimum(1,Yellow*(1-Black)+Black)
        ' Yellow  = (1-Blue-Black)/(1-Black)
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private Sub SetPDNRowCmyk(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                ' CMYK values are stored as complements, presumably to allow for some
                ' measure of compatibility with RGB-only applications.
                Dim C = 255 - CInt(context.Channels(0).ImageData(idxSrc))
                Dim M = 255 - CInt(context.Channels(1).ImageData(idxSrc))
                Dim Y = 255 - CInt(context.Channels(2).ImageData(idxSrc))
                Dim K = 255 - CInt(context.Channels(3).ImageData(idxSrc))

                Dim nRed = 255 - Math.Min(255, C * (255 - K) \ 255 + K)
                Dim nGreen = 255 - Math.Min(255, M * (255 - K) \ 255 + K)
                Dim nBlue = 255 - Math.Min(255, Y * (255 - K) \ 255 + K)

                destRow(i).R = CByte(nRed)
                destRow(i).G = CByte(nGreen)
                destRow(i).B = CByte(nBlue)

                idxSrc += context.ByteDepth
            Next
        End Sub

        Private Sub SetPDNRowBitmap(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim bitmap = context.Channels(0).ImageData
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                Dim mask As Byte = CByte(&H80 >> (idxSrc Mod 8))
                Dim bwValue As Byte = CByte(bitmap(idxSrc \ 8) And mask)
                bwValue = If(bwValue = 0, CByte(255), CByte(0))

                destRow(i).R = bwValue
                destRow(i).G = bwValue
                destRow(i).B = bwValue

                idxSrc += context.ByteDepth
            Next
        End Sub

        Private Sub SetPDNRowGrayscale(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                Dim level = context.Channels(0).ImageData(idxSrc)
                destRow(i).R = level
                destRow(i).G = level
                destRow(i).B = level
                idxSrc += context.ByteDepth
            Next
        End Sub

        Private Sub SetPDNRowIndexed(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                Dim index As Integer = CInt(context.Channels(0).ImageData(idxSrc))
                destRow(i).R = context.Layer.PsdFile.ColorModeData(index)
                destRow(i).G = context.Layer.PsdFile.ColorModeData(index + 256)
                destRow(i).B = context.Layer.PsdFile.ColorModeData(index + 2 * 256)
                idxSrc += context.ByteDepth
            Next
        End Sub

        Private Sub SetPDNRowLab(destRow As ColorBgra(), idxSrcStart As Integer, context As DecodeContext)
            Dim idxSrc = idxSrcStart
            For i As Integer = 0 To destRow.Length - 1
                Dim exL, exA, exB As Double
                exL = CDbl(context.Channels(0).ImageData(idxSrc))
                exA = CDbl(context.Channels(1).ImageData(idxSrc))
                exB = CDbl(context.Channels(2).ImageData(idxSrc))

                Dim L As Integer = CInt(Fix(exL / 2.55))
                Dim a As Integer = CInt(Fix(exA - 127.5))
                Dim b As Integer = CInt(Fix(exB - 127.5))

                ' First, convert from Lab to XYZ.
                ' Standards used Observer = 2, Illuminant = D65

                Const ref_X As Double = 95.047
                Const ref_Y As Double = 100.0
                Const ref_Z As Double = 108.883

                Dim var_Y As Double = (CDbl(L) + 16.0) / 116.0
                Dim var_X As Double = CDbl(a) / 500.0 + var_Y
                Dim var_Z As Double = var_Y - CDbl(b) / 200.0

                Dim var_X3 As Double = var_X * var_X * var_X
                Dim var_Y3 As Double = var_Y * var_Y * var_Y
                Dim var_Z3 As Double = var_Z * var_Z * var_Z

                If var_Y3 > 0.008856 Then
                    var_Y = var_Y3
                Else
                    var_Y = (var_Y - 16 / 116) / 7.787
                End If

                If var_X3 > 0.008856 Then
                    var_X = var_X3
                Else
                    var_X = (var_X - 16 / 116) / 7.787
                End If

                If var_Z3 > 0.008856 Then
                    var_Z = var_Z3
                Else
                    var_Z = (var_Z - 16 / 116) / 7.787
                End If

                Dim X As Double = ref_X * var_X
                Dim Y As Double = ref_Y * var_Y
                Dim Z As Double = ref_Z * var_Z

                ' Then, convert from XYZ to RGB.
                ' Standards used Observer = 2, Illuminant = D65
                ' ref_X = 95.047, ref_Y = 100.000, ref_Z = 108.883

                Dim var_R As Double = X * 0.032406 + Y * (-0.015372) + Z * (-0.004986)
                Dim var_G As Double = X * (-0.009689) + Y * 0.018758 + Z * 0.000415
                Dim var_B As Double = X * 0.000557 + Y * (-0.00204) + Z * 0.01057

                If var_R > 0.0031308 Then
                    var_R = 1.055 * Math.Pow(var_R, 1 / 2.4) - 0.055
                Else
                    var_R = 12.92 * var_R
                End If

                If var_G > 0.0031308 Then
                    var_G = 1.055 * Math.Pow(var_G, 1 / 2.4) - 0.055
                Else
                    var_G = 12.92 * var_G
                End If

                If var_B > 0.0031308 Then
                    var_B = 1.055 * Math.Pow(var_B, 1 / 2.4) - 0.055
                Else
                    var_B = 12.92 * var_B
                End If

                Dim nRed As Integer = CInt(Fix(var_R * 256.0))
                Dim nGreen As Integer = CInt(Fix(var_G * 256.0))
                Dim nBlue As Integer = CInt(Fix(var_B * 256.0))

                If nRed < 0 Then
                    nRed = 0
                ElseIf nRed > 255 Then
                    nRed = 255
                End If

                If nGreen < 0 Then
                    nGreen = 0
                ElseIf nGreen > 255 Then
                    nGreen = 255
                End If

                If nBlue < 0 Then
                    nBlue = 0
                ElseIf nBlue > 255 Then
                    nBlue = 255
                End If

                destRow(i).R = CByte(nRed)
                destRow(i).G = CByte(nGreen)
                destRow(i).B = CByte(nBlue)

                idxSrc += context.ByteDepth
            Next
        End Sub

#End Region

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Private ReadOnly rgbExponent As Double = 1 / 2.19921875

        ''' <summary>
        ''' Reads a 32-bit float pixel value from a byte array at the given byte
        ''' offset (native endianness - the image data has already been
        ''' byte-swapped to native order by the time it reaches this decoder)
        ''' and converts it to an 8-bit sRGB byte value.
        ''' </summary>
        Private Function RGBByteFromHDRFloat(data As Byte(), offset As Integer) As Byte
            Dim value = BitConverter.ToSingle(data, offset)
            Return Util.ToByteUnchecked(CInt(255 * Math.Pow(CDbl(value), rgbExponent)))
        End Function

    End Module

End Namespace
