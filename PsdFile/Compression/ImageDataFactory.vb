'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Compression/ImageDataFactory.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2016 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Drawing

Namespace PhotoshopFile.Compression

    Public Module ImageDataFactory

        ''' <summary>
        ''' Creates an ImageData object to compress or decompress image data.
        ''' </summary>
        ''' <param name="channel">The Channel associated with the image data.</param>
        ''' <param name="data">The image data to be decompressed, or Nothing if
        '''   image data is to be compressed.</param>
        Public Function Create(channel As Channel, data As Byte()) As ImageData
            Dim bitDepth = channel.Layer.PsdFile.BitDepth
            Dim imageData As ImageData

            Select Case channel.ImageCompression
                Case ImageCompression.Raw
                    imageData = New RawImage(data, channel.Rect.Size, bitDepth)

                Case ImageCompression.Rle
                    imageData = New RleImage(data, channel.RleRowLengths,
                        channel.Rect.Size, bitDepth)

                Case ImageCompression.Zip
                    ' Photoshop treats 32-bit Zip as 32-bit ZipPrediction
                    imageData = If(bitDepth = 32,
                        CreateZipPredict(data, channel.Rect.Size, bitDepth),
                        New ZipImage(data, channel.Rect.Size, bitDepth))

                Case ImageCompression.ZipPrediction
                    imageData = CreateZipPredict(data, channel.Rect.Size, bitDepth)

                Case Else
                    Throw New PsdInvalidException("Unknown image compression method.")
            End Select

            ' Reverse endianness of multi-byte image data
            imageData = WrapEndianness(imageData)

            Return imageData
        End Function

        Private Function CreateZipPredict(data As Byte(), size As Size, bitDepth As Integer) As ImageData
            Select Case bitDepth
                Case 16
                    Return New ZipPredict16Image(data, size)
                Case 32
                    Return New ZipPredict32Image(data, size)
                Case Else
                    Throw New PsdInvalidException(
                        "ZIP with prediction is only available for 16 and 32 bit depths.")
            End Select
        End Function

        Private Function WrapEndianness(imageData As ImageData) As ImageData
            ' Single-byte image does not require endianness reversal
            If imageData.BitDepth <= 8 Then
                Return imageData
            End If

            ' Bytes will be reordered by the compressor, so no wrapper is needed
            If (TypeOf imageData Is ZipPredict16Image) OrElse (TypeOf imageData Is ZipPredict32Image) Then
                Return imageData
            End If

            Return New EndianReverser(imageData)
        End Function

    End Module

End Namespace
