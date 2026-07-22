'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/LayerInfo.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Collections.Generic

Namespace PhotoshopFile

    Friend Module LayerInfoFactory

        Friend Sub LoadAll(reader As PsdBinaryReader, psdFile As PsdFile,
                            layerInfoList As List(Of LayerInfo), endPosition As Long,
                            globalLayerInfo As Boolean)
            ' LayerInfo has a 12-byte minimum length. Anything shorter should be
            ' ignored as padding.
            While endPosition - reader.BaseStream.Position >= 12
                Dim layerInfo = Load(reader, psdFile, globalLayerInfo)
                layerInfoList.Add(layerInfo)
            End While

            If reader.BaseStream.Position < endPosition Then
                reader.BaseStream.Position = endPosition
            ElseIf reader.BaseStream.Position > endPosition Then
                Throw New PsdInvalidException("Read past the end of the LayerInfo fields.")
            End If
        End Sub

        ''' <summary>
        ''' Loads the next LayerInfo record.
        ''' </summary>
        ''' <param name="reader">The file reader</param>
        ''' <param name="psdFile">The PSD file.</param>
        ''' <param name="globalLayerInfo">True if the LayerInfo record is being
        '''   loaded from the end of the Layer and Mask Information section;
        '''   false if it is being loaded from the end of a Layer record.</param>
        ''' <returns>LayerInfo object if it was successfully read, or Nothing if
        '''   padding was found.</returns>
        Private Function Load(reader As PsdBinaryReader, psdFile As PsdFile,
                               globalLayerInfo As Boolean) As LayerInfo
            Util.DebugMessage(reader.BaseStream, "Load, Begin, LayerInfo")

            ' Most keys have undocumented signatures, so we always accept either one.
            Dim signature = reader.ReadAsciiChars(4)
            If (signature <> "8BIM") AndAlso (signature <> "8B64") Then
                Throw New PsdInvalidException($"{NameOf(LayerInfo)} signature invalid, must be 8BIM or 8B64.")
            End If

            Dim key = reader.ReadAsciiChars(4)
            Dim hasLongLength = LayerInfoUtil.HasLongLength(signature, key, psdFile.IsLargeDocument)
            Dim length = If(hasLongLength, reader.ReadInt64(), CLng(reader.ReadInt32()))
            Dim startPosition = reader.BaseStream.Position

            Dim result As LayerInfo
            Select Case key
                Case "Layr", "Lr16", "Lr32"
                    result = New InfoLayers(reader, psdFile, key, length)
                Case "lsct", "lsdk"
                    result = New LayerSectionInfo(reader, key, CInt(length))
                Case "luni"
                    result = New LayerUnicodeName(reader)
                Case Else
                    result = New RawLayerInfo(reader, signature, key, length)
            End Select

            ' May have additional padding applied.
            Dim endPosition = startPosition + length
            If reader.BaseStream.Position < endPosition Then
                reader.BaseStream.Position = endPosition
            End If

            ' Documentation states that the length is even-padded. Actually:
            '   1. Most keys have 4-padded lengths.
            '   2. However, some keys (LMsk) have even-padded lengths.
            '   3. Other keys (Txt2, Lr16, Lr32) have unpadded lengths.
            '
            ' Photoshop writes data that is always 4-padded, even when the stated
            ' length is not a multiple of 4. The length mismatch seems to occur
            ' only on global layer info. We do not read extra padding in other
            ' cases because third-party programs are likely to follow the spec.

            If globalLayerInfo Then
                reader.ReadPadding(startPosition, 4)
            End If

            Util.DebugMessage(reader.BaseStream, $"Load, End, LayerInfo, {result.Signature}, {result.Key}")
            Return result
        End Function

    End Module

    Friend Module LayerInfoUtil

        Friend Function HasLongLength(signature As String, key As String, isLargeDocument As Boolean) As Boolean
            If Not isLargeDocument Then
                Return False
            End If

            ' Keys with 8B64 signatures always have 8-byte lengths in PSB files.
            If signature = "8B64" Then
                Return True
            End If

            Select Case key
                ' These keys are documented to have 8-byte lengths in PSB files. Some
                ' keys have 8BIM signatures. Other keys have 8B64 signatures, but this
                ' fact is undocumented, so they are still hardcoded in the list.
                Case "LMsk", "Lr16", "Lr32", "Layr", "Mt16", "Mt32", "Mtrn",
                     "Alph", "FMsk", "lnk2", "FEid", "FXid", "PxSD"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

    End Module

    Public MustInherit Class LayerInfo

        Public MustOverride ReadOnly Property Signature As String

        Public MustOverride ReadOnly Property Key As String

        Protected MustOverride Sub WriteData(writer As PsdBinaryWriter)

        Public Sub Save(writer As PsdBinaryWriter, globalLayerInfo As Boolean, isLargeDocument As Boolean)
            Util.DebugMessage(writer.BaseStream, $"Save, Begin, LayerInfo, {Signature}, {Key}")

            writer.WriteAsciiChars(Signature)
            writer.WriteAsciiChars(Key)

            Dim startPosition = writer.BaseStream.Position
            Using lengthWriter = New PsdBlockLengthWriter(writer,
                LayerInfoUtil.HasLongLength(Signature, Key, isLargeDocument))
                ' Depending on the key, the length may be unpadded, 2-padded, or
                ' 4-padded. Thus, it is up to each implementation of WriteData to
                ' pad the length correctly.
                WriteData(writer)
            End Using

            ' Data for global layer info is always padded to a multiple of 4,
            ' even if this causes the stated length to be incorrect.
            If globalLayerInfo Then
                writer.WritePadding(startPosition, 4)
            End If

            Util.DebugMessage(writer.BaseStream, $"Save, End, LayerInfo, {Signature}, {Key}")
        End Sub

    End Class

End Namespace
