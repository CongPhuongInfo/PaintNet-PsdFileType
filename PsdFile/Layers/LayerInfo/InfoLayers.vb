'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Layers/LayerInfo/InfoLayers.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Namespace PhotoshopFile

    ''' <summary>
    ''' Layers that are stored as Additional Info, rather than in the main
    ''' Layers section of the PSD file.
    ''' </summary>
    ''' <remarks>
    ''' Photoshop stores layers in the Additional Info section for 16-bit and
    ''' 32-bit depth images. The Layers section in the PSD file is left empty.
    '''
    ''' This appears to be for backward-compatibility purposes, but it is not
    ''' required. Photoshop will successfully load a high-bitdepth image that
    ''' puts the layers in the Layers section.
    ''' </remarks>
    Public Class InfoLayers
        Inherits LayerInfo

        Public Overrides ReadOnly Property Signature As String
            Get
                Return "8BIM"
            End Get
        End Property

        Private _key As String
        Public Overrides ReadOnly Property Key As String
            Get
                Return _key
            End Get
        End Property

        Public Property PsdFile As PsdFile

        Public Sub New(psdFile As PsdFile, key As String)
            Me.PsdFile = psdFile

            Select Case key
                ' The key does not have to match the bit depth, but it does have to
                ' be one of the known values.
                Case "Layr", "Lr16", "Lr32"
                    Me._key = key
                Case Else
                    Throw New PsdInvalidException($"{NameOf(InfoLayers)} key must be Layr, Lr16, or Lr32.")
            End Select
        End Sub

        Public Sub New(reader As PsdBinaryReader, psdFile As PsdFile, key As String, dataLength As Long)
            Me.New(psdFile, key)

            If psdFile.Layers.Count > 0 Then
                Throw New PsdInvalidException("Cannot have both regular layers and Additional Info layers")
            End If

            Dim endPosition = reader.BaseStream.Position + dataLength
            psdFile.LoadLayers(reader, False)

            If reader.BaseStream.Position <> endPosition Then
                Throw New PsdInvalidException($"Incorrect length for {NameOf(InfoLayers)}.")
            End If
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            PsdFile.SaveLayersData(writer)
        End Sub

    End Class

End Namespace
