'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/DocumentLoadContext.cs
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

Imports PhotoshopFile

Namespace PaintDotNet.Data.PhotoshopFileType

    ''' <summary>
    ''' Controls the loading of a PSD file into a Paint.NET Document.
    ''' </summary>
    Public Class DocumentLoadContext
        Inherits LoadContext

        Public Sub New()
            MyBase.New()
        End Sub

        Public Overrides Sub OnLoadLayersHeader(psdFile As PsdFile)
            PsdLoad.CheckSufficientMemory(psdFile)
        End Sub

        Public Overrides Sub OnLoadLayerHeader(layer As PhotoshopFile.Layer)
            Dim psdFile = layer.PsdFile
            If psdFile.ColorMode = PsdColorMode.Multichannel Then
                PsdLoad.CheckSufficientMemory(psdFile)
            End If
        End Sub

    End Class

End Namespace
