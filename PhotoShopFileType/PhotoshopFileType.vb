'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PhotoshopFileType.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2019 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports PaintDotNet
Imports System.IO

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Class PhotoshopFileTypes
        Implements IFileTypeFactory

        Public Shared ReadOnly Psd As FileType = New PhotoshopFileType()

        Private Shared fileTypes As FileType() = New FileType() {Psd}

        Public Function GetFileTypeInstances() As FileType() Implements IFileTypeFactory.GetFileTypeInstances
            Return CType(fileTypes.Clone(), FileType())
        End Function

    End Class

    Public Class PhotoshopFileType
        Inherits FileType

        Public Sub New()
            MyBase.New("Photoshop", New FileTypeOptions() With {
                .LoadExtensions = New String() {".psd", ".psb"},
                .SaveExtensions = New String() {".psd", ".psb"},
                .SupportsLayers = True
            })
        End Sub

        Public Overrides Function CreateSaveConfigWidget() As SaveConfigWidget
            Return New PsdSaveConfigWidget()
        End Function

        Protected Overrides Function OnCreateDefaultSaveConfigToken() As SaveConfigToken
            Return New PsdSaveConfigToken(True)
        End Function

        Protected Overrides Sub OnSave(input As Document, output As Stream, token As SaveConfigToken,
                                        scratchSurface As Surface, callback As ProgressEventHandler)
            ' Because the function signature takes in a Stream, we cannot force the
            ' extension to .PSB for large documents. However, Photoshop is happy
            ' to load a PSB file even if it has a .PSD extension.

            Dim psdToken = CType(token, PsdSaveConfigToken)
            PsdSave.Save(input, output, psdToken, scratchSurface, callback)
        End Sub

        Protected Overrides Function OnLoad(input As Stream) As Document
            Return PsdLoad.Load(input)
        End Function

    End Class

End Namespace
