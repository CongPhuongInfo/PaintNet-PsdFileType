'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/LoadContext.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2016 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Text

Namespace PhotoshopFile

    ''' <summary>
    ''' Contains settings and callbacks that affect the loading of a PSD file.
    ''' </summary>
    Public Class LoadContext

        Public Property Encoding As Encoding

        Public Sub New()
            Encoding = Encoding.Default
        End Sub

        Public Overridable Sub OnLoadLayersHeader(psdFile As PsdFile)
        End Sub

        Public Overridable Sub OnLoadLayerHeader(layer As Layer)
        End Sub

    End Class

End Namespace
