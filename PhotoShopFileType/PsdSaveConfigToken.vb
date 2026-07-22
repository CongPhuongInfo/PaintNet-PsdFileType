'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PsdSaveConfigToken.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2012 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports PaintDotNet

Namespace PaintDotNet.Data.PhotoshopFileType

    <Serializable>
    Public Class PsdSaveConfigToken
        Inherits SaveConfigToken

        Public Overrides Function Clone() As Object
            Return New PsdSaveConfigToken(Me)
        End Function

        Public Property RleCompress As Boolean

        Public Sub New(rleCompress As Boolean)
            Me.RleCompress = rleCompress
        End Sub

        Protected Sub New(copyMe As PsdSaveConfigToken)
            Me.RleCompress = copyMe.RleCompress
        End Sub

        Public Overrides Sub Validate()
            MyBase.Validate()
        End Sub

    End Class

End Namespace
