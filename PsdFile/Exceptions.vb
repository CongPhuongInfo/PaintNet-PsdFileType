'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Exceptions.cs
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

Namespace PhotoshopFile

    <Serializable>
    Public Class PsdInvalidException
        Inherits Exception

        Public Sub New()
        End Sub

        Public Sub New(message As String)
            MyBase.New(message)
        End Sub
    End Class

    <Serializable>
    Public Class RleException
        Inherits Exception

        Public Sub New()
        End Sub

        Public Sub New(message As String)
            MyBase.New(message)
        End Sub
    End Class

End Namespace
