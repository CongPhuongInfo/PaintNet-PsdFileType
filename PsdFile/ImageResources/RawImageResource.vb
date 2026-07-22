'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/RawImageResource.cs
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
    ''' Stores the raw data for unimplemented image resource types.
    ''' </summary>
    Public Class RawImageResource
        Inherits ImageResource

        Private _data As Byte()
        Public Property Data As Byte()
            Get
                Return _data
            End Get
            Private Set(value As Byte())
                _data = value
            End Set
        End Property

        Private _id As ResourceID
        Public Overrides ReadOnly Property ID As ResourceID
            Get
                Return _id
            End Get
        End Property

        Public Sub New(resourceId As ResourceID, name As String)
            MyBase.New(name)
            Me._id = resourceId
        End Sub

        Public Sub New(reader As PsdBinaryReader, signature As String,
                       resourceId As ResourceID, name As String, numBytes As Integer)
            MyBase.New(name)
            Me.Signature = signature
            Me._id = resourceId
            Data = reader.ReadBytes(numBytes)
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            writer.Write(Data)
        End Sub

    End Class

End Namespace
