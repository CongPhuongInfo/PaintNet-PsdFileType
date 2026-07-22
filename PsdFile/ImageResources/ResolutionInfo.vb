'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResources/ResolutionInfo.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Namespace PhotoshopFile

    ''' <summary>
    ''' Summary description for ResolutionInfo.
    ''' </summary>
    Public Class ResolutionInfo
        Inherits ImageResource

        Public Overrides ReadOnly Property ID As ResourceID
            Get
                Return ResourceID.ResolutionInfo
            End Get
        End Property

        ''' <summary>
        ''' Horizontal DPI.
        ''' </summary>
        Public Property HDpi As UFixed16_16

        ''' <summary>
        ''' Vertical DPI.
        ''' </summary>
        Public Property VDpi As UFixed16_16

        ''' <summary>
        ''' 1 = pixels per inch, 2 = pixels per centimeter
        ''' </summary>
        Public Enum ResUnit
            PxPerInch = 1
            PxPerCm = 2
        End Enum

        ''' <summary>
        ''' Display units for horizontal resolution. This only affects the
        ''' user interface; the resolution is still stored in the PSD file
        ''' as pixels/inch.
        ''' </summary>
        Public Property HResDisplayUnit As ResUnit

        ''' <summary>
        ''' Display units for vertical resolution.
        ''' </summary>
        Public Property VResDisplayUnit As ResUnit

        ''' <summary>
        ''' Physical units.
        ''' </summary>
        Public Enum Unit
            Inches = 1
            Centimeters = 2
            Points = 3
            Picas = 4
            Columns = 5
        End Enum

        Public Property WidthDisplayUnit As Unit

        Public Property HeightDisplayUnit As Unit

        Public Sub New()
            MyBase.New(String.Empty)
        End Sub

        Public Sub New(reader As PsdBinaryReader, name As String)
            MyBase.New(name)

            Me.HDpi = New UFixed16_16(reader.ReadUInt32())
            Me.HResDisplayUnit = CType(reader.ReadInt16(), ResUnit)
            Me.WidthDisplayUnit = CType(reader.ReadInt16(), Unit)

            Me.VDpi = New UFixed16_16(reader.ReadUInt32())
            Me.VResDisplayUnit = CType(reader.ReadInt16(), ResUnit)
            Me.HeightDisplayUnit = CType(reader.ReadInt16(), Unit)
        End Sub

        Protected Overrides Sub WriteData(writer As PsdBinaryWriter)
            writer.Write(HDpi.[Integer])
            writer.Write(HDpi.Fraction)
            writer.Write(CShort(HResDisplayUnit))
            writer.Write(CShort(WidthDisplayUnit))

            writer.Write(VDpi.[Integer])
            writer.Write(VDpi.Fraction)
            writer.Write(CShort(VResDisplayUnit))
            writer.Write(CShort(HeightDisplayUnit))
        End Sub

    End Class

End Namespace
