'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/BlendModeMapping.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2014 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.Runtime.CompilerServices

Imports PaintDotNet
Imports PhotoshopFile

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Module BlendModeMapping

        ''' <summary>
        ''' Convert between Paint.NET and Photoshop blend modes.
        ''' </summary>
        <Extension()>
        Public Function ToPsdBlendMode(pdnBlendMode As LayerBlendMode) As String
            Select Case pdnBlendMode
                Case LayerBlendMode.Normal
                    Return PsdBlendMode.Normal

                Case LayerBlendMode.Multiply
                    Return PsdBlendMode.Multiply
                Case LayerBlendMode.Additive
                    Return PsdBlendMode.LinearDodge
                Case LayerBlendMode.ColorBurn
                    Return PsdBlendMode.ColorBurn
                Case LayerBlendMode.ColorDodge
                    Return PsdBlendMode.ColorDodge
                Case LayerBlendMode.Overlay
                    Return PsdBlendMode.Overlay
                Case LayerBlendMode.Difference
                    Return PsdBlendMode.Difference
                Case LayerBlendMode.Lighten
                    Return PsdBlendMode.Lighten
                Case LayerBlendMode.Darken
                    Return PsdBlendMode.Darken
                Case LayerBlendMode.Screen
                    Return PsdBlendMode.Screen

                ' Paint.NET blend modes without a Photoshop equivalent are saved
                ' as Normal.
                Case LayerBlendMode.Glow, LayerBlendMode.Negation,
                     LayerBlendMode.Reflect, LayerBlendMode.Xor
                    Return PsdBlendMode.Normal

                Case Else
                    Debug.Fail("Unknown Paint.NET blend mode.")
                    Return PsdBlendMode.Normal
            End Select
        End Function

        ''' <summary>
        ''' Convert a Photoshop blend mode to a Paint.NET BlendOp.
        ''' </summary>
        Public Function FromPsdBlendMode(blendModeKey As String) As LayerBlendMode
            Select Case blendModeKey
                Case PsdBlendMode.Normal
                    Return LayerBlendMode.Normal

                Case PsdBlendMode.Multiply
                    Return LayerBlendMode.Multiply
                Case PsdBlendMode.LinearDodge
                    Return LayerBlendMode.Additive
                Case PsdBlendMode.ColorBurn
                    Return LayerBlendMode.ColorBurn
                Case PsdBlendMode.ColorDodge
                    Return LayerBlendMode.ColorDodge
                Case PsdBlendMode.Overlay
                    Return LayerBlendMode.Overlay
                Case PsdBlendMode.Difference
                    Return LayerBlendMode.Difference
                Case PsdBlendMode.Lighten
                    Return LayerBlendMode.Lighten
                Case PsdBlendMode.Darken
                    Return LayerBlendMode.Darken
                Case PsdBlendMode.Screen
                    Return LayerBlendMode.Screen

                    ' Photoshop blend modes without a Paint.NET equivalent are loaded
                    ' as Normal.
                Case Else
                    Return LayerBlendMode.Normal
            End Select
        End Function

    End Module

End Namespace
