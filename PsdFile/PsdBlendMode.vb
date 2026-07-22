'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/PsdBlendMode.cs
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

Namespace PhotoshopFile

    Public NotInheritable Class PsdBlendMode
        Private Sub New()
        End Sub

        Public Const Normal As String = "norm"
        Public Const Darken As String = "dark"
        Public Const Lighten As String = "lite"
        Public Const Hue As String = "hue "
        Public Const Saturation As String = "sat "
        Public Const Color As String = "colr"
        Public Const Luminosity As String = "lum "
        Public Const Multiply As String = "mul "
        Public Const Screen As String = "scrn"
        Public Const Dissolve As String = "diss"
        Public Const Overlay As String = "over"
        Public Const HardLight As String = "hLit"
        Public Const SoftLight As String = "sLit"
        Public Const Difference As String = "diff"
        Public Const Exclusion As String = "smud"
        Public Const ColorDodge As String = "div "
        Public Const ColorBurn As String = "idiv"
        Public Const LinearBurn As String = "lbrn"
        Public Const LinearDodge As String = "lddg"
        Public Const VividLight As String = "vLit"
        Public Const LinearLight As String = "lLit"
        Public Const PinLight As String = "pLit"
        Public Const HardMix As String = "hMix"
        Public Const PassThrough As String = "pass"
        Public Const DarkerColor As String = "dkCl"
        Public Const LighterColor As String = "lgCl"
        Public Const Subtract As String = "fsub"
        Public Const Divide As String = "fdiv"
    End Class

End Namespace
