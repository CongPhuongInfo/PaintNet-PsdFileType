'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PsdPluginResources.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Globalization
Imports System.Resources

Namespace PaintDotNet.Data.PhotoshopFileType

    Public Module PsdPluginResources

        Private rm As New ResourceManager("Photoshop.Resources", GetType(PsdPluginResources).Assembly)

        Public Function GetString(resourceName As String) As String
            ' We really ought to use .Name, but .NET 3.5 returns legacy three-letter
            ' region codes rather than the two-letter ISO 3166 codes that MSDN
            ' claims it returns. Since Paint.NET is currently translated into only
            ' one region per language, we can get by without the region for now.
            Dim languageCode = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName

            ' We currently have very few localized strings, so it's not really worth
            ' deploying satellite assemblies for each language. For now, we simply
            ' prefix the resource name with the language code.
            Dim taggedResourceName = languageCode & "_" & resourceName
            Dim s = rm.GetString(taggedResourceName)
            If s IsNot Nothing Then
                Return s
            End If

            ' If no translation is available, fall back to the untagged resource
            Return rm.GetString(resourceName)
        End Function

    End Module

End Namespace
