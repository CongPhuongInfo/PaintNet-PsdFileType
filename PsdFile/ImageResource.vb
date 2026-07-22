'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/ImageResource.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic

Namespace PhotoshopFile

    Public Enum ResourceID
        Undefined = 0
        MacPrintInfo = 1001
        ResolutionInfo = 1005
        AlphaChannelNames = 1006
        DisplayInfo = 1007
        Caption = 1008
        BorderInfo = 1009
        BackgroundColor = 1010
        PrintFlags = 1011
        MultichannelHalftoneInfo = 1012
        ColorHalftoneInfo = 1013
        DuotoneHalftoneInfo = 1014
        MultichannelTransferFunctions = 1015
        ColorTransferFunctions = 1016
        DuotoneTransferFunctions = 1017
        DuotoneImageInfo = 1018
        BlackWhiteRange = 1019
        EpsOptions = 1021
        QuickMaskInfo = 1022
        LayerStateInfo = 1024
        WorkingPathUnsaved = 1025
        LayersGroupInfo = 1026
        IptcNaa = 1028
        RawFormatImageMode = 1029
        JpegQuality = 1030
        GridGuidesInfo = 1032
        ThumbnailBgr = 1033
        CopyrightInfo = 1034
        Url = 1035
        ThumbnailRgb = 1036
        GlobalAngle = 1037
        ColorSamplersObsolete = 1038
        IccProfile = 1039
        Watermark = 1040
        IccUntagged = 1041
        EffectsVisible = 1042
        SpotHalftone = 1043
        DocumentSpecific = 1044
        UnicodeAlphaNames = 1045
        IndexedColorTableCount = 1046
        TransparentIndex = 1047
        GlobalAltitude = 1049
        Slices = 1050
        WorkflowUrl = 1051
        JumpToXpep = 1052
        AlphaIdentifiers = 1053
        UrlList = 1054
        VersionInfo = 1057
        ExifData1 = 1058
        ExifData3 = 1059
        XmpMetadata = 1060
        CaptionDigest = 1061
        PrintScale = 1062
        PixelAspectRatio = 1064
        LayerComps = 1065
        AlternateDuotoneColors = 1066
        AlternateSpotColors = 1067
        LayerSelectionIDs = 1069
        HdrToningInfo = 1070
        PrintInfo = 1071
        LayerGroupsEnabled = 1072
        ColorSamplers = 1073
        MeasurementScale = 1074
        TimelineInfo = 1075
        SheetDisclosure = 1076
        FloatDisplayInfo = 1077
        OnionSkins = 1078
        CountInfo = 1080
        PrintSettingsInfo = 1082
        PrintStyle = 1083
        MacNSPrintInfo = 1084
        WinDevMode = 1085
        AutoSaveFilePath = 1086
        AutoSaveFormat = 1087
        PathInfo = 2000 ' 2000-2999: Path Information
        ClippingPathName = 2999
        LightroomWorkflow = 8000
        PrintFlagsInfo = 10000
    End Enum

    ''' <summary>
    ''' Abstract class for Image Resources
    ''' </summary>
    Public MustInherit Class ImageResource

        Private _signature As String
        Public Property Signature As String
            Get
                Return _signature
            End Get
            Set(value As String)
                If value.Length <> 4 Then
                    Throw New ArgumentException($"{NameOf(Signature)} must be 4 characters in length.")
                End If
                _signature = value
            End Set
        End Property

        Public Property Name As String

        Public MustOverride ReadOnly Property ID As ResourceID

        Protected Sub New(name As String)
            Signature = "8BIM"
            Me.Name = name
        End Sub

        ''' <summary>
        ''' Write out the image resource block: header and data.
        ''' </summary>
        Public Sub Save(writer As PsdBinaryWriter)
            Util.DebugMessage(writer.BaseStream, "Save, Begin, ImageResource")

            writer.WriteAsciiChars(Signature)
            writer.Write(CUShort(ID))
            writer.WritePascalString(Name, 2)

            ' Length is unpadded, but data is even-padded
            Dim startPosition = writer.BaseStream.Position
            Using New PsdBlockLengthWriter(writer)
                WriteData(writer)
            End Using
            writer.WritePadding(startPosition, 2)

            Util.DebugMessage(writer.BaseStream, $"Save, End, ImageResource, {ID}")
        End Sub

        ''' <summary>
        ''' Write the data for this image resource.
        ''' </summary>
        Protected MustOverride Sub WriteData(writer As PsdBinaryWriter)

        Public Overrides Function ToString() As String
            Return $"{ID} {Name}"
        End Function

    End Class

    ''' <summary>
    ''' Creates the appropriate subclass of ImageResource.
    ''' </summary>
    Public Module ImageResourceFactory

        Public Function CreateImageResource(reader As PsdBinaryReader) As ImageResource
            Util.DebugMessage(reader.BaseStream, "Load, Begin, ImageResource")

            Dim signature = reader.ReadAsciiChars(4)
            Dim resourceIdInt = reader.ReadUInt16()
            Dim name = reader.ReadPascalString(2)
            Dim dataLength = CInt(reader.ReadUInt32())

            Dim dataPaddedLength = Util.RoundUp(dataLength, 2)
            Dim endPosition = reader.BaseStream.Position + dataPaddedLength

            Dim resource As ImageResource = Nothing
            Dim resourceId = CType(resourceIdInt, ResourceID)
            Select Case resourceId
                Case ResourceID.ResolutionInfo
                    resource = New ResolutionInfo(reader, name)
                Case ResourceID.ThumbnailRgb, ResourceID.ThumbnailBgr
                    resource = New Thumbnail(reader, resourceId, name, dataLength)
                Case ResourceID.AlphaChannelNames
                    resource = New AlphaChannelNames(reader, name, dataLength)
                Case ResourceID.UnicodeAlphaNames
                    resource = New UnicodeAlphaNames(reader, name, dataLength)
                Case ResourceID.VersionInfo
                    resource = New VersionInfo(reader, name)
                Case Else
                    resource = New RawImageResource(reader, signature, resourceId, name, dataLength)
            End Select

            Util.DebugMessage(reader.BaseStream, $"Load, End, ImageResource, {resourceId}")

            ' Reposition the reader if we do not consume the full resource block.
            ' This takes care of the even-padding, and also preserves forward-
            ' compatibility in case a resource block is later extended with
            ' additional properties.
            If reader.BaseStream.Position < endPosition Then
                reader.BaseStream.Position = endPosition
            End If

            ' However, overruns are definitely an error.
            If reader.BaseStream.Position > endPosition Then
                Throw New PsdInvalidException("Corruption detected in resource.")
            End If

            Return resource
        End Function

    End Module

    Public Class ImageResources
        Inherits List(Of ImageResource)

        Public Sub New()
            MyBase.New()
        End Sub

        Public Function [Get](id As ResourceID) As ImageResource
            Return Find(Function(x) x.ID = id)
        End Function

        Public Sub [Set](resource As ImageResource)
            Dim matchId As Predicate(Of ImageResource) =
                Function(res) res.ID = resource.ID

            Dim itemIdx = Me.FindIndex(matchId)
            Dim lastItemIdx = Me.FindLastIndex(matchId)

            If itemIdx = -1 Then
                Add(resource)
            ElseIf itemIdx <> lastItemIdx Then
                RemoveAll(matchId)
                Insert(itemIdx, resource)
            Else
                Me(itemIdx) = resource
            End If
        End Sub

    End Class

End Namespace
