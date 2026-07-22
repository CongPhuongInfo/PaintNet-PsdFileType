'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/PsdBlockLengthWriter.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2014 Tao Yue
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

Namespace PhotoshopFile

    ''' <summary>
    ''' Writes the actual length in front of the data block upon disposal.
    ''' </summary>
    Friend Class PsdBlockLengthWriter
        Implements IDisposable

        Private disposedValue As Boolean = False

        Private lengthPosition As Long
        Private startPosition As Long
        Private hasLongLength As Boolean
        Private writer As PsdBinaryWriter

        Public Sub New(writer As PsdBinaryWriter)
            Me.New(writer, False)
        End Sub

        Public Sub New(writer As PsdBinaryWriter, hasLongLength As Boolean)
            Me.writer = writer
            Me.hasLongLength = hasLongLength

            ' Store position so that we can return to it when the length is known.
            lengthPosition = writer.BaseStream.Position

            ' Write a sentinel value as a placeholder for the length.
            writer.Write(CUInt(&HFEEDFEEDL))
            If hasLongLength Then
                writer.Write(CUInt(&HFEEDFEEDL))
            End If

            ' Store the start position of the data block so that we can calculate
            ' its length when we're done writing.
            startPosition = writer.BaseStream.Position
        End Sub

        Public Sub Write()
            Dim endPosition = writer.BaseStream.Position

            writer.BaseStream.Position = lengthPosition
            Dim length As Long = endPosition - startPosition
            If hasLongLength Then
                writer.Write(length)
            Else
                writer.Write(CUInt(length))
            End If

            writer.BaseStream.Position = endPosition
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not Me.disposedValue Then
                Write()
                Me.disposedValue = True
            End If
        End Sub

    End Class

End Namespace
