'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/RleReader.cs
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
' PORTING NOTE: the original pins the buffer and walks it with a raw byte*
' for speed. VB.NET has no unsafe/fixed support, so this rewrite uses plain
' array indexing (and Util.Fill instead of the pointer fill loop). Logic and
' exceptions thrown are unchanged.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.IO

Namespace PhotoshopFile

    Public Class RleReader

        Private stream As Stream

        Public Sub New(stream As Stream)
            Me.stream = stream
        End Sub

        ''' <summary>
        ''' Decodes a PackBits RLE stream.
        ''' </summary>
        ''' <param name="buffer">Output buffer for decoded data.</param>
        ''' <param name="offset">Offset at which to begin writing.</param>
        ''' <param name="count">Number of bytes to decode from the stream.</param>
        Public Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            If Not Util.CheckBufferBounds(buffer, offset, count) Then
                Throw New ArgumentOutOfRangeException()
            End If
            If count = 0 Then
                Return 0
            End If

            Dim bytesLeft As Integer = count
            Dim bufferIdx As Integer = offset

            While bytesLeft > 0
                ' ReadByte returns an unsigned byte value (0-255), but we want it
                ' reinterpreted as a signed byte, same as the original's
                ' unchecked((sbyte)stream.ReadByte()).
                Dim flagCounter As SByte = Util.ToSByteUnchecked(CByte(stream.ReadByte()))

                If flagCounter > 0 Then
                    ' Raw packet
                    Dim readLength As Integer = CInt(flagCounter) + 1
                    If bytesLeft < readLength Then
                        Throw New RleException("Raw packet overruns the decode window.")
                    End If

                    stream.Read(buffer, bufferIdx, readLength)

                    bufferIdx += readLength
                    bytesLeft -= readLength

                ElseIf flagCounter > -128 Then
                    ' RLE packet
                    Dim runLength As Integer = 1 - CInt(flagCounter)
                    Dim byteValue As Byte = CByte(stream.ReadByte())
                    If runLength > bytesLeft Then
                        Throw New RleException("RLE packet overruns the decode window.")
                    End If

                    Util.Fill(buffer, bufferIdx, runLength, byteValue)

                    bufferIdx += runLength
                    bytesLeft -= runLength

                Else
                    ' The canonical PackBits algorithm will never emit 0x80 (-128), but
                    ' some programs do. Simply skip over the byte.
                End If
            End While

            Debug.Assert(bytesLeft = 0)
            Return count - bytesLeft
        End Function

    End Class

End Namespace
