'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/RleWriter.cs
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
' PORTING NOTE: the original walks the input with a raw byte* for speed and
' uses unchecked casts to wrap the PackBits flag-counter byte. VB.NET has no
' unsafe/fixed support and no bit-reinterpreting cast, so this rewrite uses
' array indexing and the Util.ToByteUnchecked helper. Logic is unchanged.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.IO

Namespace PhotoshopFile

    Public Class RleWriter

        Private Const maxPacketLength As Integer = 128

        ' Current task
        Private rleLock As Object
        Private stream As Stream
        Private data As Byte()
        Private offset As Integer

        ' Current packet
        Private isRepeatPacket As Boolean
        Private idxPacketStart As Integer
        Private packetLength As Integer

        Private runValue As Byte
        Private runLength As Integer

        Public Sub New(stream As Stream)
            rleLock = New Object()
            Me.stream = stream
        End Sub

        ''' <summary>
        ''' Encodes byte data using PackBits RLE compression.
        ''' </summary>
        ''' <param name="data">Raw data to be encoded.</param>
        ''' <param name="offset">Offset at which to begin transferring data.</param>
        ''' <param name="count">Number of bytes of data to transfer.</param>
        ''' <returns>Number of compressed bytes written to the stream.</returns>
        ''' <remarks>
        ''' There are multiple ways to encode two-byte runs:
        '''   1. Apple PackBits only encodes three-byte runs as repeats.
        '''   2. Adobe Photoshop encodes two-byte runs as repeats, unless preceded
        '''      by literals.
        '''   3. TIFF PackBits recommends that two-byte runs be encoded as repeats,
        '''      unless preceded *and* followed by literals.
        '''
        ''' This class adopts the Photoshop behavior, as it has slightly better
        ''' compression efficiency than Apple PackBits, and is easier to implement
        ''' than TIFF PackBits.
        ''' </remarks>
        Public Function Write(data As Byte(), offset As Integer, count As Integer) As Integer
            If Not Util.CheckBufferBounds(data, offset, count) Then
                Throw New ArgumentOutOfRangeException()
            End If

            ' We cannot encode a count of 0, because the PackBits flag-counter byte
            ' uses 0 to indicate a length of 1.
            If count = 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(count))
            End If

            SyncLock rleLock
                Dim startPosition = stream.Position

                Me.data = data
                Me.offset = offset

                Dim bytesEncoded = EncodeToStream(offset, offset + count)
                Debug.Assert(bytesEncoded = count, "Encoded byte count should match the argument.")

                Return CInt(stream.Position - startPosition)
            End SyncLock
        End Function

        Private Sub ClearPacket()
            Me.isRepeatPacket = False
            Me.packetLength = 0
        End Sub

        Private Sub WriteRepeatPacket(length As Integer)
            Dim header As Byte = Util.ToByteUnchecked(1 - length)
            stream.WriteByte(header)
            stream.WriteByte(runValue)
        End Sub

        Private Sub WriteLiteralPacket(length As Integer)
            Dim header As Byte = Util.ToByteUnchecked(length - 1)
            stream.WriteByte(header)
            stream.Write(data, idxPacketStart, length)
        End Sub

        Private Sub WritePacket()
            If isRepeatPacket Then
                WriteRepeatPacket(packetLength)
            Else
                WriteLiteralPacket(packetLength)
            End If
        End Sub

        Private Sub StartPacket(count As Integer,
                                 isRepeatPacket As Boolean, runLength As Integer, value As Byte)
            Me.isRepeatPacket = isRepeatPacket

            Me.packetLength = runLength
            Me.runLength = runLength
            Me.runValue = value

            Me.idxPacketStart = offset + count
        End Sub

        Private Sub ExtendPacketAndRun(value As Byte)
            packetLength += 1
            runLength += 1
        End Sub

        Private Sub ExtendPacketStartNewRun(value As Byte)
            packetLength += 1
            runLength = 1
            runValue = value
        End Sub

        ''' <summary>
        ''' Encodes data(startIdx) through data(endIdxExclusive - 1), inclusive.
        ''' </summary>
        Private Function EncodeToStream(startIdx As Integer, endIdxExclusive As Integer) As Integer
            ' Begin the first packet.
            Dim idx As Integer = startIdx
            StartPacket(0, False, 1, data(idx))
            Dim numBytesEncoded As Integer = 1
            idx += 1

            ' Loop invariant: Packet is never empty.
            While idx < endIdxExclusive
                Dim value As Byte = data(idx)

                If packetLength = 1 Then
                    isRepeatPacket = (value = runValue)
                    If isRepeatPacket Then
                        ExtendPacketAndRun(value)
                    Else
                        ExtendPacketStartNewRun(value)
                    End If

                ElseIf packetLength = maxPacketLength Then
                    ' Packet is full, so write it out and start a new one.
                    WritePacket()
                    StartPacket(numBytesEncoded, False, 1, value)

                ElseIf isRepeatPacket Then
                    ' Decide whether to continue the repeat packet.
                    If value = runValue Then
                        ExtendPacketAndRun(value)
                    Else
                        ' Different color, so terminate the run and start a new packet.
                        WriteRepeatPacket(packetLength)
                        StartPacket(numBytesEncoded, False, 1, value)
                    End If

                Else
                    ' Decide whether to continue the literal packet.
                    If value = runValue Then
                        ExtendPacketAndRun(value)

                        ' A 3-byte run terminates the literal and starts a new repeat
                        ' packet. That's because the 3-byte run can be encoded as a
                        ' 2-byte repeat. So even if the run ends at 3, we've already
                        ' paid for the next flag-counter byte.
                        If runLength = 3 Then
                            ' The 3-byte run can come in the middle of a literal packet,
                            ' but not at the beginning. The first 2 bytes of the run
                            ' should've triggered a repeat packet.
                            Debug.Assert(packetLength > 3)

                            ' -2 because numBytesEncoded has not yet been incremented
                            WriteLiteralPacket(packetLength - 3)
                            StartPacket(numBytesEncoded - 2, True, 3, value)
                        End If
                    Else
                        ExtendPacketStartNewRun(value)
                    End If
                End If

                idx += 1
                numBytesEncoded += 1
            End While

            ' Loop terminates with a non-empty packet waiting to be written out.
            WritePacket()
            ClearPacket()

            Return numBytesEncoded
        End Function

    End Class

End Namespace
