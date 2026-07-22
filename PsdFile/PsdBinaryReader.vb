'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/PsdBinaryReader.cs
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
' PORTING NOTE: the original reads a value with the platform's (little-endian)
' BinaryReader and then byte-swaps it in place using unsafe pointer code.
' VB.NET has no unsafe/pointer support, so here each value is converted to a
' byte array with BitConverter, reversed with Array.Reverse, and converted
' back. Same result, safe code.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Text

Namespace PhotoshopFile

    ''' <summary>
    ''' Reads PSD data types in big-endian byte order.
    ''' </summary>
    Public Class PsdBinaryReader
        Implements IDisposable

        Private reader As BinaryReader
        Private encoding As Encoding

        Public ReadOnly Property BaseStream As Stream
            Get
                Return reader.BaseStream
            End Get
        End Property

        Public Sub New(stream As Stream, otherReader As PsdBinaryReader)
            Me.New(stream, otherReader.encoding)
        End Sub

        Public Sub New(stream As Stream, encoding As Encoding)
            Me.encoding = encoding

            ' ReadPascalString and ReadUnicodeString handle encoding explicitly.
            ' BinaryReader.ReadString() is never called, so it is constructed with
            ' ASCII encoding to make accidental usage obvious.
            reader = New BinaryReader(stream, Encoding.ASCII)
        End Sub

        Public Function ReadByte() As Byte
            Return reader.ReadByte()
        End Function

        Public Function ReadBytes(count As Integer) As Byte()
            Return reader.ReadBytes(count)
        End Function

        Public Function ReadBoolean() As Boolean
            Return reader.ReadBoolean()
        End Function

        Public Function ReadInt16() As Int16
            Dim bytes = BitConverter.GetBytes(reader.ReadInt16())
            Array.Reverse(bytes)
            Return BitConverter.ToInt16(bytes, 0)
        End Function

        Public Function ReadInt32() As Int32
            Dim bytes = BitConverter.GetBytes(reader.ReadInt32())
            Array.Reverse(bytes)
            Return BitConverter.ToInt32(bytes, 0)
        End Function

        Public Function ReadInt64() As Int64
            Dim bytes = BitConverter.GetBytes(reader.ReadInt64())
            Array.Reverse(bytes)
            Return BitConverter.ToInt64(bytes, 0)
        End Function

        Public Function ReadUInt16() As UInt16
            Dim bytes = BitConverter.GetBytes(reader.ReadUInt16())
            Array.Reverse(bytes)
            Return BitConverter.ToUInt16(bytes, 0)
        End Function

        Public Function ReadUInt32() As UInt32
            Dim bytes = BitConverter.GetBytes(reader.ReadUInt32())
            Array.Reverse(bytes)
            Return BitConverter.ToUInt32(bytes, 0)
        End Function

        Public Function ReadUInt64() As UInt64
            Dim bytes = BitConverter.GetBytes(reader.ReadUInt64())
            Array.Reverse(bytes)
            Return BitConverter.ToUInt64(bytes, 0)
        End Function

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Read padding to get to the byte multiple for the block.
        ''' </summary>
        ''' <param name="startPosition">Starting position of the padded block.</param>
        ''' <param name="padMultiple">Byte multiple that the block is padded to.</param>
        Public Sub ReadPadding(startPosition As Long, padMultiple As Integer)
            ' Pad to specified byte multiple
            Dim totalLength = reader.BaseStream.Position - startPosition
            Dim padBytes = Util.GetPadding(CInt(totalLength), padMultiple)
            ReadBytes(padBytes)
        End Sub

        Public Function ReadRectangle() As Rectangle
            Dim rect = New Rectangle()
            rect.Y = ReadInt32()
            rect.X = ReadInt32()
            rect.Height = ReadInt32() - rect.Y
            rect.Width = ReadInt32() - rect.X
            Return rect
        End Function

        ''' <summary>
        ''' Read a fixed-length ASCII string.
        ''' </summary>
        Public Function ReadAsciiChars(count As Integer) As String
            Dim bytes = reader.ReadBytes(count)
            Dim s = Encoding.ASCII.GetString(bytes)
            Return s
        End Function

        ''' <summary>
        ''' Read a Pascal string using the specified encoding.
        ''' </summary>
        ''' <param name="padMultiple">Byte multiple that the Pascal string is padded to.</param>
        Public Function ReadPascalString(padMultiple As Integer) As String
            Dim startPosition = reader.BaseStream.Position

            Dim stringLength As Byte = ReadByte()
            Dim bytes = ReadBytes(CInt(stringLength))
            ReadPadding(startPosition, padMultiple)

            ' Default decoder uses best-fit fallback, so it will not throw any
            ' exceptions if unknown characters are encountered.
            Dim str = encoding.GetString(bytes)
            Return str
        End Function

        Public Function ReadUnicodeString() As String
            Dim numChars = ReadInt32()
            Dim length = 2 * numChars
            Dim data = ReadBytes(length)
            Dim str = Encoding.BigEndianUnicode.GetString(data, 0, length)

            Return str
        End Function

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#Region "IDisposable"

        Private disposedValue As Boolean = False

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            ' Check to see if Dispose has already been called.
            If disposedValue Then
                Return
            End If

            If disposing Then
                If reader IsNot Nothing Then
                    ' BinaryReader.Dispose() is protected in the original C#;
                    ' Close() achieves the same result here.
                    reader.Close()
                    reader = Nothing
                End If
            End If

            disposedValue = True
        End Sub

#End Region

    End Class

End Namespace
