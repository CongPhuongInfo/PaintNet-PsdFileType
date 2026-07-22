'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/PsdBinaryWriter.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2016 Tao Yue
'
' Portions of this file are provided under the BSD 3-clause License:
'   Copyright (c) 2006, Jonas Beckeman
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: see PsdBinaryReader.vb - endian swapping is done with
' BitConverter + Array.Reverse instead of unsafe pointer code.
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
    ''' Writes PSD data types in big-endian byte order.
    ''' </summary>
    Public Class PsdBinaryWriter
        Implements IDisposable

        Private writer As BinaryWriter
        Private encoding As Encoding

        Friend ReadOnly Property BaseStream As Stream
            Get
                ' Flush the writer so that the Stream.Position is correct.
                Flush()
                Return writer.BaseStream
            End Get
        End Property

        Public Sub New(stream As Stream, encoding As Encoding)
            Me.encoding = encoding

            ' Specifying ASCII encoding will help catch any accidental calls to
            ' BinaryWriter.Write(String). Since we do not own the Stream, the
            ' constructor is called with leaveOpen = True.
            writer = New BinaryWriter(stream, Encoding.ASCII, True)
        End Sub

        Public Sub Flush()
            writer.Flush()
        End Sub

        Public Sub Write(rect As Rectangle)
            Write(rect.Top)
            Write(rect.Left)
            Write(rect.Bottom)
            Write(rect.Right)
        End Sub

        ''' <summary>
        ''' Pad the length of a block to a multiple.
        ''' </summary>
        ''' <param name="startPosition">Starting position of the padded block.</param>
        ''' <param name="padMultiple">Byte multiple to pad to.</param>
        Public Sub WritePadding(startPosition As Long, padMultiple As Integer)
            Dim length = writer.BaseStream.Position - startPosition
            Dim padBytes = Util.GetPadding(CInt(length), padMultiple)
            For i As Long = 0 To padBytes - 1
                writer.Write(CByte(0))
            Next
        End Sub

        ''' <summary>
        ''' Write string as ASCII characters, without a length prefix.
        ''' </summary>
        Public Sub WriteAsciiChars(s As String)
            Dim bytes = Encoding.ASCII.GetBytes(s)
            writer.Write(bytes)
        End Sub

        ''' <summary>
        ''' Writes a Pascal string using the specified encoding.
        ''' </summary>
        ''' <param name="s">Unicode string to convert to the encoding.</param>
        ''' <param name="padMultiple">Byte multiple that the Pascal string is padded to.</param>
        ''' <param name="maxBytes">Maximum number of bytes to write.</param>
        Public Sub WritePascalString(s As String, padMultiple As Integer, Optional maxBytes As Byte = 255)
            Dim startPosition = writer.BaseStream.Position

            Dim bytesArray = encoding.GetBytes(s)
            If bytesArray.Length > maxBytes Then
                Dim tempArray(CInt(maxBytes) - 1) As Byte
                Array.Copy(bytesArray, tempArray, CInt(maxBytes))
                bytesArray = tempArray
            End If

            writer.Write(CByte(bytesArray.Length))
            writer.Write(bytesArray)
            WritePadding(startPosition, padMultiple)
        End Sub

        ''' <summary>
        ''' Write a Unicode string to the stream.
        ''' </summary>
        Public Sub WriteUnicodeString(s As String)
            Write(s.Length)
            Dim data = Encoding.BigEndianUnicode.GetBytes(s)
            Write(data)
        End Sub

        Public Sub Write(value As Boolean)
            writer.Write(value)
        End Sub

        Public Sub Write(value As Byte())
            writer.Write(value)
        End Sub

        Public Sub Write(value As Byte)
            writer.Write(value)
        End Sub

        Public Sub Write(value As Int16)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToInt16(bytes, 0))
        End Sub

        Public Sub Write(value As Int32)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToInt32(bytes, 0))
        End Sub

        Public Sub Write(value As Int64)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToInt64(bytes, 0))
        End Sub

        Public Sub Write(value As UInt16)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToUInt16(bytes, 0))
        End Sub

        Public Sub Write(value As UInt32)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToUInt32(bytes, 0))
        End Sub

        Public Sub Write(value As UInt64)
            Dim bytes = BitConverter.GetBytes(value)
            Array.Reverse(bytes)
            writer.Write(BitConverter.ToUInt64(bytes, 0))
        End Sub

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
                If writer IsNot Nothing Then
                    ' BinaryWriter.Dispose() is protected in the original C#;
                    ' Close() achieves the same result and flushes automatically.
                    writer.Close()
                    writer = Nothing
                End If
            End If

            disposedValue = True
        End Sub

#End Region

    End Class

End Namespace
