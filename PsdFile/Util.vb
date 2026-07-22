'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/Util.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2020 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
' PORTING NOTE: The original C# uses `unsafe`/pointer code (fixed, byte*) to
' fill, invert, and byte-swap buffers in place for speed. VB.NET does not
' support unsafe/pointer code at all, so every one of those routines below
' has been rewritten using plain safe array indexing. Behavior is identical;
' only the implementation technique changed.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.CompilerServices

Namespace PhotoshopFile

    Public Module Util

        <DebuggerDisplay("Top = {Top}, Bottom = {Bottom}, Left = {Left}, Right = {Right}")>
        Public Structure RectanglePosition
            Public Property Top As Integer
            Public Property Bottom As Integer
            Public Property Left As Integer
            Public Property Right As Integer
        End Structure

        <Extension()>
        Public Function IntersectWith(thisRect As Rectangle, rect As Rectangle) As Rectangle
            Dim result = thisRect
            result.Intersect(rect)
            Return result
        End Function

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Fills a range of a buffer with a byte value.
        ''' </summary>
        Public Sub Fill(buffer As Byte(), startIndex As Integer, count As Integer, value As Byte)
            For i As Integer = startIndex To startIndex + count - 1
                buffer(i) = value
            Next
        End Sub

        ''' <summary>
        ''' Inverts (bitwise-NOT) a range of a buffer.
        ''' </summary>
        Public Sub Invert(buffer As Byte(), startIndex As Integer, count As Integer)
            For i As Integer = startIndex To startIndex + count - 1
                buffer(i) = CByte(buffer(i) Xor &HFF)
            Next
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Reverses the endianness of a word of arbitrary length, in place,
        ''' starting at the given offset within the buffer.
        ''' </summary>
        Public Sub SwapBytes(buffer As Byte(), offset As Integer, length As Integer)
            Dim i As Integer = 0
            While i < length \ 2
                Dim t As Byte = buffer(offset + i)
                buffer(offset + i) = buffer(offset + length - i - 1)
                buffer(offset + length - i - 1) = t
                i += 1
            End While
        End Sub

        ''' <summary>
        ''' Reverses the endianness of a 2-byte word, in place.
        ''' </summary>
        Public Sub SwapBytes2(buffer As Byte(), offset As Integer)
            SwapBytes(buffer, offset, 2)
        End Sub

        ''' <summary>
        ''' Reverses the endianness of a 4-byte word, in place.
        ''' </summary>
        Public Sub SwapBytes4(buffer As Byte(), offset As Integer)
            SwapBytes(buffer, offset, 4)
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public Sub SwapByteArray(bitDepth As Integer, byteArray As Byte(), startIdx As Integer, count As Integer)
            Select Case bitDepth
                Case 1, 8
                    ' No swapping needed for 1-bit or 8-bit depths.
                Case 16
                    SwapByteArray2(byteArray, startIdx, count)
                Case 32
                    SwapByteArray4(byteArray, startIdx, count)
                Case Else
                    Throw New Exception("Byte-swapping implemented only for 16-bit and 32-bit depths.")
            End Select
        End Sub

        ''' <summary>
        ''' Reverses the endianness of 2-byte words in a byte array.
        ''' </summary>
        ''' <param name="byteArray">Byte array containing the sequence on which to swap endianness</param>
        ''' <param name="startIdx">Byte index of the first word to swap</param>
        ''' <param name="count">Number of words to swap</param>
        Public Sub SwapByteArray2(byteArray As Byte(), startIdx As Integer, count As Integer)
            Dim endIdx As Integer = startIdx + count * 2
            If byteArray.Length < endIdx Then
                Throw New IndexOutOfRangeException()
            End If

            Dim idx As Integer = startIdx
            While idx < endIdx
                SwapBytes2(byteArray, idx)
                idx += 2
            End While
        End Sub

        ''' <summary>
        ''' Reverses the endianness of 4-byte words in a byte array.
        ''' </summary>
        ''' <param name="byteArray">Byte array containing the sequence on which to swap endianness</param>
        ''' <param name="startIdx">Byte index of the first word to swap</param>
        ''' <param name="count">Number of words to swap</param>
        Public Sub SwapByteArray4(byteArray As Byte(), startIdx As Integer, count As Integer)
            Dim endIdx As Integer = startIdx + count * 4
            If byteArray.Length < endIdx Then
                Throw New IndexOutOfRangeException()
            End If

            Dim idx As Integer = startIdx
            While idx < endIdx
                SwapBytes4(byteArray, idx)
                idx += 4
            End While
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Calculates the number of bytes required to store a row of an image
        ''' with the specified bit depth.
        ''' </summary>
        ''' <param name="size">The size of the image in pixels.</param>
        ''' <param name="bitDepth">The bit depth of the image.</param>
        ''' <returns>The number of bytes needed to store a row of the image.</returns>
        Public Function BytesPerRow(size As Size, bitDepth As Integer) As Integer
            Select Case bitDepth
                Case 1
                    Return (size.Width + 7) \ 8
                Case Else
                    Return size.Width * BytesFromBitDepth(bitDepth)
            End Select
        End Function

        ''' <summary>
        ''' Round the integer to a multiple.
        ''' </summary>
        Public Function RoundUp(value As Integer, multiple As Integer) As Integer
            If value = 0 Then
                Return 0
            End If

            If Math.Sign(value) <> Math.Sign(multiple) Then
                Throw New ArgumentException("value and multiple cannot have opposite signs.")
            End If

            Dim result = value
            Dim remainder = result Mod multiple
            If remainder > 0 Then
                result += (multiple - remainder)
            End If
            Return result
        End Function

        ''' <summary>
        ''' Get number of bytes required to pad to the specified multiple.
        ''' </summary>
        Public Function GetPadding(length As Integer, padMultiple As Integer) As Integer
            If (length < 0) OrElse (padMultiple < 0) Then
                Throw New ArgumentException()
            End If

            Dim remainder = length Mod padMultiple
            If remainder = 0 Then
                Return 0
            End If

            Dim padding = padMultiple - remainder
            Return padding
        End Function

        ''' <summary>
        ''' Returns the number of bytes needed to store a single pixel of the
        ''' specified bit depth.
        ''' </summary>
        Public Function BytesFromBitDepth(depth As Integer) As Integer
            Select Case depth
                Case 1, 8
                    Return 1
                Case 16
                    Return 2
                Case 32
                    Return 4
                Case Else
                    Throw New ArgumentException("Invalid bit depth.")
            End Select
        End Function

        <Extension()>
        Public Function MinChannelCount(colorMode As PsdColorMode) As Short
            Select Case colorMode
                Case PsdColorMode.Bitmap, PsdColorMode.Duotone, PsdColorMode.Grayscale,
                     PsdColorMode.Indexed, PsdColorMode.Multichannel
                    Return 1
                Case PsdColorMode.Lab, PsdColorMode.RGB
                    Return 3
                Case PsdColorMode.CMYK
                    Return 4
            End Select

            Throw New ArgumentException("Unknown color mode.")
        End Function

        ''' <summary>
        ''' Verify that the offset and count will remain within the bounds of the
        ''' buffer.
        ''' </summary>
        ''' <returns>True if in bounds, false if out of bounds.</returns>
        Public Function CheckBufferBounds(data As Byte(), offset As Integer, count As Integer) As Boolean
            If offset < 0 Then
                Return False
            End If
            If count < 0 Then
                Return False
            End If
            If offset + count > data.Length Then
                Return False
            End If

            Return True
        End Function

        Public Sub CheckByteArrayLength(length As Long)
            If length < 0 Then
                Throw New Exception("Byte array cannot have a negative length.")
            End If
            If length > &H7FFFFFC7L Then
                Throw New OutOfMemoryException("Byte array cannot exceed 2,147,483,591 in length.")
            End If
        End Sub

        ''' <summary>
        ''' Writes a message to the debug console, indicating the current position
        ''' in the stream in both decimal and hexadecimal formats.
        ''' </summary>
        <Conditional("DEBUG")>
        Public Sub DebugMessage(stream As Stream, message As String)
            Debug.WriteLine($"0x{stream.Position:x}, {stream.Position}, {message}")
        End Sub

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' PORTING NOTE: helpers replacing C#'s "unchecked((byte)/(sbyte) ...)"
        ' bit-reinterpreting casts, which have no direct VB.NET equivalent
        ' (VB numeric conversions perform range-checked value conversion, not
        ' bit-pattern reinterpretation). Used by RleReader/RleWriter.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''' <summary>
        ''' Reinterprets a Byte's bit pattern as an SByte (equivalent to C#'s
        ''' unchecked((sbyte)byteValue)).
        ''' </summary>
        Public Function ToSByteUnchecked(value As Byte) As SByte
            If value > 127 Then
                Return CSByte(CInt(value) - 256)
            Else
                Return CSByte(value)
            End If
        End Function

        ''' <summary>
        ''' Reinterprets an Integer's low 8 bits as a Byte (equivalent to C#'s
        ''' unchecked((byte)intValue)).
        ''' </summary>
        Public Function ToByteUnchecked(value As Integer) As Byte
            Return CByte(((value Mod 256) + 256) Mod 256)
        End Function

    End Module

    ''' <summary>
    ''' Fixed-point decimal, with 16-bit integer and 16-bit fraction.
    ''' </summary>
    Public Class UFixed16_16

        Public Property [Integer] As UInt16
        Public Property Fraction As UInt16

        Public Sub New(integerPart As UInt16, fractionPart As UInt16)
            [Integer] = integerPart
            Fraction = fractionPart
        End Sub

        ''' <summary>
        ''' Split the high and low words of a 32-bit unsigned integer into a
        ''' fixed-point number.
        ''' </summary>
        Public Sub New(value As UInt32)
            [Integer] = CType(value >> 16, UInt16)
            Fraction = CType(value And &HFFFFUI, UInt16)
        End Sub

        Public Sub New(value As Double)
            If value >= 65536.0 Then
                Throw New OverflowException()
            End If
            If value < 0 Then
                Throw New OverflowException()
            End If

            [Integer] = CType(value, UInt16)

            ' Round instead of truncate, because doubles may not represent the
            ' fraction exactly.
            Fraction = CType((value - [Integer]) * 65536 + 0.5, UInt16)
        End Sub

        Public Shared Widening Operator CType(value As UFixed16_16) As Double
            Return CDbl(value.[Integer]) + value.Fraction / 65536.0
        End Operator

    End Class

End Namespace
