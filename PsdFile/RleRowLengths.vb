'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PsdFile/RleRowLengths.cs
'
' This software is provided under the MIT License:
'   Copyright (c) 2006-2007 Frank Blumenberg
'   Copyright (c) 2010-2017 Tao Yue
'
' See LICENSE.txt for complete licensing and attribution information.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Strict On
Option Explicit On

Imports System.Linq

Namespace PhotoshopFile

    Public Class RleRowLengths

        Public Property Values As Integer()

        Public ReadOnly Property Total As Long
            Get
                Return Values.Sum(Function(x) CLng(x))
            End Get
        End Property

        Default Public Property Item(i As Integer) As Integer
            Get
                Return Values(i)
            End Get
            Set(value As Integer)
                Values(i) = value
            End Set
        End Property

        Public Sub New(rowCount As Integer)
            Values = New Integer(rowCount - 1) {}
        End Sub

        Public Sub New(reader As PsdBinaryReader, rowCount As Integer, isLargeDocument As Boolean)
            Me.New(rowCount)
            For i As Integer = 0 To rowCount - 1
                Values(i) = If(isLargeDocument, reader.ReadInt32(), CInt(reader.ReadUInt16()))
            Next
        End Sub

        Public Sub Write(writer As PsdBinaryWriter, isLargeDocument As Boolean)
            For i As Integer = 0 To Values.Length - 1
                If isLargeDocument Then
                    writer.Write(Values(i))
                Else
                    writer.Write(CType(Values(i), UInt16))
                End If
            Next
        End Sub

    End Class

End Namespace
