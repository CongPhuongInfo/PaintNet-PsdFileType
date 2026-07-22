'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Photoshop PSD FileType Plugin for Paint.NET
' VB.NET port of PhotoShopFileType/PsdSaveConfigWidget.cs
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

Imports System
Imports System.Windows.Forms

Imports PaintDotNet

Namespace PaintDotNet.Data.PhotoshopFileType

    ''' <summary>
    ''' Summary description for TgaSaveConfigWidget.
    ''' </summary>
    Public Class PsdSaveConfigWidget
        Inherits PaintDotNet.SaveConfigWidget(Of PhotoshopFileType, PsdSaveConfigToken)

        Private rleCompressCheckBox As System.Windows.Forms.CheckBox

        ''' <summary>
        ''' Required designer variable.
        ''' </summary>
        Private components As System.ComponentModel.Container = Nothing

        Public Sub New()
            MyBase.New(New PhotoshopFileType())

            ' This call is required by the Windows.Forms Form Designer.
            InitializeComponent()

            'Me.bpp24Radio.Text = PdnResources.GetString("TgaSaveConfigWidget.Bpp24Radio.Text")
            'Me.bpp32Radio.Text = PdnResources.GetString("TgaSaveConfigWidget.Bpp32Radio.Text")
            'Me.bppLabel.Text = PdnResources.GetString("TgaSaveConfigWidget.BppLabel.Text")
            'Me.rleCompressCheckBox.Text = PdnResources.GetString("TgaSaveConfigWidget.RleCompressCheckBox.Text")
        End Sub

        Protected Overrides Function CreateTokenFromWidget() As PsdSaveConfigToken
            Return New PsdSaveConfigToken(Me.rleCompressCheckBox.Checked)
        End Function

        Protected Overrides Sub InitWidgetFromToken(token As PsdSaveConfigToken)
            If token IsNot Nothing Then
                Me.rleCompressCheckBox.Checked = token.RleCompress
            Else
                Me.rleCompressCheckBox.Checked = True
            End If
        End Sub

        ''' <summary>
        ''' Clean up any resources being used.
        ''' </summary>
        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If components IsNot Nothing Then
                    components.Dispose()
                End If
            End If

            MyBase.Dispose(disposing)
        End Sub

#Region "Component Designer generated code"
        ''' <summary>
        ''' Required method for Designer support - do not modify
        ''' the contents of this method with the code editor.
        ''' </summary>
        Private Sub InitializeComponent()
            Me.rleCompressCheckBox = New System.Windows.Forms.CheckBox()
            Me.SuspendLayout()
            '
            ' rleCompressCheckBox
            '
            Me.rleCompressCheckBox.Location = New System.Drawing.Point(0, 0)
            Me.rleCompressCheckBox.Name = "rleCompressCheckBox"
            Me.rleCompressCheckBox.Size = New System.Drawing.Size(184, 24)
            Me.rleCompressCheckBox.TabIndex = 0
            Me.rleCompressCheckBox.Text = PsdPluginResources.GetString("SaveDialog_RleCompression")
            AddHandler Me.rleCompressCheckBox.CheckedChanged, AddressOf Me.OnCheckedChanged
            '
            ' PsdSaveConfigWidget
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Controls.Add(Me.rleCompressCheckBox)
            Me.Name = "PsdSaveConfigWidget"
            Me.Size = New System.Drawing.Size(180, 104)
            Me.ResumeLayout(False)
        End Sub
#End Region

        Private Sub OnCheckedChanged(sender As Object, e As System.EventArgs)
            Me.UpdateToken()
        End Sub

    End Class

End Namespace
