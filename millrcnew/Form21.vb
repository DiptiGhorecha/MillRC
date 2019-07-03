﻿''' <summary>
''' report view form - receipt printing for godown
''' In this module we are displaying .dat file in reachtextbox
''' </summary>
Public Class Form21
    Dim formloaded As Boolean = False
    Private checkPrint As Integer

    Private Sub Form21_Load(sender As Object, e As EventArgs) Handles Me.Load
        '''''set position,size of form and richtextbox
        Me.MdiParent = MainMDIForm
        Me.Top = MainMDIForm.Label1.Height + MainMDIForm.MainMenuStrip.Height
        Me.Left = 0
        Me.Width = Parent.Width - 8
        Me.Height = Parent.Height - 95
        RichTextBox1.Width = Me.Width - 40
        RichTextBox1.Height = Me.Height - 60
        formloaded = True
    End Sub

    Private Sub Form21_Move(sender As Object, e As EventArgs) Handles Me.Move
        '''''''keep the form position fix
        If formloaded Then
            If (Right > Parent.ClientSize.Width) Then Left = Parent.ClientSize.Width - Width
            If (Bottom > Parent.ClientSize.Height) Then Top = Parent.ClientSize.Height - Height
            If (Left < 0) Then Left = 0
            If (Top < 0) Then Top = 0
            If (Top < 87) Then Top = 87
        End If
    End Sub
    Private Sub PrintDocument1_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles PrintDocument1.BeginPrint
        checkPrint = 0
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        checkPrint = RichTextBox1.Print(checkPrint, RichTextBox1.TextLength, e)
        If checkPrint < RichTextBox1.TextLength Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
        End If
        '  End If
    End Sub
End Class