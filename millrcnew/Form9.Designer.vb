﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmInvSummary
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmInvSummary))
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.TxtSrch = New System.Windows.Forms.TextBox()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.HSNRadio2 = New System.Windows.Forms.RadioButton()
        Me.HSNRadio1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.B2BRadio2 = New System.Windows.Forms.RadioButton()
        Me.B2BRadio3 = New System.Windows.Forms.RadioButton()
        Me.B2BRadio1 = New System.Windows.Forms.RadioButton()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Items.AddRange(New Object() {"2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030", "2031", "2032", "2033", "2034", "2035", "2036", "2037", "2038", "2039", "2040", "2041", "2042", "2043", "2044", "2045", "2046", "2047", "2048", "2049", "2050"})
        Me.ComboBox4.Location = New System.Drawing.Point(108, 46)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(154, 21)
        Me.ComboBox4.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Year"
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"})
        Me.ComboBox3.Location = New System.Drawing.Point(109, 10)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(153, 21)
        Me.ComboBox3.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(37, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Month"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(110, 114)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(152, 20)
        Me.TextBox2.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(78, 13)
        Me.Label5.TabIndex = 62
        Me.Label5.Text = "To Invoice No."
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(110, 79)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(152, 20)
        Me.TextBox1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 79)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 13)
        Me.Label1.TabIndex = 59
        Me.Label1.Text = "From Invoice No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(465, 71)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 13)
        Me.Label2.TabIndex = 66
        Me.Label2.Text = "Press F1 for help"
        Me.Label2.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(128, 295)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(50, 22)
        Me.Button3.TabIndex = 12
        Me.Button3.Text = "Cancel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(72, 295)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(50, 22)
        Me.Button2.TabIndex = 11
        Me.Button2.Text = "Print"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(16, 295)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(50, 22)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "View"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AllowUserToResizeColumns = False
        Me.DataGridView2.AllowUserToResizeRows = False
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView2.Location = New System.Drawing.Point(15, 328)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(617, 172)
        Me.DataGridView2.TabIndex = 67
        Me.DataGridView2.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.TxtSrch)
        Me.GroupBox5.Location = New System.Drawing.Point(429, 266)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(200, 50)
        Me.GroupBox5.TabIndex = 68
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Search by godown"
        '
        'TxtSrch
        '
        Me.TxtSrch.Location = New System.Drawing.Point(12, 19)
        Me.TxtSrch.Name = "TxtSrch"
        Me.TxtSrch.Size = New System.Drawing.Size(177, 20)
        Me.TxtSrch.TabIndex = 8
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(520, 128)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(120, 20)
        Me.TextBox3.TabIndex = 69
        Me.TextBox3.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(504, 96)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(120, 20)
        Me.TextBox4.TabIndex = 70
        Me.TextBox4.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(15, 140)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(511, 50)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "HSN No."
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"All", "Rental Or Leasing Services Involving Own Or Leased Residential Property", "Rental Or Leasing Services Involving Own Or Leased Non-residential Property"})
        Me.ComboBox1.Location = New System.Drawing.Point(6, 17)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(499, 21)
        Me.ComboBox1.TabIndex = 5
        Me.ComboBox1.Text = "All"
        '
        'HSNRadio2
        '
        Me.HSNRadio2.AutoSize = True
        Me.HSNRadio2.Location = New System.Drawing.Point(556, 157)
        Me.HSNRadio2.Name = "HSNRadio2"
        Me.HSNRadio2.Size = New System.Drawing.Size(36, 17)
        Me.HSNRadio2.TabIndex = 74
        Me.HSNRadio2.TabStop = True
        Me.HSNRadio2.Text = "All"
        Me.HSNRadio2.UseVisualStyleBackColor = True
        Me.HSNRadio2.Visible = False
        '
        'HSNRadio1
        '
        Me.HSNRadio1.AutoSize = True
        Me.HSNRadio1.Location = New System.Drawing.Point(552, 180)
        Me.HSNRadio1.Name = "HSNRadio1"
        Me.HSNRadio1.Size = New System.Drawing.Size(72, 17)
        Me.HSNRadio1.TabIndex = 73
        Me.HSNRadio1.TabStop = True
        Me.HSNRadio1.Text = "HSN wise"
        Me.HSNRadio1.UseVisualStyleBackColor = True
        Me.HSNRadio1.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.B2BRadio2)
        Me.GroupBox2.Controls.Add(Me.B2BRadio3)
        Me.GroupBox2.Controls.Add(Me.B2BRadio1)
        Me.GroupBox2.Location = New System.Drawing.Point(15, 197)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(247, 50)
        Me.GroupBox2.TabIndex = 6
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Filter2"
        '
        'B2BRadio2
        '
        Me.B2BRadio2.AutoSize = True
        Me.B2BRadio2.Enabled = False
        Me.B2BRadio2.Location = New System.Drawing.Point(90, 19)
        Me.B2BRadio2.Name = "B2BRadio2"
        Me.B2BRadio2.Size = New System.Drawing.Size(45, 17)
        Me.B2BRadio2.TabIndex = 7
        Me.B2BRadio2.TabStop = True
        Me.B2BRadio2.Text = "B2C"
        Me.B2BRadio2.UseVisualStyleBackColor = True
        '
        'B2BRadio3
        '
        Me.B2BRadio3.AutoSize = True
        Me.B2BRadio3.Location = New System.Drawing.Point(173, 19)
        Me.B2BRadio3.Name = "B2BRadio3"
        Me.B2BRadio3.Size = New System.Drawing.Size(36, 17)
        Me.B2BRadio3.TabIndex = 8
        Me.B2BRadio3.TabStop = True
        Me.B2BRadio3.Text = "All"
        Me.B2BRadio3.UseVisualStyleBackColor = True
        '
        'B2BRadio1
        '
        Me.B2BRadio1.AutoSize = True
        Me.B2BRadio1.Enabled = False
        Me.B2BRadio1.Location = New System.Drawing.Point(15, 20)
        Me.B2BRadio1.Name = "B2BRadio1"
        Me.B2BRadio1.Size = New System.Drawing.Size(45, 17)
        Me.B2BRadio1.TabIndex = 6
        Me.B2BRadio1.TabStop = True
        Me.B2BRadio1.Text = "B2B"
        Me.B2BRadio1.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(13, 255)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(91, 41)
        Me.Label6.TabIndex = 75
        Me.Label6.Text = "Report File Name (Without Extn)"
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(110, 255)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(152, 20)
        Me.TextBox5.TabIndex = 9
        Me.TextBox5.Text = "Invoices_checklist_format1"
        '
        'FrmInvSummary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(644, 512)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.HSNRadio2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.HSNRadio1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.Label3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FrmInvSummary"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Invoice Checklist- Format1"
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBox4 As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents ComboBox3 As ComboBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Button3 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents TxtSrch As TextBox
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents HSNRadio2 As RadioButton
    Friend WithEvents HSNRadio1 As RadioButton
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents B2BRadio3 As RadioButton
    Friend WithEvents B2BRadio1 As RadioButton
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents B2BRadio2 As RadioButton
    Friend WithEvents Label6 As Label
    Friend WithEvents TextBox5 As TextBox
End Class
