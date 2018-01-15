<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmcalculator
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmcalculator))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtsource = New System.Windows.Forms.TextBox()
        Me.btnequal = New System.Windows.Forms.Button()
        Me.btnclear = New System.Windows.Forms.Button()
        Me.btnaddminus = New System.Windows.Forms.Button()
        Me.btnx = New System.Windows.Forms.Button()
        Me.btndivide = New System.Windows.Forms.Button()
        Me.btnmultiply = New System.Windows.Forms.Button()
        Me.btnminus = New System.Windows.Forms.Button()
        Me.btnadd = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btndecimal = New System.Windows.Forms.Button()
        Me.btn9 = New System.Windows.Forms.Button()
        Me.btn8 = New System.Windows.Forms.Button()
        Me.btn7 = New System.Windows.Forms.Button()
        Me.btn6 = New System.Windows.Forms.Button()
        Me.btn5 = New System.Windows.Forms.Button()
        Me.btn4 = New System.Windows.Forms.Button()
        Me.Btn3 = New System.Windows.Forms.Button()
        Me.Btn2 = New System.Windows.Forms.Button()
        Me.btn1 = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtsource)
        Me.GroupBox1.Controls.Add(Me.btnequal)
        Me.GroupBox1.Controls.Add(Me.btnclear)
        Me.GroupBox1.Controls.Add(Me.btnaddminus)
        Me.GroupBox1.Controls.Add(Me.btnx)
        Me.GroupBox1.Controls.Add(Me.btndivide)
        Me.GroupBox1.Controls.Add(Me.btnmultiply)
        Me.GroupBox1.Controls.Add(Me.btnminus)
        Me.GroupBox1.Controls.Add(Me.btnadd)
        Me.GroupBox1.Controls.Add(Me.Button2)
        Me.GroupBox1.Controls.Add(Me.btndecimal)
        Me.GroupBox1.Controls.Add(Me.btn9)
        Me.GroupBox1.Controls.Add(Me.btn8)
        Me.GroupBox1.Controls.Add(Me.btn7)
        Me.GroupBox1.Controls.Add(Me.btn6)
        Me.GroupBox1.Controls.Add(Me.btn5)
        Me.GroupBox1.Controls.Add(Me.btn4)
        Me.GroupBox1.Controls.Add(Me.Btn3)
        Me.GroupBox1.Controls.Add(Me.Btn2)
        Me.GroupBox1.Controls.Add(Me.btn1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(261, 214)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'txtsource
        '
        Me.txtsource.Location = New System.Drawing.Point(13, 182)
        Me.txtsource.Name = "txtsource"
        Me.txtsource.Size = New System.Drawing.Size(231, 20)
        Me.txtsource.TabIndex = 19
        '
        'btnequal
        '
        Me.btnequal.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnequal.Location = New System.Drawing.Point(179, 133)
        Me.btnequal.Name = "btnequal"
        Me.btnequal.Size = New System.Drawing.Size(65, 32)
        Me.btnequal.TabIndex = 18
        Me.btnequal.Text = "="
        Me.btnequal.UseVisualStyleBackColor = True
        '
        'btnclear
        '
        Me.btnclear.Location = New System.Drawing.Point(109, 133)
        Me.btnclear.Name = "btnclear"
        Me.btnclear.Size = New System.Drawing.Size(64, 32)
        Me.btnclear.TabIndex = 17
        Me.btnclear.Text = "C"
        Me.btnclear.UseVisualStyleBackColor = True
        '
        'btnaddminus
        '
        Me.btnaddminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnaddminus.Location = New System.Drawing.Point(206, 95)
        Me.btnaddminus.Name = "btnaddminus"
        Me.btnaddminus.Size = New System.Drawing.Size(38, 32)
        Me.btnaddminus.TabIndex = 16
        Me.btnaddminus.Text = "+ -"
        Me.btnaddminus.UseVisualStyleBackColor = True
        '
        'btnx
        '
        Me.btnx.Location = New System.Drawing.Point(157, 95)
        Me.btnx.Name = "btnx"
        Me.btnx.Size = New System.Drawing.Size(38, 32)
        Me.btnx.TabIndex = 15
        Me.btnx.Text = "1/x"
        Me.btnx.UseVisualStyleBackColor = True
        '
        'btndivide
        '
        Me.btndivide.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndivide.Location = New System.Drawing.Point(206, 57)
        Me.btndivide.Name = "btndivide"
        Me.btndivide.Size = New System.Drawing.Size(38, 32)
        Me.btndivide.TabIndex = 14
        Me.btndivide.Text = "/"
        Me.btndivide.UseVisualStyleBackColor = True
        '
        'btnmultiply
        '
        Me.btnmultiply.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnmultiply.Location = New System.Drawing.Point(157, 57)
        Me.btnmultiply.Name = "btnmultiply"
        Me.btnmultiply.Size = New System.Drawing.Size(38, 32)
        Me.btnmultiply.TabIndex = 13
        Me.btnmultiply.Text = "*"
        Me.btnmultiply.UseVisualStyleBackColor = True
        '
        'btnminus
        '
        Me.btnminus.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnminus.Location = New System.Drawing.Point(206, 19)
        Me.btnminus.Name = "btnminus"
        Me.btnminus.Size = New System.Drawing.Size(38, 32)
        Me.btnminus.TabIndex = 12
        Me.btnminus.Text = "-"
        Me.btnminus.UseVisualStyleBackColor = True
        '
        'btnadd
        '
        Me.btnadd.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnadd.Location = New System.Drawing.Point(157, 19)
        Me.btnadd.Name = "btnadd"
        Me.btnadd.Size = New System.Drawing.Size(38, 32)
        Me.btnadd.TabIndex = 11
        Me.btnadd.Text = "+"
        Me.btnadd.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(62, 133)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(38, 32)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "0"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btndecimal
        '
        Me.btndecimal.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btndecimal.Location = New System.Drawing.Point(13, 133)
        Me.btndecimal.Name = "btndecimal"
        Me.btndecimal.Size = New System.Drawing.Size(38, 32)
        Me.btndecimal.TabIndex = 9
        Me.btndecimal.Text = "."
        Me.btndecimal.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btndecimal.UseVisualStyleBackColor = True
        '
        'btn9
        '
        Me.btn9.Location = New System.Drawing.Point(109, 95)
        Me.btn9.Name = "btn9"
        Me.btn9.Size = New System.Drawing.Size(38, 32)
        Me.btn9.TabIndex = 8
        Me.btn9.Text = "9"
        Me.btn9.UseVisualStyleBackColor = True
        '
        'btn8
        '
        Me.btn8.Location = New System.Drawing.Point(62, 95)
        Me.btn8.Name = "btn8"
        Me.btn8.Size = New System.Drawing.Size(38, 32)
        Me.btn8.TabIndex = 7
        Me.btn8.Text = "8"
        Me.btn8.UseVisualStyleBackColor = True
        '
        'btn7
        '
        Me.btn7.Location = New System.Drawing.Point(13, 95)
        Me.btn7.Name = "btn7"
        Me.btn7.Size = New System.Drawing.Size(38, 32)
        Me.btn7.TabIndex = 6
        Me.btn7.Text = "7"
        Me.btn7.UseVisualStyleBackColor = True
        '
        'btn6
        '
        Me.btn6.Location = New System.Drawing.Point(109, 57)
        Me.btn6.Name = "btn6"
        Me.btn6.Size = New System.Drawing.Size(38, 32)
        Me.btn6.TabIndex = 5
        Me.btn6.Text = "6"
        Me.btn6.UseVisualStyleBackColor = True
        '
        'btn5
        '
        Me.btn5.Location = New System.Drawing.Point(62, 57)
        Me.btn5.Name = "btn5"
        Me.btn5.Size = New System.Drawing.Size(38, 32)
        Me.btn5.TabIndex = 4
        Me.btn5.Text = "5"
        Me.btn5.UseVisualStyleBackColor = True
        '
        'btn4
        '
        Me.btn4.Location = New System.Drawing.Point(13, 57)
        Me.btn4.Name = "btn4"
        Me.btn4.Size = New System.Drawing.Size(38, 32)
        Me.btn4.TabIndex = 4
        Me.btn4.Text = "4"
        Me.btn4.UseVisualStyleBackColor = True
        '
        'Btn3
        '
        Me.Btn3.Location = New System.Drawing.Point(109, 19)
        Me.Btn3.Name = "Btn3"
        Me.Btn3.Size = New System.Drawing.Size(38, 32)
        Me.Btn3.TabIndex = 2
        Me.Btn3.Text = "3"
        Me.Btn3.UseVisualStyleBackColor = True
        '
        'Btn2
        '
        Me.Btn2.Location = New System.Drawing.Point(62, 19)
        Me.Btn2.Name = "Btn2"
        Me.Btn2.Size = New System.Drawing.Size(38, 32)
        Me.Btn2.TabIndex = 1
        Me.Btn2.Text = "2"
        Me.Btn2.UseVisualStyleBackColor = True
        '
        'btn1
        '
        Me.btn1.Location = New System.Drawing.Point(13, 19)
        Me.btn1.Name = "btn1"
        Me.btn1.Size = New System.Drawing.Size(38, 32)
        Me.btn1.TabIndex = 0
        Me.btn1.Text = "1"
        Me.btn1.UseVisualStyleBackColor = True
        '
        'frmcalculator
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(285, 236)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmcalculator"
        Me.Text = "Calculator"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents btnequal As Button
    Friend WithEvents btnclear As Button
    Friend WithEvents btnaddminus As Button
    Friend WithEvents btnx As Button
    Friend WithEvents btndivide As Button
    Friend WithEvents btnmultiply As Button
    Friend WithEvents btnminus As Button
    Friend WithEvents btnadd As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents btndecimal As Button
    Friend WithEvents btn9 As Button
    Friend WithEvents btn8 As Button
    Friend WithEvents btn7 As Button
    Friend WithEvents btn6 As Button
    Friend WithEvents btn5 As Button
    Friend WithEvents btn4 As Button
    Friend WithEvents Btn3 As Button
    Friend WithEvents Btn2 As Button
    Friend WithEvents btn1 As Button
    Friend WithEvents txtsource As TextBox
End Class
