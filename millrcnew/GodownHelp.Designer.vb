<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GodownHelp
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GodownHelp))
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.TxtSrch = New System.Windows.Forms.TextBox()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView3
        '
        Me.DataGridView3.AllowUserToAddRows = False
        Me.DataGridView3.AllowUserToDeleteRows = False
        Me.DataGridView3.AllowUserToResizeColumns = False
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView3.Location = New System.Drawing.Point(12, 12)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(446, 412)
        Me.DataGridView3.TabIndex = 106
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.TxtSrch)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 430)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(200, 50)
        Me.GroupBox5.TabIndex = 114
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Search by tenant name"
        '
        'TxtSrch
        '
        Me.TxtSrch.Location = New System.Drawing.Point(12, 19)
        Me.TxtSrch.Name = "TxtSrch"
        Me.TxtSrch.Size = New System.Drawing.Size(177, 20)
        Me.TxtSrch.TabIndex = 0
        '
        'GodownHelp
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(470, 486)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.DataGridView3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "GodownHelp"
        Me.Text = "GodownHelp"
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView3 As DataGridView
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents TxtSrch As TextBox
End Class
