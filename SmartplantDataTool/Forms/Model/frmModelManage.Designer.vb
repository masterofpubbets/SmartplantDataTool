<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModelManage
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmModelManage))
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.lstModel = New System.Windows.Forms.ListBox()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButton2 = New DevExpress.XtraEditors.SimpleButton()
        Me.pp = New DevExpress.XtraWaitForm.ProgressPanel()
        Me.SuspendLayout()
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.DodgerBlue
        Me.LabelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl1.LineColor = System.Drawing.Color.DodgerBlue
        Me.LabelControl1.LineLocation = DevExpress.XtraEditors.LineLocation.Bottom
        Me.LabelControl1.LineVisible = True
        Me.LabelControl1.Location = New System.Drawing.Point(12, 41)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(636, 62)
        Me.LabelControl1.TabIndex = 3
        Me.LabelControl1.Text = "Existing Models"
        '
        'lstModel
        '
        Me.lstModel.FormattingEnabled = True
        Me.lstModel.ItemHeight = 16
        Me.lstModel.Location = New System.Drawing.Point(75, 109)
        Me.lstModel.Name = "lstModel"
        Me.lstModel.Size = New System.Drawing.Size(413, 340)
        Me.lstModel.TabIndex = 4
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!)
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Location = New System.Drawing.Point(590, 506)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(169, 53)
        Me.SimpleButton1.TabIndex = 15
        Me.SimpleButton1.Text = "Close"
        '
        'SimpleButton2
        '
        Me.SimpleButton2.Appearance.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.SimpleButton2.Appearance.Options.UseFont = True
        Me.SimpleButton2.Image = CType(resources.GetObject("SimpleButton2.Image"), System.Drawing.Image)
        Me.SimpleButton2.Location = New System.Drawing.Point(494, 158)
        Me.SimpleButton2.Name = "SimpleButton2"
        Me.SimpleButton2.Size = New System.Drawing.Size(126, 37)
        Me.SimpleButton2.TabIndex = 16
        Me.SimpleButton2.Text = "Delete"
        '
        'pp
        '
        Me.pp.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.pp.Appearance.Options.UseBackColor = True
        Me.pp.Location = New System.Drawing.Point(494, 262)
        Me.pp.Name = "pp"
        Me.pp.Size = New System.Drawing.Size(199, 81)
        Me.pp.TabIndex = 17
        Me.pp.Text = "ProgressPanel1"
        Me.pp.Visible = False
        '
        'frmModelManage
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(786, 619)
        Me.ControlBox = False
        Me.Controls.Add(Me.pp)
        Me.Controls.Add(Me.SimpleButton2)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.lstModel)
        Me.Controls.Add(Me.LabelControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmModelManage"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Model Management"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lstModel As System.Windows.Forms.ListBox
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButton2 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents pp As DevExpress.XtraWaitForm.ProgressPanel
End Class
