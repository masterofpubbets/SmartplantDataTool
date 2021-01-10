<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAppend
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAppend))
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.opnFle = New System.Windows.Forms.OpenFileDialog()
        Me.SimpleButton2 = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.LabelControl2 = New DevExpress.XtraEditors.LabelControl()
        Me.SimpleButton3 = New DevExpress.XtraEditors.SimpleButton()
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
        Me.LabelControl1.Location = New System.Drawing.Point(12, 29)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(636, 62)
        Me.LabelControl1.TabIndex = 1
        Me.LabelControl1.Text = "Append New Smartplant Model to Local Database"
        '
        'opnFle
        '
        Me.opnFle.Filter = "Smartplant File|*.vue"
        '
        'SimpleButton2
        '
        Me.SimpleButton2.Appearance.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.SimpleButton2.Appearance.Options.UseFont = True
        Me.SimpleButton2.Location = New System.Drawing.Point(55, 121)
        Me.SimpleButton2.Name = "SimpleButton2"
        Me.SimpleButton2.Size = New System.Drawing.Size(134, 40)
        Me.SimpleButton2.TabIndex = 5
        Me.SimpleButton2.Text = "Select Model"
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!)
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Location = New System.Drawing.Point(817, 568)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(169, 53)
        Me.SimpleButton1.TabIndex = 6
        Me.SimpleButton1.Text = "Close"
        '
        'lblInfo
        '
        Me.lblInfo.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.lblInfo.ForeColor = System.Drawing.Color.DimGray
        Me.lblInfo.Location = New System.Drawing.Point(125, 227)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(737, 190)
        Me.lblInfo.TabIndex = 7
        Me.lblInfo.Text = "No Model"
        '
        'LabelControl2
        '
        Me.LabelControl2.Appearance.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.LabelControl2.Appearance.ForeColor = System.Drawing.Color.DimGray
        Me.LabelControl2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl2.LineColor = System.Drawing.Color.DimGray
        Me.LabelControl2.LineLocation = DevExpress.XtraEditors.LineLocation.Center
        Me.LabelControl2.LineVisible = True
        Me.LabelControl2.Location = New System.Drawing.Point(55, 194)
        Me.LabelControl2.Name = "LabelControl2"
        Me.LabelControl2.Size = New System.Drawing.Size(807, 30)
        Me.LabelControl2.TabIndex = 8
        Me.LabelControl2.Text = "Model Info"
        '
        'SimpleButton3
        '
        Me.SimpleButton3.Appearance.Font = New System.Drawing.Font("Tahoma", 9.0!)
        Me.SimpleButton3.Appearance.Options.UseFont = True
        Me.SimpleButton3.Image = CType(resources.GetObject("SimpleButton3.Image"), System.Drawing.Image)
        Me.SimpleButton3.Location = New System.Drawing.Point(128, 430)
        Me.SimpleButton3.Name = "SimpleButton3"
        Me.SimpleButton3.Size = New System.Drawing.Size(134, 40)
        Me.SimpleButton3.TabIndex = 9
        Me.SimpleButton3.Text = "Append"
        '
        'pp
        '
        Me.pp.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.pp.Appearance.Options.UseBackColor = True
        Me.pp.Location = New System.Drawing.Point(408, 304)
        Me.pp.Name = "pp"
        Me.pp.Size = New System.Drawing.Size(199, 81)
        Me.pp.TabIndex = 10
        Me.pp.Text = "ProgressPanel1"
        Me.pp.Visible = False
        '
        'frmAppend
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1038, 665)
        Me.ControlBox = False
        Me.Controls.Add(Me.pp)
        Me.Controls.Add(Me.SimpleButton3)
        Me.Controls.Add(Me.LabelControl2)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.SimpleButton2)
        Me.Controls.Add(Me.LabelControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmAppend"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Append New Model"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents opnFle As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SimpleButton2 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents LabelControl2 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SimpleButton3 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents pp As DevExpress.XtraWaitForm.ProgressPanel
End Class
