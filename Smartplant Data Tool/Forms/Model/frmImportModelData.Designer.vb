<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportModelData
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
        Me.pp = New DevExpress.XtraWaitForm.ProgressPanel()
        Me.pbAll = New DevExpress.XtraEditors.ProgressBarControl()
        Me.LabelControl1 = New DevExpress.XtraEditors.LabelControl()
        Me.lblAllProgress = New DevExpress.XtraEditors.LabelControl()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.SimpleButton2 = New DevExpress.XtraEditors.SimpleButton()
        Me.lblModelName = New DevExpress.XtraEditors.LabelControl()
        CType(Me.pbAll.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pp
        '
        Me.pp.Appearance.BackColor = System.Drawing.Color.Transparent
        Me.pp.Appearance.Options.UseBackColor = True
        Me.pp.Location = New System.Drawing.Point(387, 323)
        Me.pp.Name = "pp"
        Me.pp.Size = New System.Drawing.Size(199, 81)
        Me.pp.TabIndex = 0
        Me.pp.Text = "ProgressPanel1"
        Me.pp.Visible = False
        '
        'pbAll
        '
        Me.pbAll.Location = New System.Drawing.Point(168, 237)
        Me.pbAll.Name = "pbAll"
        Me.pbAll.Properties.EndColor = System.Drawing.Color.Black
        Me.pbAll.Properties.Maximum = 4
        Me.pbAll.Properties.StartColor = System.Drawing.Color.Gainsboro
        Me.pbAll.Size = New System.Drawing.Size(636, 39)
        Me.pbAll.TabIndex = 1
        '
        'LabelControl1
        '
        Me.LabelControl1.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.LabelControl1.Appearance.ForeColor = System.Drawing.Color.DodgerBlue
        Me.LabelControl1.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.LabelControl1.LineColor = System.Drawing.Color.DodgerBlue
        Me.LabelControl1.LineLocation = DevExpress.XtraEditors.LineLocation.Bottom
        Me.LabelControl1.LineVisible = True
        Me.LabelControl1.Location = New System.Drawing.Point(12, 33)
        Me.LabelControl1.Name = "LabelControl1"
        Me.LabelControl1.Size = New System.Drawing.Size(636, 62)
        Me.LabelControl1.TabIndex = 2
        Me.LabelControl1.Text = "Importing Model to Local Database"
        '
        'lblAllProgress
        '
        Me.lblAllProgress.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.lblAllProgress.Appearance.ForeColor = System.Drawing.Color.CadetBlue
        Me.lblAllProgress.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center
        Me.lblAllProgress.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblAllProgress.LineColor = System.Drawing.Color.CadetBlue
        Me.lblAllProgress.LineLocation = DevExpress.XtraEditors.LineLocation.Top
        Me.lblAllProgress.LineVisible = True
        Me.lblAllProgress.Location = New System.Drawing.Point(168, 282)
        Me.lblAllProgress.Name = "lblAllProgress"
        Me.lblAllProgress.Size = New System.Drawing.Size(636, 36)
        Me.lblAllProgress.TabIndex = 3
        Me.lblAllProgress.Text = "Overall Progress: 0"
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!)
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Location = New System.Drawing.Point(775, 569)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(169, 53)
        Me.SimpleButton1.TabIndex = 13
        Me.SimpleButton1.Text = "Close"
        '
        'SimpleButton2
        '
        Me.SimpleButton2.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!)
        Me.SimpleButton2.Appearance.Options.UseFont = True
        Me.SimpleButton2.Location = New System.Drawing.Point(600, 569)
        Me.SimpleButton2.Name = "SimpleButton2"
        Me.SimpleButton2.Size = New System.Drawing.Size(169, 53)
        Me.SimpleButton2.TabIndex = 14
        Me.SimpleButton2.Text = "Start"
        '
        'lblModelName
        '
        Me.lblModelName.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.lblModelName.Appearance.ForeColor = System.Drawing.Color.Black
        Me.lblModelName.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblModelName.LineColor = System.Drawing.Color.DodgerBlue
        Me.lblModelName.LineLocation = DevExpress.XtraEditors.LineLocation.Bottom
        Me.lblModelName.Location = New System.Drawing.Point(33, 96)
        Me.lblModelName.Name = "lblModelName"
        Me.lblModelName.Size = New System.Drawing.Size(851, 46)
        Me.lblModelName.TabIndex = 15
        Me.lblModelName.Text = "Model:"
        '
        'frmImportModelData
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(970, 665)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblModelName)
        Me.Controls.Add(Me.SimpleButton2)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.lblAllProgress)
        Me.Controls.Add(Me.LabelControl1)
        Me.Controls.Add(Me.pbAll)
        Me.Controls.Add(Me.pp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmImportModelData"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        CType(Me.pbAll.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents pp As DevExpress.XtraWaitForm.ProgressPanel
    Friend WithEvents pbAll As DevExpress.XtraEditors.ProgressBarControl
    Friend WithEvents LabelControl1 As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblAllProgress As DevExpress.XtraEditors.LabelControl
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents SimpleButton2 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents lblModelName As DevExpress.XtraEditors.LabelControl
End Class
