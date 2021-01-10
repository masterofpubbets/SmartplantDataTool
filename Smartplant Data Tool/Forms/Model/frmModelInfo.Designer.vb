<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmModelInfo
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
        Me.tree = New System.Windows.Forms.TreeView()
        Me.SimpleButton1 = New DevExpress.XtraEditors.SimpleButton()
        Me.lblModelName = New DevExpress.XtraEditors.LabelControl()
        Me.SuspendLayout()
        '
        'tree
        '
        Me.tree.Location = New System.Drawing.Point(28, 78)
        Me.tree.Name = "tree"
        Me.tree.Size = New System.Drawing.Size(593, 440)
        Me.tree.TabIndex = 0
        '
        'SimpleButton1
        '
        Me.SimpleButton1.Appearance.Font = New System.Drawing.Font("Tahoma", 11.0!)
        Me.SimpleButton1.Appearance.Options.UseFont = True
        Me.SimpleButton1.Location = New System.Drawing.Point(477, 541)
        Me.SimpleButton1.Name = "SimpleButton1"
        Me.SimpleButton1.Size = New System.Drawing.Size(169, 53)
        Me.SimpleButton1.TabIndex = 14
        Me.SimpleButton1.Text = "Close"
        '
        'lblModelName
        '
        Me.lblModelName.Appearance.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.lblModelName.Appearance.ForeColor = System.Drawing.Color.DodgerBlue
        Me.lblModelName.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None
        Me.lblModelName.LineColor = System.Drawing.Color.DodgerBlue
        Me.lblModelName.LineLocation = DevExpress.XtraEditors.LineLocation.Bottom
        Me.lblModelName.LineVisible = True
        Me.lblModelName.Location = New System.Drawing.Point(10, 12)
        Me.lblModelName.Name = "lblModelName"
        Me.lblModelName.Size = New System.Drawing.Size(636, 50)
        Me.lblModelName.TabIndex = 15
        Me.lblModelName.Text = "Model: "
        '
        'frmModelInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(235, Byte), Integer), CType(CType(236, Byte), Integer), CType(CType(239, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(671, 638)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblModelName)
        Me.Controls.Add(Me.SimpleButton1)
        Me.Controls.Add(Me.tree)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmModelInfo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Model Info"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tree As System.Windows.Forms.TreeView
    Friend WithEvents SimpleButton1 As DevExpress.XtraEditors.SimpleButton
    Friend WithEvents lblModelName As DevExpress.XtraEditors.LabelControl
End Class
