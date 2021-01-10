<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMsg
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
        Me.Message = New System.Windows.Forms.Label()
        Me.picYes = New System.Windows.Forms.PictureBox()
        Me.picCritical = New System.Windows.Forms.PictureBox()
        Me.picDel = New System.Windows.Forms.PictureBox()
        Me.picEx = New System.Windows.Forms.PictureBox()
        Me.picInfo = New System.Windows.Forms.PictureBox()
        Me.gCancel = New System.Windows.Forms.Button()
        Me.gOK = New System.Windows.Forms.Button()
        Me.gNo = New System.Windows.Forms.Button()
        Me.gYes = New System.Windows.Forms.Button()
        CType(Me.picYes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCritical, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picDel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picEx, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Message
        '
        Me.Message.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Message.ForeColor = System.Drawing.Color.Maroon
        Me.Message.Location = New System.Drawing.Point(182, 32)
        Me.Message.Name = "Message"
        Me.Message.Size = New System.Drawing.Size(432, 95)
        Me.Message.TabIndex = 67
        Me.Message.Text = "Header"
        Me.Message.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'picYes
        '
        Me.picYes.Image = Global.Smartplant_Data_Tool.My.Resources.Resources.question_mark_icon
        Me.picYes.Location = New System.Drawing.Point(12, 14)
        Me.picYes.Name = "picYes"
        Me.picYes.Size = New System.Drawing.Size(131, 113)
        Me.picYes.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picYes.TabIndex = 71
        Me.picYes.TabStop = False
        Me.picYes.Visible = False
        '
        'picCritical
        '
        Me.picCritical.Image = Global.Smartplant_Data_Tool.My.Resources.Resources.Exclamationmark_icon
        Me.picCritical.Location = New System.Drawing.Point(12, 14)
        Me.picCritical.Name = "picCritical"
        Me.picCritical.Size = New System.Drawing.Size(131, 113)
        Me.picCritical.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picCritical.TabIndex = 70
        Me.picCritical.TabStop = False
        Me.picCritical.Visible = False
        '
        'picDel
        '
        Me.picDel.Image = Global.Smartplant_Data_Tool.My.Resources.Resources.delete2
        Me.picDel.Location = New System.Drawing.Point(12, 14)
        Me.picDel.Name = "picDel"
        Me.picDel.Size = New System.Drawing.Size(131, 113)
        Me.picDel.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picDel.TabIndex = 69
        Me.picDel.TabStop = False
        Me.picDel.Visible = False
        '
        'picEx
        '
        Me.picEx.Image = Global.Smartplant_Data_Tool.My.Resources.Resources.exclamationmark
        Me.picEx.Location = New System.Drawing.Point(12, 14)
        Me.picEx.Name = "picEx"
        Me.picEx.Size = New System.Drawing.Size(131, 113)
        Me.picEx.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picEx.TabIndex = 68
        Me.picEx.TabStop = False
        Me.picEx.Visible = False
        '
        'picInfo
        '
        Me.picInfo.Image = Global.Smartplant_Data_Tool.My.Resources.Resources.exclamationmark
        Me.picInfo.Location = New System.Drawing.Point(12, 14)
        Me.picInfo.Name = "picInfo"
        Me.picInfo.Size = New System.Drawing.Size(131, 113)
        Me.picInfo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picInfo.TabIndex = 72
        Me.picInfo.TabStop = False
        Me.picInfo.Visible = False
        '
        'gCancel
        '
        Me.gCancel.Location = New System.Drawing.Point(472, 140)
        Me.gCancel.Name = "gCancel"
        Me.gCancel.Size = New System.Drawing.Size(142, 49)
        Me.gCancel.TabIndex = 73
        Me.gCancel.Text = "Cancel"
        Me.gCancel.UseVisualStyleBackColor = True
        '
        'gOK
        '
        Me.gOK.Location = New System.Drawing.Point(324, 140)
        Me.gOK.Name = "gOK"
        Me.gOK.Size = New System.Drawing.Size(142, 49)
        Me.gOK.TabIndex = 74
        Me.gOK.Text = "OK"
        Me.gOK.UseVisualStyleBackColor = True
        '
        'gNo
        '
        Me.gNo.Location = New System.Drawing.Point(472, 140)
        Me.gNo.Name = "gNo"
        Me.gNo.Size = New System.Drawing.Size(142, 49)
        Me.gNo.TabIndex = 75
        Me.gNo.Text = "No"
        Me.gNo.UseVisualStyleBackColor = True
        '
        'gYes
        '
        Me.gYes.Location = New System.Drawing.Point(324, 140)
        Me.gYes.Name = "gYes"
        Me.gYes.Size = New System.Drawing.Size(142, 49)
        Me.gYes.TabIndex = 76
        Me.gYes.Text = "Yes"
        Me.gYes.UseVisualStyleBackColor = True
        '
        'frmMsg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(626, 201)
        Me.ControlBox = False
        Me.Controls.Add(Me.gYes)
        Me.Controls.Add(Me.gNo)
        Me.Controls.Add(Me.gOK)
        Me.Controls.Add(Me.gCancel)
        Me.Controls.Add(Me.picInfo)
        Me.Controls.Add(Me.picYes)
        Me.Controls.Add(Me.picCritical)
        Me.Controls.Add(Me.picDel)
        Me.Controls.Add(Me.picEx)
        Me.Controls.Add(Me.Message)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frmMsg"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.picYes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCritical, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picDel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picEx, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Message As System.Windows.Forms.Label
    Friend WithEvents picEx As System.Windows.Forms.PictureBox
    Friend WithEvents picDel As System.Windows.Forms.PictureBox
    Friend WithEvents picCritical As System.Windows.Forms.PictureBox
    Friend WithEvents picYes As System.Windows.Forms.PictureBox
    Friend WithEvents picInfo As System.Windows.Forms.PictureBox
    Friend WithEvents gCancel As System.Windows.Forms.Button
    Friend WithEvents gOK As System.Windows.Forms.Button
    Friend WithEvents gNo As System.Windows.Forms.Button
    Friend WithEvents gYes As System.Windows.Forms.Button
End Class
