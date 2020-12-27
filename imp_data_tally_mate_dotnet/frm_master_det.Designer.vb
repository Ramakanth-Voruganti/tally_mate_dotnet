<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_master_det
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
        Me.tab_import_main = New System.Windows.Forms.TabControl()
        Me.tab_import_master = New System.Windows.Forms.TabPage()
        Me.btn_get_master_det = New System.Windows.Forms.Button()
        Me.dgrid_master_actual = New System.Windows.Forms.DataGridView()
        Me.tab_test = New System.Windows.Forms.TabPage()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.pbr_frm_import_sales_to_tally = New System.Windows.Forms.ProgressBar()
        Me.btn_post_xml = New System.Windows.Forms.Button()
        Me.lbl_name = New System.Windows.Forms.Label()
        Me.txt_name = New System.Windows.Forms.TextBox()
        Me.txt_address = New System.Windows.Forms.TextBox()
        Me.lbl_address = New System.Windows.Forms.Label()
        Me.txt_group = New System.Windows.Forms.TextBox()
        Me.lbl_group = New System.Windows.Forms.Label()
        Me.tab_import_main.SuspendLayout()
        Me.tab_import_master.SuspendLayout()
        CType(Me.dgrid_master_actual, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tab_test.SuspendLayout()
        Me.SuspendLayout()
        '
        'tab_import_main
        '
        Me.tab_import_main.Appearance = System.Windows.Forms.TabAppearance.Buttons
        Me.tab_import_main.Controls.Add(Me.tab_import_master)
        Me.tab_import_main.Controls.Add(Me.tab_test)
        Me.tab_import_main.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tab_import_main.Location = New System.Drawing.Point(12, 83)
        Me.tab_import_main.Name = "tab_import_main"
        Me.tab_import_main.Padding = New System.Drawing.Point(100, 2)
        Me.tab_import_main.SelectedIndex = 0
        Me.tab_import_main.Size = New System.Drawing.Size(983, 477)
        Me.tab_import_main.TabIndex = 5
        '
        'tab_import_master
        '
        Me.tab_import_master.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tab_import_master.Controls.Add(Me.btn_get_master_det)
        Me.tab_import_master.Controls.Add(Me.dgrid_master_actual)
        Me.tab_import_master.Font = New System.Drawing.Font("Courier New", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tab_import_master.Location = New System.Drawing.Point(4, 28)
        Me.tab_import_master.Name = "tab_import_master"
        Me.tab_import_master.Padding = New System.Windows.Forms.Padding(3)
        Me.tab_import_master.Size = New System.Drawing.Size(975, 445)
        Me.tab_import_master.TabIndex = 0
        Me.tab_import_master.Text = "Master"
        Me.tab_import_master.UseVisualStyleBackColor = True
        '
        'btn_get_master_det
        '
        Me.btn_get_master_det.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_get_master_det.Location = New System.Drawing.Point(17, 6)
        Me.btn_get_master_det.Name = "btn_get_master_det"
        Me.btn_get_master_det.Size = New System.Drawing.Size(141, 39)
        Me.btn_get_master_det.TabIndex = 8
        Me.btn_get_master_det.Text = "Get Details From Tally"
        Me.btn_get_master_det.UseVisualStyleBackColor = True
        '
        'dgrid_master_actual
        '
        Me.dgrid_master_actual.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgrid_master_actual.Location = New System.Drawing.Point(17, 63)
        Me.dgrid_master_actual.Name = "dgrid_master_actual"
        Me.dgrid_master_actual.Size = New System.Drawing.Size(935, 358)
        Me.dgrid_master_actual.TabIndex = 0
        '
        'tab_test
        '
        Me.tab_test.Controls.Add(Me.TextBox1)
        Me.tab_test.Location = New System.Drawing.Point(4, 28)
        Me.tab_test.Name = "tab_test"
        Me.tab_test.Size = New System.Drawing.Size(975, 445)
        Me.tab_test.TabIndex = 1
        Me.tab_test.Text = "Test"
        Me.tab_test.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.TextBox1.ForeColor = System.Drawing.Color.White
        Me.TextBox1.Location = New System.Drawing.Point(12, 14)
        Me.TextBox1.MaxLength = 0
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.TextBox1.Size = New System.Drawing.Size(945, 409)
        Me.TextBox1.TabIndex = 6
        '
        'pbr_frm_import_sales_to_tally
        '
        Me.pbr_frm_import_sales_to_tally.Location = New System.Drawing.Point(869, 25)
        Me.pbr_frm_import_sales_to_tally.Name = "pbr_frm_import_sales_to_tally"
        Me.pbr_frm_import_sales_to_tally.Size = New System.Drawing.Size(145, 16)
        Me.pbr_frm_import_sales_to_tally.TabIndex = 8
        '
        'btn_post_xml
        '
        Me.btn_post_xml.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_post_xml.Location = New System.Drawing.Point(778, 9)
        Me.btn_post_xml.Name = "btn_post_xml"
        Me.btn_post_xml.Size = New System.Drawing.Size(70, 39)
        Me.btn_post_xml.TabIndex = 6
        Me.btn_post_xml.Text = "Post XML"
        Me.btn_post_xml.UseVisualStyleBackColor = True
        '
        'lbl_name
        '
        Me.lbl_name.AutoSize = True
        Me.lbl_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_name.Location = New System.Drawing.Point(13, 9)
        Me.lbl_name.Name = "lbl_name"
        Me.lbl_name.Size = New System.Drawing.Size(87, 15)
        Me.lbl_name.TabIndex = 9
        Me.lbl_name.Text = "Account Name"
        '
        'txt_name
        '
        Me.txt_name.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_name.Location = New System.Drawing.Point(16, 31)
        Me.txt_name.Name = "txt_name"
        Me.txt_name.Size = New System.Drawing.Size(269, 21)
        Me.txt_name.TabIndex = 10
        '
        'txt_address
        '
        Me.txt_address.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_address.Location = New System.Drawing.Point(309, 31)
        Me.txt_address.Name = "txt_address"
        Me.txt_address.Size = New System.Drawing.Size(203, 21)
        Me.txt_address.TabIndex = 12
        '
        'lbl_address
        '
        Me.lbl_address.AutoSize = True
        Me.lbl_address.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_address.Location = New System.Drawing.Point(306, 9)
        Me.lbl_address.Name = "lbl_address"
        Me.lbl_address.Size = New System.Drawing.Size(51, 15)
        Me.lbl_address.TabIndex = 11
        Me.lbl_address.Text = "Address"
        '
        'txt_group
        '
        Me.txt_group.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_group.Location = New System.Drawing.Point(533, 31)
        Me.txt_group.Name = "txt_group"
        Me.txt_group.Size = New System.Drawing.Size(203, 21)
        Me.txt_group.TabIndex = 14
        '
        'lbl_group
        '
        Me.lbl_group.AutoSize = True
        Me.lbl_group.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_group.Location = New System.Drawing.Point(530, 9)
        Me.lbl_group.Name = "lbl_group"
        Me.lbl_group.Size = New System.Drawing.Size(126, 15)
        Me.lbl_group.TabIndex = 13
        Me.lbl_group.Text = "Account Group / Head"
        '
        'frm_master_det
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1041, 574)
        Me.Controls.Add(Me.txt_group)
        Me.Controls.Add(Me.lbl_group)
        Me.Controls.Add(Me.txt_address)
        Me.Controls.Add(Me.lbl_address)
        Me.Controls.Add(Me.txt_name)
        Me.Controls.Add(Me.lbl_name)
        Me.Controls.Add(Me.tab_import_main)
        Me.Controls.Add(Me.pbr_frm_import_sales_to_tally)
        Me.Controls.Add(Me.btn_post_xml)
        Me.MaximizeBox = False
        Me.Name = "frm_master_det"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frm_master_det"
        Me.tab_import_main.ResumeLayout(False)
        Me.tab_import_master.ResumeLayout(False)
        CType(Me.dgrid_master_actual, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tab_test.ResumeLayout(False)
        Me.tab_test.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tab_import_main As TabControl
    Friend WithEvents tab_import_master As TabPage
    Friend WithEvents dgrid_master_actual As DataGridView
    Friend WithEvents tab_test As TabPage
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents pbr_frm_import_sales_to_tally As ProgressBar
    Friend WithEvents btn_post_xml As Button
    Friend WithEvents btn_get_master_det As Button
    Friend WithEvents lbl_name As Label
    Friend WithEvents txt_name As TextBox
    Friend WithEvents txt_address As TextBox
    Friend WithEvents lbl_address As Label
    Friend WithEvents txt_group As TextBox
    Friend WithEvents lbl_group As Label
End Class
