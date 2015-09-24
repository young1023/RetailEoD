<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DailySales
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DailySales))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Button_Previous = New System.Windows.Forms.Button()
        Me.Button_Next = New System.Windows.Forms.Button()
        Me.Combobox_Station = New System.Windows.Forms.ComboBox()
        Me.PageNo1 = New System.Windows.Forms.TextBox()
        Me.Tran_Search_Text1 = New System.Windows.Forms.TextBox()
        Me.Convert_CSV = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.DataConvertionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TransactionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExcelShopToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExportedFilesTrackingToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Search_Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.From = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Rows_No = New System.Windows.Forms.ComboBox()
        Me.Button_First = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Initial_Value = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(14, 148)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowTemplate.Height = 24
        Me.DataGridView1.Size = New System.Drawing.Size(1106, 408)
        Me.DataGridView1.TabIndex = 0
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(75, 51)
        Me.DateTimePicker1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(142, 23)
        Me.DateTimePicker1.TabIndex = 1
        Me.DateTimePicker1.Value = New Date(2014, 2, 27, 0, 0, 0, 0)
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(301, 51)
        Me.DateTimePicker2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(136, 23)
        Me.DateTimePicker2.TabIndex = 2
        Me.DateTimePicker2.Value = New Date(2014, 2, 28, 0, 0, 0, 0)
        '
        'Button_Previous
        '
        Me.Button_Previous.Location = New System.Drawing.Point(327, 111)
        Me.Button_Previous.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button_Previous.Name = "Button_Previous"
        Me.Button_Previous.Size = New System.Drawing.Size(38, 29)
        Me.Button_Previous.TabIndex = 3
        Me.Button_Previous.Text = "<"
        Me.Button_Previous.UseVisualStyleBackColor = True
        '
        'Button_Next
        '
        Me.Button_Next.Location = New System.Drawing.Point(401, 110)
        Me.Button_Next.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button_Next.Name = "Button_Next"
        Me.Button_Next.Size = New System.Drawing.Size(38, 29)
        Me.Button_Next.TabIndex = 4
        Me.Button_Next.Text = ">"
        Me.Button_Next.UseVisualStyleBackColor = True
        '
        'Combobox_Station
        '
        Me.Combobox_Station.FormattingEnabled = True
        Me.Combobox_Station.Items.AddRange(New Object() {"ALL", "134", "136", "185", "187", "223", "224", "225", "226", "227", "229", "230", "231", "233", "243", "250", "255", "256", "262", "265", "301", "302", "303", "307", "313", "320", "325", "328", "333", "341", "342", "347", "356", "357", "366", "369", "372", "373", "374", "386", "388", "399", "402", "403", "408", "409", "822", "888"})
        Me.Combobox_Station.Location = New System.Drawing.Point(574, 111)
        Me.Combobox_Station.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Combobox_Station.Name = "Combobox_Station"
        Me.Combobox_Station.Size = New System.Drawing.Size(125, 23)
        Me.Combobox_Station.TabIndex = 5
        '
        'PageNo1
        '
        Me.PageNo1.Location = New System.Drawing.Point(75, 108)
        Me.PageNo1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.PageNo1.Name = "PageNo1"
        Me.PageNo1.Size = New System.Drawing.Size(47, 23)
        Me.PageNo1.TabIndex = 6
        Me.PageNo1.Text = "1"
        '
        'Tran_Search_Text1
        '
        Me.Tran_Search_Text1.Location = New System.Drawing.Point(574, 51)
        Me.Tran_Search_Text1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Tran_Search_Text1.Name = "Tran_Search_Text1"
        Me.Tran_Search_Text1.Size = New System.Drawing.Size(137, 23)
        Me.Tran_Search_Text1.TabIndex = 7
        '
        'Convert_CSV
        '
        Me.Convert_CSV.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Convert_CSV.Location = New System.Drawing.Point(868, 50)
        Me.Convert_CSV.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Convert_CSV.Name = "Convert_CSV"
        Me.Convert_CSV.Size = New System.Drawing.Size(87, 29)
        Me.Convert_CSV.TabIndex = 8
        Me.Convert_CSV.Text = "CSV"
        Me.Convert_CSV.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DataConvertionToolStripMenuItem, Me.ReportToolStripMenuItem, Me.ExportedFilesTrackingToolStripMenuItem, Me.ExitToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(7, 2, 0, 2)
        Me.MenuStrip1.Size = New System.Drawing.Size(1134, 24)
        Me.MenuStrip1.TabIndex = 9
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'DataConvertionToolStripMenuItem
        '
        Me.DataConvertionToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.DataConvertionToolStripMenuItem.Name = "DataConvertionToolStripMenuItem"
        Me.DataConvertionToolStripMenuItem.Size = New System.Drawing.Size(105, 20)
        Me.DataConvertionToolStripMenuItem.Text = "Data Convertion"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TransactionToolStripMenuItem, Me.SalesToolStripMenuItem, Me.ExcelShopToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'TransactionToolStripMenuItem
        '
        Me.TransactionToolStripMenuItem.Name = "TransactionToolStripMenuItem"
        Me.TransactionToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.TransactionToolStripMenuItem.Text = "Transaction Detail"
        '
        'SalesToolStripMenuItem
        '
        Me.SalesToolStripMenuItem.Name = "SalesToolStripMenuItem"
        Me.SalesToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.SalesToolStripMenuItem.Text = "Sales"
        '
        'ExcelShopToolStripMenuItem
        '
        Me.ExcelShopToolStripMenuItem.Name = "ExcelShopToolStripMenuItem"
        Me.ExcelShopToolStripMenuItem.Size = New System.Drawing.Size(169, 22)
        Me.ExcelShopToolStripMenuItem.Text = "Excel Shop"
        '
        'ExportedFilesTrackingToolStripMenuItem
        '
        Me.ExportedFilesTrackingToolStripMenuItem.Name = "ExportedFilesTrackingToolStripMenuItem"
        Me.ExportedFilesTrackingToolStripMenuItem.Size = New System.Drawing.Size(140, 20)
        Me.ExportedFilesTrackingToolStripMenuItem.Text = "Exported Files Tracking"
        '
        'ExitToolStripMenuItem1
        '
        Me.ExitToolStripMenuItem1.Name = "ExitToolStripMenuItem1"
        Me.ExitToolStripMenuItem1.Size = New System.Drawing.Size(37, 20)
        Me.ExitToolStripMenuItem1.Text = "Exit"
        '
        'Search_Button1
        '
        Me.Search_Button1.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Search_Button1.Location = New System.Drawing.Point(750, 49)
        Me.Search_Button1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Search_Button1.Name = "Search_Button1"
        Me.Search_Button1.Size = New System.Drawing.Size(87, 29)
        Me.Search_Button1.TabIndex = 11
        Me.Search_Button1.Text = "Search"
        Me.Search_Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 111)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(33, 15)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Page"
        '
        'From
        '
        Me.From.AutoSize = True
        Me.From.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.From.Location = New System.Drawing.Point(14, 61)
        Me.From.Name = "From"
        Me.From.Size = New System.Drawing.Size(35, 15)
        Me.From.TabIndex = 20
        Me.From.Text = "From"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(464, 115)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 15)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Station"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(234, 61)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(21, 15)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Symbol", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(875, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 15)
        Me.Label3.TabIndex = 27
        Me.Label3.Text = "Rows Per Page"
        '
        'Rows_No
        '
        Me.Rows_No.FormattingEnabled = True
        Me.Rows_No.Items.AddRange(New Object() {"10", "20", "25", "30", "35"})
        Me.Rows_No.Location = New System.Drawing.Point(750, 114)
        Me.Rows_No.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Rows_No.Name = "Rows_No"
        Me.Rows_No.Size = New System.Drawing.Size(101, 23)
        Me.Rows_No.TabIndex = 28
        Me.Rows_No.Text = "10"
        '
        'Button_First
        '
        Me.Button_First.Location = New System.Drawing.Point(269, 111)
        Me.Button_First.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Button_First.Name = "Button_First"
        Me.Button_First.Size = New System.Drawing.Size(38, 29)
        Me.Button_First.TabIndex = 29
        Me.Button_First.Text = "<<"
        Me.Button_First.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(464, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(83, 15)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Transaction ID"
        '
        'Initial_Value
        '
        Me.Initial_Value.AutoSize = True
        Me.Initial_Value.Location = New System.Drawing.Point(17, 28)
        Me.Initial_Value.Name = "Initial_Value"
        Me.Initial_Value.Size = New System.Drawing.Size(13, 15)
        Me.Initial_Value.TabIndex = 31
        Me.Initial_Value.Text = "0"
        Me.Initial_Value.Visible = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(1056, 28)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(78, 79)
        Me.PictureBox1.TabIndex = 32
        Me.PictureBox1.TabStop = False
        '
        'DailySales
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1134, 614)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Initial_Value)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Button_First)
        Me.Controls.Add(Me.Rows_No)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.From)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Search_Button1)
        Me.Controls.Add(Me.Convert_CSV)
        Me.Controls.Add(Me.Tran_Search_Text1)
        Me.Controls.Add(Me.PageNo1)
        Me.Controls.Add(Me.Combobox_Station)
        Me.Controls.Add(Me.Button_Next)
        Me.Controls.Add(Me.Button_Previous)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "DailySales"
        Me.Text = "EOD Sales Data Convertion  - Daily Sales"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Button_Previous As System.Windows.Forms.Button
    Friend WithEvents Button_Next As System.Windows.Forms.Button
    Friend WithEvents Combobox_Station As System.Windows.Forms.ComboBox
    Friend WithEvents PageNo1 As System.Windows.Forms.TextBox
    Friend WithEvents Tran_Search_Text1 As System.Windows.Forms.TextBox
    Friend WithEvents Convert_CSV As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents DataConvertionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TransactionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Search_Button1 As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents From As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Rows_No As System.Windows.Forms.ComboBox
    Friend WithEvents ExportedFilesTrackingToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Button_First As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Initial_Value As System.Windows.Forms.Label
    Friend WithEvents ExcelShopToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
