Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System
Imports System.Text
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Main


    Private Sub Form1_Load(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        ListBox1.Items.Clear()

        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = eodsales"
        MysqlConn.Open()
        Try
            Dim MyAdapter As New MySqlDataAdapter
            Dim SqlQuery = "SELECT * From SystemSetting"
            Dim Command As New MySqlCommand
            Command.Connection = MysqlConn

            Command.CommandText = SqlQuery
            MyAdapter.SelectCommand = Command
            Dim Mydata As MySqlDataReader
            Mydata = Command.ExecuteReader
            Mydata.Read()


            If Mydata.HasRows Then

                Label8.Text = Mydata.GetValue(0)

                Label9.Text = Mydata.GetString(1)

                Label10.Text = Mydata.GetString(2)

                Label11.Text = Mydata.GetString(3)

                Label13.Text = Mydata.GetString(4)

                Label15.Text = Mydata.GetString(5)

                Dim Files() As System.IO.FileInfo
                Dim DirInfo As New System.IO.DirectoryInfo(Mydata.GetValue(0))
                Dim counter = My.Computer.FileSystem.GetFiles(Mydata.GetValue(0))
                Files = DirInfo.GetFiles("*." & Mydata.GetString(1))
                Dim Filelist As String

                Mydata.Close()

                ' Display raw files on screen
                ListBox1.Items.Add("File Name                                                                                   Creation Date")
                For Each File In Files

                    File.Attributes = FileAttributes.Archive


                    ' Check depulicate files
                    Dim MyAdapter1 As New MySqlDataAdapter
                    Dim Command1 As New MySqlCommand
                    Command1.Connection = MysqlConn
                    Dim SqlQuery1 = "SELECT Count(1) From file_Header where File_Name ='" & File.Name & "'"
                    Command1.CommandText = SqlQuery1
                    MyAdapter1.SelectCommand = Command1
                    Dim Mydata1 As MySqlDataReader
                    Mydata1 = Command1.ExecuteReader
                    Mydata1.Read()
                    If Mydata1.GetValue(0) = 1 Then
                        File.Attributes = FileAttributes.Normal
                        Filelist = File.ToString() & " - Converted       " & File.LastAccessTime.ToString()
                    Else
                        Filelist = File.ToString() & "                             " & File.LastAccessTime.ToString()
                    End If
                    Mydata1.Close()

                    ' Show the files
                    ListBox1.Items.Add(Filelist)
                Next

                Label1.Text = "Number of files:  "
                Label7.Text = CStr(counter.Count)



            Else
                MsgBox("System Setting Not Set.")
            End If




            MysqlConn.Close()

        Catch myerror As MySqlException

            MessageBox.Show("Cannot connect to database: " & myerror.Message)

        Finally

            MysqlConn.Dispose()

        End Try


    End Sub

    Private Sub Convert_Click(sender As Object, e As EventArgs) Handles Convert.Click

        ' minmize the form to background running
        Me.WindowState = FormWindowState.Minimized
        Visible = False

        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1; user id=root; password=robin; database = EODSales;"
        MysqlConn.Open()
        Dim MyAdapter As New MySqlDataAdapter
        Dim Command As New MySqlCommand
        Command.Connection = MysqlConn

        ' Get System Setting
        Dim sysSqlQuery = "select * From SystemSetting"
        Command.CommandText = sysSqlQuery
        MyAdapter.SelectCommand = Command
        Dim Mysysdata As MySqlDataReader
        Mysysdata = Command.ExecuteReader
        Mysysdata.Read()
        Dim Dir_Info As String = Mysysdata.GetString(0)
        Dim File_Ext As String = Mysysdata.GetString(1)
        Dim separator As String = Mysysdata.GetString(2)
        Dim Log_Dir As String = Mysysdata.GetString(3)
        ' number of raw files
        Dim counter = My.Computer.FileSystem.GetFiles(Mysysdata.GetValue(0))
        Mysysdata.Close()

        Dim Directory = Dir_Info
        Dim File_Date As String = ""
        File_Date = DateTime.Now.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HHmmss")
        Dim Files() As System.IO.FileInfo
        Dim DirInfo As New System.IO.DirectoryInfo(Directory)
        Files = DirInfo.GetFiles("*." & File_Ext)
        Dim FILE_NAME As String = Log_Dir & "\EODSales_" & File_Date & ".log"
        File.Create(FILE_NAME).Dispose()
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        Dim Field_Name As String = ""
        Dim RecordType As String = ""
        Dim seperator() As String = {"|"}
        Dim Filename As String
        Dim i As Integer = 1
        Dim Total_files As String = Label7.Text



        For Each File In Files



            If File.Attributes = FileAttributes.Archive Then

                ' Setting notify icon in system tray
                NotifyIcon1 = New System.Windows.Forms.NotifyIcon(components)
                NotifyIcon1.Icon = New Icon("favicon.ico")

                ' text on notifyIcon 
                NotifyIcon1.Text = ("Converting " & i & " of total " & Total_files & " Files ")
                NotifyIcon1.Visible = True

                ' Define reader to read line
                Dim textReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(File.FullName, System.Text.Encoding.UTF8)
                Dim aLine As String = textReader.ReadLine()
                Dim Content As String = ""
                Dim TransactionDate As String = ""
                Dim Transaction_ID As String = ""

                Filename = File.Name

                Dim Station_id As String = Mid(Filename, 14, 3)

                ' objWriter.Write(Filename & Chr(13) & Chr(10))

                While Not (aLine Is Nothing)

                    ' Content of each line

                    Content = aLine & "|" & Filename
                    Content = Content.Replace("""", "")
                    Content = Replace(Content, """", "").ToString
                    Content = Replace(Content, "'", "").ToString
                    Content = Replace(Content, "|", "','").ToString

                    aLine = textReader.ReadLine()

                    RecordType = Content.Substring(0, 3)

                    Try

                        Select Case RecordType
                                  Case "000"

                                Dim SqlQuery0 = "insert into file_header (Report_Type, Country_Code, Site_ID, Incremental_Counter, File_Type, Business_Day, Start_Date, Start_Time, End_Date, End_Time, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0


                            Case "500"

                                TransactionDate = Mid(Content, 7, 8)

                                Dim SqlQuery0 = "insert into ticket_header (Report_Type, Transaction_Date, Transaction_Start_Time, Transaction_id, Sales_Operator, POS_ID, Ticket_Total, File_Name) Values ('" & Content & "') ; Insert into ExcelShop (Station_ID, Transaction_Date, Transaction_id, File_Name) Values ('" & Station_id & "', '" & TransactionDate & "' ,  (Select Transaction_ID from ticket_header order by idticket_header desc limit 1), '" & Filename & "') "
                                'Dim SqlQuery0 = "insert into ticket_header (Report_Type, Transaction_Date, Transaction_Start_Time, Transaction_id, Sales_Operator, POS_ID, Ticket_Total, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0


                            Case "501"

                                Dim SqlQuery0 = "insert into sales (Report_Type,  Item_No, Transaction_id, Category_id, Extended_Sales_Price, Original_Amount, Extended_VAT, Mark_Down_Indicator, Sales_Mode, Product_EAN_Code, Product_ID, Product_Description, Quantity, Unit_Sales_Price, Unit_VAT, File_Name) Values ('" & Content & "'); Update ExcelShop Set Product_id = (Select Product_ID from sales order by idSales_n_Returns desc limit 1) , unit_Sales_Price = (Select unit_Sales_Price from sales order by idSales_n_Returns desc limit 1) where transaction_id = (Select Transaction_ID from ticket_header order by idticket_header desc limit 1) and file_name = '" & Filename & "' "
                                'Dim SqlQuery0 = "insert into sales (Report_Type,  Item_No, Transaction_id, Category_id, Extended_Sales_Price, Original_Amount, Extended_VAT, Mark_Down_Indicator, Sales_Mode, Product_EAN_Code, Product_ID, Product_Description, Quantity, Unit_Sales_Price, Unit_VAT, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0

                            Case "502"

                                Dim SqlQuery0 = "insert into sales (Report_Type,  Item_No, Transaction_id, Category_id, Extended_Sales_Price, Original_Amount, Extended_VAT, Mark_Down_Indicator, Sales_Mode, Product_EAN_Code, Product_ID, Product_Description, Quantity, Unit_Sales_Price, Unit_VAT, File_Name) Values ('" & Content & "'); Update ExcelShop Set Product_id = (Select Product_ID from sales order by idSales_n_Returns desc limit 1) , unit_Sales_Price = (Select unit_Sales_Price from sales order by idSales_n_Returns desc limit 1) where transaction_id = (Select Transaction_ID from ticket_header order by idticket_header desc limit 1) and file_name = '" & Filename & "' "
                                'Dim SqlQuery0 = "insert into sales (Report_Type,  Item_No, Transaction_id, Category_id, Extended_Sales_Price, Original_Amount, Extended_VAT, Mark_Down_Indicator, Sales_Mode, Product_EAN_Code, Product_ID, Product_Description, Quantity, Unit_Sales_Price, Unit_VAT, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0

                            Case "550"

                                Dim SqlQuery0 = "insert into price_correction (Report_Type,  Item_No, Temp_1, Transaction_ID, Price_Correction_Type, Price_Event_ID, Discount_PAN, Quantity, Amount, Value, VAT_Amount, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0


                            Case "600"

                                Dim SqlQuery0 = "insert into payments (Report_Type,  Item_No, Transaction_ID, Method_Of_Payment, MOP_Description, PAN, Transaction_Currency, Transaction_Amount, Transaction_Amount_Home, File_Name) Values ('" & Content & "'); Update ExcelShop Set mop = (Select Method_Of_Payment from Payments order by idPayments desc limit 1) ,  Transaction_Amount_Home = (Select Transaction_Amount_Home from Payments order by idPayments desc limit 1) where transaction_id = (Select Transaction_ID from ticket_header order by idticket_header desc limit 1) and file_name = '" & Filename & "' "
                                ' Dim SqlQuery0 = "insert into payments (Report_Type,  Item_No, Transaction_ID, Method_Of_Payment, MOP_Description, PAN, Transaction_Currency, Transaction_Amount, Transaction_Amount_Home, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0


                            Case "999"

                                Dim SqlQuery0 = "insert into trailer (Report_Type,  Start_Time, End_Time, Record_Count, Hash_Total_1, Hash_Total_2, File_Name) Values ('" & Content & "')"
                                Command.CommandText = SqlQuery0


                        End Select


                        MyAdapter.SelectCommand = Command
                        Command.ExecuteNonQuery()

                        objWriter.Write(Content & Chr(13) & Chr(10))

                    Catch ex As Exception

                        MsgBox(ex.Message)

                    End Try



                End While

                i = i + 1

                NotifyIcon1.Dispose()

                ' File.Refresh()
                ' File.MoveTo(Label15.Text & File.Name)

            End If



        Next
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Visible = True

        objWriter.Flush()
        objWriter.Close()
        MsgBox("Import Done Succeefully.")

        MysqlConn.Close()


        ListBox1.Items.Clear()

        Form1_Load(Me, New System.EventArgs)

    End Sub


    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        Select Case TabControl1.SelectedTab.Name
            Case "TabPage1"

                'Refresh listbox1
                ListBox1.Items.Clear()
                Form1_Load(Me, New System.EventArgs)




            Case "TabPage2"

                Dim MysqlConn As MySqlConnection
                MysqlConn = New MySqlConnection()
                MysqlConn.ConnectionString = "server=localhost ; user id=root; password=robin; database = EODSales"
                MysqlConn.Open()

                Dim MyAdapter As New MySqlDataAdapter
                Dim SqlQuery = "SELECT idMop as 'ID' , Payment_EOD as 'Payment EOD' ,  Payment, Group_Payment_Type as 'Payment Type', Group_Name as 'Group Name' From MOP"
                Dim Command As New MySqlCommand
                Command.Connection = MysqlConn
                Command.CommandText = SqlQuery
                MyAdapter.SelectCommand = Command

               
                Try

                    Dim dt As New DataTable()
                    MyAdapter.Fill(dt)
                    DataGridView1.DataSource = dt
                    DataGridView1.ReadOnly = False
                    MysqlConn.Close()

                    DataGridView1.Columns("ID").ReadOnly = True

                    ' Add Check Box Column in DataGridView
                    If DataGridView1.Columns.Count = 5 And DataGridView1.ColumnCount > 0 Then
                        Dim Chckbx_Column1 As New DataGridViewCheckBoxColumn
                        Chckbx_Column1.HeaderText = "Delete"
                        Chckbx_Column1.Name = "Delete_File"
                        Me.DataGridView1.Columns.Add(Chckbx_Column1)
                    End If


                Catch myerror As MySqlException

                    MessageBox.Show("Cannot connect to database: " & myerror.Message)

                Finally

                    MysqlConn.Dispose()

                End Try

            Case "TabPage5"

                Dim MysqlConn As MySqlConnection
                MysqlConn = New MySqlConnection()
                MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales"
                MysqlConn.Open()

                Dim MyAdapter As New MySqlDataAdapter
                Dim SqlQuery = "SELECT * From SystemSetting"
                Dim Command As New MySqlCommand
                Command.Connection = MysqlConn

                Command.CommandText = SqlQuery
                MyAdapter.SelectCommand = Command
                Dim Mydata As MySqlDataReader
                Mydata = Command.ExecuteReader
                Mydata.Read()


                If Mydata.HasRows Then


                    Dim Files() As System.IO.FileInfo
                    Dim DirInfo As New System.IO.DirectoryInfo(Mydata.GetValue(3))
                    Dim counter = My.Computer.FileSystem.GetFiles(Mydata.GetValue(3))
                    Files = DirInfo.GetFiles("*.log")
                    Dim Filelist As String

                    ListBox2.Items.Clear()

                    ListBox2.Items.Add("File Name")

                    For Each File In Files

                        Filelist = File.ToString()

                        ListBox2.Items.Add(Filelist)

                    Next

                    Label6.Text = "Log Directory: " & Mydata.GetString(3) & "        Number of files: " & CStr(counter.Count)

                End If

          

            Case "TabPage6"

                Dim MysqlConn As MySqlConnection
                MysqlConn = New MySqlConnection()
                MysqlConn.ConnectionString = "server=localhost ; user id=root; password=robin; database = EODSales"
                MysqlConn.Open()

                Dim MyAdapter As New MySqlDataAdapter
                Dim SqlQuery = "SELECT Dir_Name as 'Import Directory' , File_Extension as 'Extension',File_Separator as 'Separator', Log_Dir as 'Log Directory' , Export_Dir as 'Export Directory', Archive_Dir as 'Archive Directory' From SystemSetting"
                Dim Command As New MySqlCommand
                Command.Connection = MysqlConn
                Command.CommandText = SqlQuery
                MyAdapter.SelectCommand = Command

                Try

                    Dim dt As New DataTable()
                    MyAdapter.Fill(dt)
                    DataGridView4.DataSource = dt
                    DataGridView4.ReadOnly = False
                    MysqlConn.Close()

                Catch myerror As MySqlException

                    MessageBox.Show("Cannot connect to database: " & myerror.Message)

                Finally

                    MysqlConn.Dispose()

                End Try



        End Select

    End Sub
 
   


    

    Private Sub Setting_Save_Click(sender As Object, e As EventArgs) Handles Setting_Save.Click

        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=localhost ; user id=root; password=robin; database = EODSales"
        MysqlConn.Open()

        Dim Command As New MySqlCommand
        Command.Connection = MysqlConn

        Dim MyAdapter As New MySqlDataAdapter

        Dim SqlQuery0 = "Update systemsetting set dir_name=@dir_name, file_extension=@file_ext, File_Separator=@file_sep, Log_Dir=@Log_dir, Export_Dir=@export_dir , Archive_Dir=@archive_dir"

        Command.Parameters.AddWithValue("@dir_name", DataGridView4.Rows(0).Cells(0).Value.ToString())
        Command.Parameters.AddWithValue("@file_ext", DataGridView4.Rows(0).Cells(1).Value.ToString())
        Command.Parameters.AddWithValue("@file_sep", DataGridView4.Rows(0).Cells(2).Value.ToString())
        Command.Parameters.AddWithValue("@Log_dir", DataGridView4.Rows(0).Cells(3).Value.ToString())
        Command.Parameters.AddWithValue("@export_dir", DataGridView4.Rows(0).Cells(4).Value.ToString())
        Command.Parameters.AddWithValue("@archive_dir", DataGridView4.Rows(0).Cells(5).Value.ToString())


        Command.CommandText = SqlQuery0

        MyAdapter.SelectCommand = Command
        Command.ExecuteNonQuery()

        MsgBox("Record Successfully Updated")

        MysqlConn.Close()

        MysqlConn.Dispose()

    End Sub

    Private Sub OnRowValidating(ByVal sender As Object, ByVal args As DataGridViewCellCancelEventArgs) Handles DataGridView4.RowValidating

        Dim row As DataGridViewRow = DataGridView4.Rows(args.RowIndex)

        For Each cell As DataGridViewCell In row.Cells
            If String.IsNullOrEmpty(cell.Value.ToString()) Then
                MsgBox("Please enter a value in the field.")
            End If
        Next

        Dim cellVal0 As String = DataGridView4.Rows(0).Cells(0).Value.ToString()
        Dim cellVal1 As String = DataGridView4.Rows(0).Cells(1).Value.ToString()
        Dim cellVal2 As String = DataGridView4.Rows(0).Cells(2).Value.ToString()
        Dim cellVal3 As String = DataGridView4.Rows(0).Cells(3).Value.ToString()


        If cellVal0.Length > 45 Then

            MsgBox("The maximum size of the directory are 45 charaters only.")

        End If

        If cellVal1.Length > 3 Then

            MsgBox("The size of the file extension three charaters only.")

        End If

        If cellVal2.Length > 1 Then

            MsgBox("The size of the file separator is one charater only.")

        End If


        If cellVal3.Length > 45 Then

            MsgBox("The maximum size of the log directory are 45 charaters only.")

        End If

    End Sub

    Private Sub SalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesToolStripMenuItem.Click

        Sales.Show()
        Me.Hide()

    End Sub

    Private Sub Transaction_Search_Button_Click(sender As Object, e As EventArgs)

        Form1_Load(Me, New System.EventArgs)

    End Sub

    Private Sub TransactionDetailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransactionDetailToolStripMenuItem.Click

        Transaction.Show()
        Me.Hide()

    End Sub

    Private Sub DailySalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DailySalesToolStripMenuItem.Click

        DailySales.Show()
        Me.Hide()

    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click

        Close()

    End Sub

   
    Private Sub ExportedFilesTrackingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportedFilesTrackingToolStripMenuItem.Click

        ExportedFile.Show()
        Me.Hide()

    End Sub

    

    Private Sub Button_Open_File_Click(sender As Object, e As EventArgs) Handles Button_Open_File.Click

        Dim fullPath As String = Label11.Text & ListBox2.SelectedItem.ToString
        System.Diagnostics.Process.Start("notepad.exe", fullPath)

    End Sub


    Private Sub ButtonMove_Files_Click(sender As Object, e As EventArgs) Handles ButtonMove_Files.Click


        Dim destination_Dir As String = Label15.Text.ToString & DateTime.Now.ToString("yyyyMMdd")
        Dim Source_Dir As String = Label8.Text.ToString
        MsgBox(destination_Dir)
        If (Not System.IO.Directory.Exists(destination_Dir)) Then
            System.IO.Directory.CreateDirectory(destination_Dir)
        End If

        For Each foundFile As String In My.Computer.FileSystem.GetFiles(Source_Dir, Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*")

            Dim foundFileInfo As New System.IO.FileInfo(foundFile)
            My.Computer.FileSystem.MoveFile(foundFile, destination_Dir & "\" & foundFileInfo.Name)
        Next

        ListBox1.Items.Clear()

        Form1_Load(Me, New System.EventArgs)

    End Sub

    Private Sub ExcelShopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelShopToolStripMenuItem.Click

        ExcelShop.Show()
        Me.Hide()

    End Sub


    Private Sub Btn_Update_Click(sender As Object, e As EventArgs) Handles Btn_Update.Click

        Try
            Dim MysqlConn As MySqlConnection
            MysqlConn = New MySqlConnection()
            MysqlConn.ConnectionString = "server=localhost ; user id=root; password=robin; database = EODSales"
            MysqlConn.Open()

            Dim k As Integer

            Do Until k = DataGridView1.RowCount - 1

                Dim Command As New MySqlCommand
                Command.Connection = MysqlConn
                Dim MyAdapter As New MySqlDataAdapter

                Dim SqlQuery1 = "Update mop set Payment_EOD=@Payment_EOD, Payment=@Payment, Group_Payment_Type=@Group_Payment_Type, Group_Name=@Group_Name Where idMop= " & DataGridView1.Rows(k).Cells(0).Value & ""

                Command.Parameters.AddWithValue("@Payment_EOD", DataGridView1.Rows(k).Cells(1).Value)
                Command.Parameters.AddWithValue("@Payment", DataGridView1.Rows(k).Cells(2).Value.ToString())
                Command.Parameters.AddWithValue("@Group_Payment_Type", DataGridView1.Rows(k).Cells(3).Value)
                Command.Parameters.AddWithValue("@Group_Name", DataGridView1.Rows(k).Cells(4).Value.ToString())

                Command.CommandText = SqlQuery1

                MyAdapter.SelectCommand = Command
                Command.ExecuteNonQuery()

                k = k + 1

            Loop

            MsgBox("Record Successfully Updated")

            TabControl1.SelectedIndex = 1

            TabControl1.SelectedIndex = 3

            MysqlConn.Close()

            MysqlConn.Dispose()



        Catch exc As Exception
            MsgBox(exc.Message, MsgBoxStyle.OkOnly, "Connection Failed")
        End Try

    End Sub

    Private Sub Btn_MOP_Insert_Click(sender As Object, e As EventArgs) Handles Btn_MOP_Insert.Click

        Try
            Dim MysqlConn As MySqlConnection
            MysqlConn = New MySqlConnection()
            MysqlConn.ConnectionString = "server=localhost ; user id=root; password=robin; database = EODSales"
            MysqlConn.Open()

            Dim l As Integer = DataGridView1.RowCount - 2
        
            Dim Command As New MySqlCommand
            Command.Connection = MysqlConn
            Dim MyAdapter As New MySqlDataAdapter

            Dim SqlQuery2 = "insert into mop (Payment_EOD, Payment, Group_Payment_Type, Group_Name) Values (@Payment_EOD, @Payment, @Group_Payment_Type, @Group_Name)"

            Command.Parameters.AddWithValue("@Payment_EOD", DataGridView1.Rows(l).Cells(1).Value)
            Command.Parameters.AddWithValue("@Payment", DataGridView1.Rows(l).Cells(2).Value.ToString())
            Command.Parameters.AddWithValue("@Group_Payment_Type", DataGridView1.Rows(l).Cells(3).Value)
            Command.Parameters.AddWithValue("@Group_Name", DataGridView1.Rows(l).Cells(4).Value.ToString())

            Command.CommandText = SqlQuery2

            MyAdapter.SelectCommand = Command
            Command.ExecuteNonQuery()

            MsgBox("Record Successfully Updated")

            TabControl1.SelectedIndex = 1

            TabControl1.SelectedIndex = 3

            MysqlConn.Close()

            MysqlConn.Dispose()



        Catch exc As Exception
            MsgBox(exc.Message, MsgBoxStyle.OkOnly, "Connection Failed")
        End Try

    End Sub

    Private Sub Del_MOP_Btn_Click(sender As Object, e As EventArgs) Handles Del_MOP_Btn.Click


        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales"
        MysqlConn.Open()

        Try

            Dim MyAdapter As New MySqlDataAdapter
            Dim Command As New MySqlCommand
            Command.Connection = MysqlConn

            Dim CheckedRows =
                (
                    From Rows In DataGridView1.Rows.Cast(Of DataGridViewRow)()
                    Where CBool(Rows.Cells("Delete_File").Value) = True
                ).ToList
            If CheckedRows.Count = 0 Then

                MessageBox.Show("Nothing selected")

            ElseIf CheckedRows.Count > 1 Then

                MessageBox.Show("Only one file is allowed.")

            Else

                Dim sb As New System.Text.StringBuilder
                For Each row As DataGridViewRow In CheckedRows
                    sb.AppendLine(row.Cells("ID").Value.ToString)
                Next

                Command.CommandType = CommandType.StoredProcedure
                Command.CommandText = "Delete_MOP"
                Command.Parameters.AddWithValue("@ID", sb)
                MyAdapter.SelectCommand = Command
                Command.ExecuteNonQuery()

                MessageBox.Show("Row number " & sb.ToString & " was deleted")

                TabControl1.SelectedIndex = 1

                TabControl1.SelectedIndex = 3



            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try



    End Sub
End Class
