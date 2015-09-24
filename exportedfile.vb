Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System
Imports System.Text
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExportedFile


    Private Sub ExportedFile_Load(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load


        ' ExportedFile
        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales; default command timeout= 500;"
        MysqlConn.Open()

        ' Handle the previous button
        Dim Page_No As Int16
        Page_No = PageNo1.Text.Trim
        If Page_No = 1 Then
            Button_Previous.Hide()
        Else
            Button_Previous.Show()
        End If

        Try
            Dim MyAdapter As New MySqlDataAdapter
            Dim Command As New MySqlCommand("SET ARITHABORT ON", MysqlConn)
            'Command.Connection = MysqlConn

            ' Set Initial Date Range
            If Initial_Value.Text = "0" Then
                DateTimePicker1.Value = DateAdd(DateInterval.Month, -1, Now)
                DateTimePicker2.Value = Now
            End If

            ' Date Range
            Dim Start_Date As String = DateTimePicker1.Value
            Dim End_Date As String = DateTimePicker2.Value

            If DateTimePicker1.Value > DateTimePicker2.Value Then
                MsgBox("The end date must be great than the start date.")
            End If

           
            ' Set no. of rows per page
            Dim Row_No As Integer
            Row_No = CInt(Rows_No.SelectedItem)

            ' Set No. of page to show
            Dim Offset_sp As Integer
            Offset_sp = PageNo1.Text
            Offset_sp = (Offset_sp - 1) * Row_No

            ' Handle Check Box
            Dim Check_file_value As Integer
            If Check_Missing_File.Checked Then
                Check_file_value = 1
            Else
                Check_file_value = 0
            End If

          

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandText = "Retrieve_Exported_Files"
            Command.Parameters.AddWithValue("@offset_sp", Offset_sp)
            Command.Parameters.AddWithValue("@Rows_No", Row_No)
            Command.Parameters.AddWithValue("@start_date1", Start_Date)
            Command.Parameters.AddWithValue("@End_date1", End_Date)
            Command.Parameters.AddWithValue("@Check_File", Check_file_value)

            MyAdapter.SelectCommand = Command

            Dim dt As New DataTable()
            MyAdapter.Fill(dt)
            DataGridView1.DataSource = dt

            ' Add Check Box Column in DataGridView

            If DataGridView1.Columns.Count = 6 And DataGridView1.ColumnCount > 0 Then
                Dim Chckbx_Column As New DataGridViewCheckBoxColumn
                Chckbx_Column.HeaderText = "Delete"
                Chckbx_Column.Name = "Delete_File"
                Me.DataGridView1.Columns.Add(Chckbx_Column)
            End If



        Catch myerror As MySqlException

            MessageBox.Show("Cannot connect to database: " & myerror.Message)

        Finally

            MysqlConn.Close()
            MysqlConn.Dispose()

        End Try

    End Sub


   

    Private Sub FileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.Click

        Main.Show()
        Me.Hide()

    End Sub

    Private Sub Search_Button1_Click(sender As Object, e As EventArgs)

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Button_Next_Click(sender As Object, e As EventArgs) Handles Button_Next.Click

        Dim Page_No As Int16

        Page_No = PageNo1.Text.Trim

        Page_No = Page_No + 1

        PageNo1.Text = Page_No

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub


    Private Sub Convert_CSV_Click(sender As Object, e As EventArgs)

        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales; default command timeout= 500;"
        MysqlConn.Open()

        Dim Start_Date As Date = DateTimePicker1.Value
        Dim End_Date As Date = DateTimePicker2.Value

        ' Prepare CSV
        Dim File_Date As String = ""
        File_Date = DateTime.Now.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HHmmss")

        Dim Export_Dir As String = Main.Label13.Text

        Dim FILE_NAME As String = Export_Dir & File_Date & ".csv"
        File.Create(FILE_NAME).Dispose()
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)


        Try
            Dim MyAdapter As New MySqlDataAdapter
            Dim Command As New MySqlCommand("SET ARITHABORT ON", MysqlConn)
            'Command.Connection = MysqlConn

          

            Command.CommandText = "Retrieve_ExportedFile"

            Command.CommandType = CommandType.StoredProcedure
            Command.Parameters.AddWithValue("@offset_sp", 0)
            Command.Parameters.AddWithValue("@Rows_No", 999999)
            Command.Parameters.AddWithValue("@start_date", Start_Date)
            Command.Parameters.AddWithValue("@End_date", End_Date)

            MyAdapter.SelectCommand = Command

            Command.ExecuteNonQuery()


            Dim dr As MySqlDataReader
            Dim sMsg As String

            MyAdapter.SelectCommand = Command

            Dim dt As New DataTable()
            MyAdapter.Fill(dt)


            dr = Command.ExecuteReader


            Dim j As Integer = 0

            sMsg = ""

            For Each column As DataColumn In dt.Columns

                sMsg &= column.ColumnName


                If j < dr.FieldCount - 1 Then

                    sMsg &= ","

                End If

                j = j + 1
            Next



            sMsg &= Environment.NewLine
            'write to file
            objWriter.Write(sMsg)


            If dr.HasRows Then


                Do While dr.Read

                    sMsg = ""

                    For index As Integer = 0 To dr.FieldCount - 1

                        sMsg &= dr.GetValue(index).ToString

                        If index < dr.FieldCount - 1 Then

                            sMsg &= ","

                        End If

                    Next

                    sMsg &= Environment.NewLine
                    'write to file
                    objWriter.Write(sMsg)

                Loop

            End If


            objWriter.Close()
            MsgBox("CSV file converted")
            MysqlConn.Close()

        Catch myerror As MySqlException

            MessageBox.Show("Cannot connect to database: " & myerror.Message)

        Finally

            MysqlConn.Dispose()

        End Try


    End Sub

    Private Sub Button_Previous_Click(sender As Object, e As EventArgs) Handles Button_Previous.Click


        Dim Page_No As Int16

        Page_No = PageNo1.Text.Trim

        Dim butt As Button = DirectCast(sender, Button)

        Page_No = Page_No - 1

        PageNo1.Text = Page_No

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Rows_No_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Rows_No.SelectedIndexChanged

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub

  
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs)

        Close()

    End Sub

    Private Sub ExitToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click

        Close()

    End Sub

    Private Sub TransactionDetailToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransactionDetailToolStripMenuItem.Click

        Transaction.Show()
        Me.Hide()

    End Sub

    Private Sub SalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesToolStripMenuItem.Click

        Sales.Show()
        Me.Hide()

    End Sub

    Private Sub DailySalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DailySalesToolStripMenuItem.Click

        DailySales.Show()
        Me.Hide()

    End Sub

    Private Sub Check_Missing_File_CheckedChanged(sender As Object, e As EventArgs) Handles Check_Missing_File.CheckedChanged

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Button_First_Click(sender As Object, e As EventArgs) Handles Button_First.Click

        Dim Page_No As Int16

        Page_No = 1

        PageNo1.Text = Page_No

        ExportedFile_Load(Me, New System.EventArgs)

    End Sub


    Private Sub ButtonSearch_Click(sender As Object, e As EventArgs) Handles ButtonSearch.Click

        Initial_Value.Text = "1"
        ExportedFile_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Button_Delete_Click(sender As Object, e As EventArgs) Handles Button_Delete.Click

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


                MessageBox.Show("Deleting.")

                Dim sb As New System.Text.StringBuilder
                For Each row As DataGridViewRow In CheckedRows
                    sb.AppendLine(row.Cells("File Name").Value.ToString)
                Next

                Command.CommandType = CommandType.StoredProcedure
                Command.CommandText = "Delete_Exported_Record"
                Command.Parameters.AddWithValue("@File_Name", sb)
                MyAdapter.SelectCommand = Command
                Command.ExecuteNonQuery()


                ExportedFile_Load(Me, New System.EventArgs)

                MessageBox.Show(sb.ToString & " was deleted")

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub ExcelShopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelShopToolStripMenuItem.Click
        ExcelShop.Show()
        Me.Hide()
    End Sub
End Class
