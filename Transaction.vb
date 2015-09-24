Imports MySql.Data
Imports MySql.Data.MySqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System
Imports System.Text
Imports System.Data.OleDb
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Transaction


    Private Sub Transaction_Load(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load


        ' Transaction
        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales;  default command timeout=500;"
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
            Dim Command As New MySqlCommand
            Command.Connection = MysqlConn

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


            ' handle Search Value
            Dim Search_Value As String
            Search_Value = Tran_Search_Text1.Text

            ' handle Sation
            Dim Search_Station As String
            Search_Station = Trim(ComboBox_Station.SelectedItem)

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandText = "Retrieve_Transaction"
            Command.Parameters.AddWithValue("@offset_sp", Offset_sp)
            Command.Parameters.AddWithValue("@Rows_No", Row_No)
            Command.Parameters.AddWithValue("@search_Value", Search_Value)
            Command.Parameters.AddWithValue("@search_station", Search_Station)
            Command.Parameters.AddWithValue("@start_date1", Start_Date)
            Command.Parameters.AddWithValue("@End_date1", End_Date)
        
            MyAdapter.SelectCommand = Command

            Dim dt As New DataTable()
            MyAdapter.Fill(dt)
            DataGridView1.DataSource = dt


        Catch myerror As MySqlException

            MessageBox.Show("Cannot connect to database: " & myerror.Message)

        Finally

            MysqlConn.Close()
            MysqlConn.Dispose()

        End Try

    End Sub


    Private Sub Excel1_Click(sender As Object, e As EventArgs)

        ' Get present data for file name
        Dim File_Date As String = ""
        File_Date = DateTime.Now.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HHmmss")

        Dim xlApp As Excel.Application = New Excel.Application

        xlApp.SheetsInNewWorkbook = 1

        Dim xlWorkBook As Excel.Workbook = xlApp.Workbooks.Add
        Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Worksheets.Item(1)

        'Dim misValue As Object = System.Reflection.Missing.Value
        'xlWorkBook = xlApp.Workbooks.Add(misValue)

        xlWorkSheet.Name = "EODSales_Export"

        ' Write header to the Excel
        For Each col As DataGridViewColumn In Me.DataGridView1.Columns
            xlWorkSheet.Cells(1, col.Index + 1) = col.HeaderText.ToString
        Next

        'Write content
        For nRow = 1 To DataGridView1.Rows.Count - 1

            For nCol = 0 To DataGridView1.Columns.Count - 1
                xlWorkSheet.Cells(nRow + 1, nCol + 1) = DataGridView1.Rows(nRow).Cells(nCol).Value
            Next nCol

        Next nRow

        xlApp.DisplayAlerts = False

        xlWorkBook.SaveAs("C:\LOG\EODSales_" & File_Date & "_.xls", Excel.XlFileFormat.xlExcel5, Type.Missing, Type.Missing, Type.Missing, Type.Missing, _
                         Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges)

        MsgBox("Excel Exported.")

        xlWorkBook.Close()
        xlApp.Quit()
    End Sub

    Private Sub FileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.Click

        Main.Show()
        Me.Hide()

    End Sub

    Private Sub Search_Button1_Click(sender As Object, e As EventArgs) Handles Search_Button1.Click

        Initial_Value.Text = "1"
        Transaction_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Button_Next_Click(sender As Object, e As EventArgs) Handles Button_Next.Click

        Dim Page_No As Int16

        Page_No = PageNo1.Text.Trim

        Page_No = Page_No + 1

        PageNo1.Text = Page_No

        Transaction_Load(Me, New System.EventArgs)

    End Sub


    Private Sub Convert_CSV_Click(sender As Object, e As EventArgs) Handles Convert_CSV.Click

        Dim MysqlConn As MySqlConnection
        MysqlConn = New MySqlConnection()
        MysqlConn.ConnectionString = "server=127.0.0.1 ; user id=root; password=robin; database = EODSales;  default command timeout=500;"
        MysqlConn.Open()

        ' Set Initial Date Range
        If Initial_Value.Text = "0" Then
            DateTimePicker1.Value = DateAdd(DateInterval.Month, -1, Now)
            DateTimePicker2.Value = Now
        End If

        Dim Start_Date As String = DateTimePicker1.Value
        Dim End_Date As String = DateTimePicker2.Value

        ' Prepare CSV
        Dim File_Date As String = ""
        File_Date = DateTime.Now.ToString("yyyyMMdd") & "_" & DateTime.Now.ToString("HHmmss")

        Dim Export_Dir As String = Main.Label13.Text

        Dim FILE_NAME As String = Export_Dir & "Transaction_" & File_Date & ".csv"
        File.Create(FILE_NAME).Dispose()
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME, True, System.Text.Encoding.Unicode)


        Try
            Dim MyAdapter As New MySqlDataAdapter
            Dim Command As New MySqlCommand
            Command.Connection = MysqlConn

            Dim Search_Value As String
            Search_Value = Tran_Search_Text1.Text.ToString

            ' handle Sation
            Dim Search_Station As String
            Search_Station = Trim(ComboBox_Station.SelectedItem)

            Command.CommandType = CommandType.StoredProcedure
            Command.CommandText = "Retrieve_Transaction"
            Command.Parameters.AddWithValue("@offset_sp", 0)
            Command.Parameters.AddWithValue("@Rows_No", 999999)
            Command.Parameters.AddWithValue("@search_Value", Search_Value)
            Command.Parameters.AddWithValue("@search_station", Search_Station)
            Command.Parameters.AddWithValue("@start_date1", Start_Date)
            Command.Parameters.AddWithValue("@End_date1", End_Date)

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

                    sMsg &= Chr(9)

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

                            sMsg &= Chr(9)

                        End If

                    Next

                    sMsg &= Environment.NewLine
                    'write to file
                    objWriter.Write(sMsg)

                Loop

            End If


        Catch myerror As MySqlException

            MessageBox.Show("Cannot connect to database: " & myerror.Message)

        Finally

            MsgBox("CSV file converted")
            objWriter.Close()
            MysqlConn.Close()
            MysqlConn.Dispose()

        End Try


    End Sub

    Private Sub Button_Previous_Click(sender As Object, e As EventArgs) Handles Button_Previous.Click


        Dim Page_No As Int16

        Page_No = PageNo1.Text.Trim

        Dim butt As Button = DirectCast(sender, Button)

        Page_No = Page_No - 1

        PageNo1.Text = Page_No

        Transaction_Load(Me, New System.EventArgs)

    End Sub

    Private Sub Rows_No_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Rows_No.SelectedIndexChanged

        Transaction_Load(Me, New System.EventArgs)

    End Sub

  
    Private Sub SalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesToolStripMenuItem.Click

        Sales.Show()
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

   
    Private Sub Button_First_Click(sender As Object, e As EventArgs) Handles Button_First.Click


        Dim Page_No As Int16

        Page_No = 1

        PageNo1.Text = Page_No

        Transaction_Load(Me, New System.EventArgs)

    End Sub

    Private Sub ComboBox_Station_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Station.SelectedIndexChanged

        Transaction_Load(Me, New System.EventArgs)

    End Sub

    Private Sub ExcelShopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExcelShopToolStripMenuItem.Click
        ExcelShop.Show()
        Me.Hide()

    End Sub
End Class
