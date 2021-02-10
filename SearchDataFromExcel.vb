    Imports System.Data.OleDb

    Public Class frmSearchExcelDataTable

        Dim DT As New DataTable

        ' / --------------------------------------------------------------------
        ' / เลือกไฟล์ Excel
        Private Sub btnOpenExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenExcel.Click
            '/ ประกาศใช้งาน Open File Dialog ในแบบ Run Time
            Dim dlgOpenFile As OpenFileDialog = New OpenFileDialog()

            ' / ตั้งค่าการใช้งาน Open File Dialog
            With dlgOpenFile
                .InitialDirectory = MyPath(Application.StartupPath)
                .Title = "เลือกไฟล์ MS Excel"
                .Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|XLS Files (*.xls)|*xls"
                .FilterIndex = 0
                .RestoreDirectory = True
            End With
            '/ หากเลือกปุ่ม OK หลังจากการ Browse ...
            If dlgOpenFile.ShowDialog() = DialogResult.OK Then
                txtFileName.Text = dlgOpenFile.FileName
                Dim strConn As String = _
                    " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                    dlgOpenFile.FileName & ";" & _
                    " Extended Properties=""Excel 12.0 Xml; HDR=YES"";"
                Dim Conn As New OleDbConnection(strConn)
                Conn.Open()
                '/ มอง WorkSheet ให้เป็นตารางข้อมูล (Table)
                Dim dtSheets As DataTable = Conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                Dim drSheet As DataRow
                '//
                cmbSheetName.Items.Clear()
                '/ นำรายชื่อ WorkSheet ทั้งหมด มาเก็บไว้ที่ ComboBox เพื่อรอให้ User เลือกนำไปใช้งาน
                For Each drSheet In dtSheets.Rows
                    cmbSheetName.Items.Add(drSheet("TABLE_NAME").ToString)
                Next
                Conn.Close()
            End If
        End Sub

        ' / --------------------------------------------------------------------
        '// เลือก WorkSheet แล้วแสดงผลข้อมูลลงในตารางกริด
        Private Sub cmbSheetName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSheetName.SelectedIndexChanged
            Dim Conn As OleDbConnection
            Dim Cmd As OleDbCommand
            Dim DA As New OleDbDataAdapter
            Dim DS As DataSet
            Dim strConn As String = _
                " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                txtFileName.Text & ";" & _
                " Extended Properties=""Excel 12.0 Xml; HDR=YES"";"
            Try
                dgvData.DataSource = Nothing
                DT = New DataTable
                Conn = New OleDbConnection
                Conn.ConnectionString = strConn
                Cmd = New OleDbCommand
                '/ เสมือน WorkSheet เป็น Table ในฐานข้อมูล
                Cmd.CommandText = "Select * FROM [" & cmbSheetName.Text & "]"
                Cmd.Connection = Conn
                DA.SelectCommand = Cmd
                DA.Fill(DT)
                DT.TableName = cmbSheetName.Text.Replace("$", "")
                dgvData.DataSource = DT
                '//
                lblRecCount.Text = "[จำนวน : " & dgvData.Rows.Count & " รายการ]"
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                Conn = Nothing
                Cmd = Nothing
                DA = Nothing
                DS = Nothing
            End Try
        End Sub

        ' / --------------------------------------------------------------------
        '// การค้นหาข้อมูลจาก DataTable
        Private Sub txtSearch_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSearch.KeyPress
            '// ยังไม่ได้โหลดข้อมูลเข้ามา
            If DT.Rows.Count = 0 Then
                txtSearch.Clear()
                Return
            End If
            '// ตัดตัวอักขระที่ไม่พึงประสงค์ ที่มีผลต่อการค้นหาข้อมูล
            txtSearch.Text = txtSearch.Text.Trim.Replace("'", "").Replace("*", "").Replace("%", "")
            '// เช็คว่ากด Enter ในช่อง TextBox ที่ใช้ค้นหาหรือไม่
            If Asc(e.KeyChar) = 13 Then
                Dim DTSearch As New DataTable
                dgvData.DataSource = Nothing
                '// เพิ่มหลัก (Column) เข้ามาใหม่ให้กับ DataTable
                DTSearch.Columns.Add("ID", GetType(String))
                DTSearch.Columns.Add("Name", GetType(String))
                DTSearch.Columns.Add("Point", GetType(Integer))
                '// ค้นหาข้อมูลด้วย LIKE
                Dim result() As DataRow = DT.Select("NAME LIKE '*" & txtSearch.Text & "*'")
                '// หากพบข้อมูลก็ทำการลูปรายการแถวข้อมูล
                For Each row As DataRow In result
                    '// เพิ่มรายการแถวข้อมูลเข้าไปยัง DataTable
                    DTSearch.Rows.Add(row(0).ToString, row(1).ToString, row(2).ToString)
                Next
                dgvData.DataSource = DTSearch
                txtSearch.Clear()
                lblRecCount.Text = "[จำนวน : " & dgvData.Rows.Count & " รายการ]"
            End If
        End Sub

        Private Sub dgvData_DoubleClick(sender As Object, e As System.EventArgs) Handles dgvData.DoubleClick
            If dgvData.Rows.Count = 0 Then Return
            MessageBox.Show("คุณเลือก: " & vbCrLf & _
                            "ID = " & dgvData.SelectedRows(0).Cells(0).Value.ToString & vbCrLf & _
                            "NAME = " & dgvData.SelectedRows(0).Cells(1).Value.ToString & vbCrLf & _
                            "POINT = " & dgvData.SelectedRows(0).Cells(2).Value.ToString)
        End Sub

        Private Sub frmSearchExcelDataTable_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
            Call SetupDGVData()
            lblRecCount.Text = "[จำนวน : 0 รายการ]"
        End Sub

        ' / --------------------------------------------------------------------
        '// Initialize DataGridView @Run Time
        Private Sub SetupDGVData()
            With dgvData
                .RowHeadersVisible = False
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AllowUserToResizeRows = False
                .MultiSelect = False
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect
                .ReadOnly = True
                .Font = New Font("Tahoma", 9)
                ' Autosize Column
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                '// Even-Odd Color
                .AlternatingRowsDefaultCellStyle.BackColor = Color.AliceBlue
                ' Adjust Header Styles
                With .ColumnHeadersDefaultCellStyle
                    .BackColor = Color.Navy
                    .ForeColor = Color.Black ' Color.White
                    .Font = New Font("Tahoma", 9, FontStyle.Bold)
                End With
            End With
        End Sub

        ' / ------------------------------------------------------------------
        ' / ฟังค์ชั่นที่เราสามารถกำหนด Path ให้กับโปรแกรมของเราเอง
        ' / Ex.
        ' / AppPath = C:\My Project\bin\debug
        ' / Replace "\bin\debug" with ""
        ' / Return : C:\My Project\
        Function MyPath(AppPath As String) As String
            AppPath = AppPath.ToLower()
            MyPath = AppPath.Replace("\bin\debug", "").Replace("\bin\release", "")
            '// If not found folder then put the \ (BackSlash ASCII Code = 92) at the end.
            If Microsoft.VisualBasic.Right(MyPath, 1) <> Chr(92) Then MyPath = MyPath & Chr(92)
        End Function

        Private Sub btnExit_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
            Me.Close()
        End Sub

        Private Sub frmSearchExcelDataTable_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
            Me.Dispose()
            GC.SuppressFinalize(Me)
            Application.Exit()
        End Sub

    End Class
    
    
    ' https://www.g2gnet.com/webboard/forum.php?mod=viewthread&tid=368&extra=page%3D1
