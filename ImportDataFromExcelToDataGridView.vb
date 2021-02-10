' [VB.NET] Import ข้อมูลใน Excel เข้ามาแสดงผลในตารางกริด

    Imports System.Data.OleDb

    Public Class frmExcel2DataGrid

        ' / --------------------------------------------------------------------
        ' / เลือกไฟล์ Excel
        Private Sub btnOpenExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenExcel.Click
            '/ ประกาศใช้งาน Open File Dialog ในแบบ Run Time
            Dim dlgOpenFile As OpenFileDialog = New OpenFileDialog()

            ' / ตั้งค่าการใช้งาน Open File Dialog
            With dlgOpenFile
                .InitialDirectory = MyPath(Application.StartupPath)
                .Title = "เลือกไฟล์ MS Excel"
                .Filter = "All Files (*.*)|*.*|Excel files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|XLS Files (*.xls)|*xls"
                .FilterIndex = 1
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
            Dim Comm As OleDbCommand
            Dim DAP As OleDbDataAdapter
            Dim DS As DataSet

            Dim strConn As String = _
                " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                txtFileName.Text & ";" & _
                " Extended Properties=""Excel 12.0 Xml; HDR=YES"";"
            Try
                Conn = New OleDbConnection
                Conn.ConnectionString = strConn
                Comm = New OleDbCommand
                '/ เสมือน WorkSheet เป็น Table ในฐานข้อมูล
                Comm.CommandText = "Select * FROM [" & cmbSheetName.Text & "]"
                Comm.Connection = Conn
                DAP = New OleDbDataAdapter(Comm)
                DS = New DataSet
                Conn.Open()
                DAP.Fill(DS, "Sheet1")
                '/ ผูกข้อมูล (Bound Data) เข้ากับ Sheet1
                dgvData.DataSource = DS.Tables("Sheet1")
                '//
                Call SetupDGVData()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                Conn = Nothing
                Comm = Nothing
                DAP = Nothing
                DS = Nothing
            End Try
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

        Private Sub frmExcel2DataGrid_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
            Me.Dispose()
            Application.Exit()
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

    End Class
