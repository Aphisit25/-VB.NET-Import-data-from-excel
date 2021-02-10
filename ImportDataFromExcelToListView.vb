        ' / --------------------------------------------------------------------
        '// เลือก WorkSheet แล้วแสดงผลข้อมูลลงใน ListView
        Private Sub cmbSheetName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSheetName.SelectedIndexChanged
            Dim Conn As OleDbConnection
            Dim DA As OleDbDataAdapter
            Dim DT As New DataTable
            Dim strConn As String = _
                " Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                txtFileName.Text & ";" & _
                " Extended Properties=""Excel 12.0 Xml; HDR=YES"";"
            Try
                Conn = New OleDbConnection
                Conn.ConnectionString = strConn
                If Conn.State = ConnectionState.Closed Then Conn.Open()
                DA = New OleDbDataAdapter("Select * FROM [" & cmbSheetName.Text & "]", Conn)
                DA.Fill(DT)
                With ListView1
                    .Clear()
                    .View = View.Details
                    .GridLines = True
                    .FullRowSelect = True
                    .HideSelection = False
                    .MultiSelect = False
                End With
                '// อ่านจำนวนหลักทั้งหมดเข้ามาก่อน
                For iCol = 0 To DT.Columns.Count - 1
                    ListView1.Columns.Add(DT.Columns(iCol).ColumnName)
                    '// ปรับระยะความกว้างใหม่
                    ListView1.Columns(iCol).Width = ListView1.Width \ (DT.Columns.Count - 1)
                Next
                '// อ่านข้อมูลในแต่ละแถว
                For sRow = 0 To DT.Rows.Count - 1
                    Dim LV As New ListViewItem
                    Dim dRow As DataRow = DT.Rows(sRow)
                    LV = ListView1.Items.Add(dRow.Item(0))  '// --> Primary Node
                    For iCol = 1 To DT.Columns.Count - 1
                        LV.SubItems.Add(dRow(iCol).ToString())
                    Next
                Next
                Conn.Close()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            Finally
                Conn = Nothing
                DT = Nothing
                DA = Nothing
            End Try
        End Sub
