    Function SetupAutoID() As String
        strSQL = _
            " SELECT MAX(Patient.HNPK) AS MaxPK FROM Patient "
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Cmd = New OleDb.OleDbCommand(strSQL, Conn)
        Dim MaxPK As Long
        '/ ตรวจสอบว่ามีข้อมูลอยู่หรือไม่
        If IsDBNull(Cmd.ExecuteScalar) Then
            MaxPK = 1
        Else
            MaxPK = Cmd.ExecuteScalar + 1
        End If
        '/ ตัวอย่างของการสร้างรูปแบบ ID อัตโนมัติ HN-ปีพ.ศ.XXXXX เช่น HN-6000013
        '/ เช่น ได้ MaxPK = 13 เอามาเรียงต่อกันกับ 0 จำนวน 5 ตัว ก็จะได้ "00000" & "13" = "0000013"
        '/ ให้นับกลับมาจากทางขวาเข้ามาทางซ้าย ก็จะได้ "00013"
        SetupAutoID = "HN-" & Microsoft.VisualBasic.Right(Year(Now), 2) & Microsoft.VisualBasic.Right("00000" & MaxPK, 5)
    End Function
    
    
    
    ' https://www.g2gnet.com/webboard/forum.php?mod=viewthread&tid=31&extra=page%3D1
