# first_github

 Sub calender_table()

Dim dm As Variant
Dim m, w, d, r As Integer

y = InputBox("input_year")
m_begin = InputBox("input_begin_month")


m_end = InputBox("input_end_month")
d = DateSerial(y, m_begin, 1)
w = Weekday(d)
r = 3

Range("1£º100").Delete
Cells(1, 1) = y & "year from" & m_begin & "to" & m_end & " calendar"
Cells(2, 1) = "星期日"
Cells(2, 2) = "星期一"
Cells(2, 3) = "星期二"
Cells(2, 4) = "星期三"
Cells(2, 5) = "星期四"
Cells(2, 6) = "星期五"
Cells(2, 7) = "星期六"
Range("a2:g2").Interior.ColorIndex = 3
Range("a1").Interior.ColorIndex = 50
Range("a1:g2").Font.Name = "黑体"
Range("a1:g2").Font.Bold = True
dm = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)

If ((y Mod 400 = 0) Or (y Mod 4 = 0 And y Mod 100 <> 0)) Then
dm(1) = 29
End If
For m = m_begin To m_end
   Cells(r, 1) = m & "月"
   Cells(r, 1).Interior.ColorIndex = 10
   Cells(r, 1).Font.Bold = True
   r = r + 1

    For d = 1 To dm(m - 1)
    Cells(r, w) = d
    w = w + 1
    If w > 7 Then
    w = 1
    r = r + 1
    End If
    Next
    r = r + 1
    
    
Next
End Sub

