Public Sub ball()
Dim x0, y0, x1, y1, x2, y2, x3, y3, x, y As Long
Dim r1, r2, r As Long
Dim c1, c2, c3, d1, d2, d3, z1, z2, z3, n As Long
Dim color As String
Dim m2, m3 As Long


'Rows.RowHeight = 14 '行高
'Columns.ColumnWidth = 1.8 '`列宽
Rows.RowHeight = 9 '行高
Columns.ColumnWidth = 1 '`列宽


'x0 = 3
'y0 = 5
'r = 40
x0 = Application.RandBetween(0, 100)
y0 = Application.RandBetween(0, 200)
r = Application.RandBetween(20, 40)

x1 = x0 + 2 * r
y1 = y0 + 2 * r
x2 = x0 + r
y2 = y0 + r
x3 = x2 - 0.3 * r
y3 = y2 - 0.3 * r

'r1 = 50
r1 = r


c1 = Application.RandBetween(0, 120)
c2 = Application.RandBetween(0, 120)
c3 = Application.RandBetween(0, 150)


z1 = (255 - c1) / r1 '单位单元格色差
z2 = (255 - c2) / r1
z3 = (255 - c3) / r1


x = x0
y = y0
Do While y <= y1
Do While x <= x1
m2 = (x - x2) ^ 2 + (y - y2) ^ 2 - r ^ 2
m3 = (x - x3) ^ 2 + (y - y3) ^ 2 - r1 ^ 2
If m2 <= 0 Then
If m3 <= 0 Then
n = ((x - x3) ^ 2 + (y - y3) ^ 2) ^ 0.5
d1 = c1 + (r1 - n) * z1
d2 = c2 + (r1 - n) * z2
d3 = c3 + (r1 - n) * z3
color = RGB(d1, d2, d3)
Cells(x, y).Interior.color = color
Else
color = RGB(c1, c2, c3)
Cells(x, y).Interior.color = color
End If
End If
x = x + 1
Loop
x = x0
y = y + 1
Loop



End Sub

