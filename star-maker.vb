Public Sub star()
Dim pi, a, hw, hd As Long
Dim x0, y0, x1, y1, x, y, min, max, r, s, angle As String
Dim ax, ay, bx, by, cx, cy, dx, dy, ex, ey As Long
Dim c1, c2, c3, z1, z2, z3 As Long
Dim color As String


'Rows.RowHeight = 14 '行高
'Columns.ColumnWidth = 1.8 '`列宽
Rows.RowHeight = 9 '行高
Columns.ColumnWidth = 1 '`列宽



c1 = Application.RandBetween(100, 255)
c2 = Application.RandBetween(100, 255)
c3 = Application.RandBetween(100, 255)
'c1 = 255
'c2 = 255
'c3 = 255
z1 = 1
z2 = 1
z3 = 1
color = RGB(c1, c2, c3)

'color = 65535





pi = Application.pi()
x0 = 3
y0 = 5
hd = 60
x0 = Application.RandBetween(0, 50)
y0 = Application.RandBetween(0, 100)
hd = Application.RandBetween(20, 80)

a = hd / (Cos(54 / 180 * pi) + Sin(72 / 180 * pi))
hw = 2 * a * Sin(54 / 180 * pi)
x1 = x0 + hd
y1 = y0 + hw

ax = x0
ay = y0 + a * Sin(54 / 180 * pi)
bx = x0 + a * Cos(54 / 180 * pi)
by = y0
cx = x0 + a * Cos(54 / 180 * pi)
cy = y0 + 2 * a * Sin(54 / 180 * pi)
dx = x0 + hd
dy = y0 + a * Cos(72 / 180 * pi)
ex = x0 + hd
ey = y0 + 2 * a * Sin(54 / 180 * pi) - a * Cos(72 / 180 * pi)





x = x0
y = y0
angle = (180 - 72 - 72) / 180 * pi  '星星角36°

Do While x >= x0 And x < bx '注意这里用and不要用&！！！
Do While y <= y1
r = Sin(angle / 2) * (x - x0)
min = ay - r
max = ay + r
If y >= min And y <= max Then '注意这里用and不要用&！！！
'c1 = Application.RandBetween(0, 255)'搭配使用-A组
'c2 = Application.RandBetween(0, 255)'搭配使用-A组
z1 = Application.RandBetween(-1, 1)
z2 = Application.RandBetween(-1, 1)
z3 = Application.RandBetween(-1, 1)
c1 = c1 - z1
c2 = c2 - z2
c3 = c3 - z3
color = RGB(c1, c2, c3)
Cells(x, y).Interior.color = color
End If
y = y + 1
Loop
y = y0
x = x + 1
Loop





'Do While x >= bx And x < ex - Sin(2 * angle) * (ey - dy) / 2 / Cos(angle) 'And x < bx + hw * Tan(angle) / 2
'Do While y <= y0 + hw / 2
's = (x - bx) / y
'min = 0
'max = Tan(angle)
'If s >= min And s < max Then
'Cells(x, y).Interior.Color = color
'End If
'y = y + 1
'Loop
'y = y0
'x = x + 1
'Loop

Do While x >= bx And x < ex - Sin(2 * angle) * (ey - dy) / 2 / Cos(angle) 'And x < bx + hw * Tan(angle) / 2
Do While y <= y1
r = (x - bx) / Tan(angle)
min = by + r
max = cy - r
If y >= min And y <= max Then
'c1 = Application.RandBetween(0, 255)'搭配使用-A组
'c2 = Application.RandBetween(0, 255)'搭配使用-A组
z1 = Application.RandBetween(-1, 1)
z2 = Application.RandBetween(-1, 1)
z3 = Application.RandBetween(-1, 1)
c1 = c1 - z1
c2 = c2 - z2
c3 = c3 - z3
color = RGB(c1, c2, c3)
Cells(x, y).Interior.color = color
End If
y = y + 1
Loop
y = y0
x = x + 1
Loop





'Do While x >= ex - Sin(2 * angle) * (ey - dy) / 2 / Cos(angle) And x <= x1
'Do While y < y1
's = (dx - x) / (y - dy)
'min = (ex - cx) / ey
'max = hd / (ay - dy)
'If s >= min And s <= max Then
'Cells(x, y).Interior.Color = color
'End If
's = (ex - x) / (ey - y + 0.0001)
'If s >= min And s <= max Then
'Cells(x, y).Interior.Color = color
'End If
'y = y + 1
'Loop
'y = y0
'x = x + 1
'Loop
Do While x >= ex - Sin(2 * angle) * (ey - dy) / 2 / Cos(angle) And x <= x1
Do While y < y1
s = (dx - x) / (y - dy)
min = Tan(angle)
max = Tan(2 * angle)
If s >= min And s <= max Then
'c1 = Application.RandBetween(0, 255)'搭配使用-A组
'c2 = Application.RandBetween(0, 255)'搭配使用-A组
z1 = Application.RandBetween(-1, 1)
z2 = Application.RandBetween(-1, 1)
z3 = Application.RandBetween(-1, 1)
c1 = c1 - z1
c2 = c2 - z2
c3 = c3 - z3
color = RGB(c1, c2, c3)
Cells(x, y).Interior.color = color
End If
s = (ex - x) / (ey - y + 0.0001)
If s >= min And s <= max Then
'c1 = Application.RandBetween(0, 255)
'c2 = Application.RandBetween(0, 255)
z1 = Application.RandBetween(-1, 1)
z2 = Application.RandBetween(-1, 1)
z3 = Application.RandBetween(-1, 1)
c1 = c1 - z1
c2 = c2 - z2
c3 = c3 - z3
color = RGB(c1, c2, c3)
Cells(x, y).Interior.color = color
End If
y = y + 1
Loop
y = y0
x = x + 1
Loop





'Cells(ax, ay).Interior.color = color
'Cells(bx, by).Interior.color = color
'Cells(cx, cy).Interior.color = color
'Cells(dx, dy).Interior.color = color
'Cells(ex, ey).Interior.color = color

'Cells(ax, ay).Interior.Color = RGB(229, 245, 255)

End Sub

