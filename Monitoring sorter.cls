VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub septupdate_Click()
Dim a As Integer, b As Integer, c As Integer, d As Integer, e As Integer
Dim f As Integer, g As Integer, h As Integer, i As Integer, j As Integer
Dim k  As Integer, l As Integer, m As Integer, n As Integer, o As Integer
Dim p As Integer, q As Integer, r As Integer, s As Integer, t As Integer
Dim u As Integer, v As Integer, w As Integer, x As Integer, y As Integer
Dim z As Integer, aa As Integer, bb As Integer, cc As Integer, dd As Integer
Dim ee As Integer, ff As Integer, code As Variant

Application.ScreenUpdating = False

For a = 4 To 43                                     'week one weekly am tasks
code = Sheet9.Cells(a, 3).Value
    For b = 4 To 43
    If Sheet8.Cells(b, 3) = code Then
    Sheet9.Cells(b, 4) = Sheet8.Cells(b, 5)
    End If
    Next b
Next a

For c = 4 To 43                                     'week one weekly pm tasks
code = Sheet9.Cells(c, 6).Value
    For d = 4 To 43
    If Sheet8.Cells(d, 6) = code Then
    Sheet9.Cells(d, 7) = Sheet8.Cells(d, 8)
    End If
    Next d
Next c

For e = 4 To 43                                     'week one 2 weekly tasks
code = Sheet9.Cells(e, 9).Value
    For f = 4 To 43
    If Sheet8.Cells(f, 9) = code Then
    Sheet9.Cells(f, 10) = Sheet8.Cells(f, 11)
    End If
    Next f
Next e

For g = 4 To 43                                     'week one monthly tasks
code = Sheet9.Cells(g, 12).Value
    For h = 4 To 43
    If Sheet8.Cells(h, 12) = code Then
    Sheet9.Cells(h, 13) = Sheet8.Cells(h, 14)
    End If
    Next h
Next g

Application.ScreenUpdating = False

For i = 46 To 81                                     'week two weekly am tasks
code = Sheet9.Cells(i, 3).Value
    For j = 46 To 81
    If Sheet8.Cells(j, 3) = code Then
    Sheet9.Cells(j, 4) = Sheet8.Cells(j, 5)
    End If
    Next j
Next i

For k = 46 To 81                                     'week two weekly pm tasks
code = Sheet9.Cells(k, 6).Value
    For l = 46 To 81
    If Sheet8.Cells(l, 6) = code Then
    Sheet9.Cells(l, 7) = Sheet8.Cells(l, 8)
    End If
    Next l
Next k

For m = 46 To 81                                     'week two 2 weekly tasks
code = Sheet9.Cells(m, 9).Value
    For n = 46 To 81
    If Sheet8.Cells(n, 9) = code Then
    Sheet9.Cells(n, 10) = Sheet8.Cells(n, 11)
    End If
    Next n
Next m

For o = 46 To 81                                     'week two monthly tasks
code = Sheet9.Cells(o, 12).Value
    For p = 46 To 81
    If Sheet8.Cells(p, 12) = code Then
    Sheet9.Cells(p, 13) = Sheet8.Cells(p, 14)
    End If
    Next p
Next o

Application.ScreenUpdating = False

For q = 84 To 123                                     'week three weekly am tasks
code = Sheet9.Cells(q, 3).Value
    For r = 84 To 123
    If Sheet8.Cells(r, 3) = code Then
    Sheet9.Cells(r, 4) = Sheet8.Cells(r, 5)
    End If
    Next r
Next q

For s = 84 To 123                                     'week three weekly pm tasks
code = Sheet9.Cells(s, 6).Value
    For t = 84 To 123
    If Sheet8.Cells(t, 6) = code Then
    Sheet9.Cells(t, 7) = Sheet8.Cells(t, 8)
    End If
    Next t
Next s

For u = 84 To 123                                     'week three 2 weekly tasks
code = Sheet9.Cells(u, 9).Value
    For v = 84 To 123
    If Sheet8.Cells(v, 9) = code Then
    Sheet9.Cells(v, 10) = Sheet8.Cells(v, 11)
    End If
    Next v
Next u

For w = 84 To 123                                     'week three monthly tasks
code = Sheet9.Cells(w, 12).Value
    For x = 84 To 123
    If Sheet8.Cells(x, 12) = code Then
    Sheet9.Cells(x, 13) = Sheet8.Cells(x, 14)
    End If
    Next x
Next w


Application.ScreenUpdating = False

For y = 126 To 163                                     'week four weekly am tasks
code = Sheet9.Cells(y, 3).Value
    For z = 126 To 163
    If Sheet8.Cells(z, 3) = code Then
    Sheet9.Cells(z, 4) = Sheet8.Cells(z, 5)
    End If
    Next z
Next y

For aa = 126 To 163                                     'week four weekly pm tasks
code = Sheet9.Cells(aa, 6).Value
    For bb = 126 To 163
    If Sheet8.Cells(bb, 6) = code Then
    Sheet9.Cells(bb, 7) = Sheet8.Cells(bb, 8)
    End If
    Next bb
Next aa

For cc = 126 To 163                                     'week four 2 weekly tasks
code = Sheet9.Cells(cc, 9).Value
    For dd = 126 To 163
    If Sheet8.Cells(dd, 9) = code Then
    Sheet9.Cells(dd, 10) = Sheet8.Cells(dd, 11)
    End If
    Next dd
Next cc

For ee = 126 To 163                                     'week four monthly tasks
code = Sheet9.Cells(ee, 12).Value
    For ff = 126 To 163
    If Sheet8.Cells(ff, 12) = code Then
    Sheet9.Cells(ff, 13) = Sheet8.Cells(ff, 14)
    End If
    Next ff
Next ee

Application.ScreenUpdating = True

End Sub
