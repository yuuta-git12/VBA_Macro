Attribute VB_Name = "Array_TEST"
Option Explicit

Sub sample1()
    Dim i As Long
    Application.ScreenUpdating = False
    For i = 1 To 100000
        Cells(i, 3) = Cells(i, 1) * Cells(i, 2)
    Next i
    Debug.Print ("èIóπ")
    Application.ScreenUpdating = False
End Sub

Sub sample2()
    Dim i As Long, t As Long, j As Long
    Dim MyArray1
    Dim MyArray2
    'MyArray1 = Range("A1:B100000")
    t = 1
    j = 10000
    MyArray1 = Range(Cells(t, 1), Cells(j, 2))
    ReDim MyArray2(1 To 100000, 1 To 1)
    For i = LBound(MyArray1, 1) To UBound(MyArray1, 1)
        MyArray2(i, 1) = MyArray1(i, 1) * MyArray1(i, 2)
    Next i
    
    Range("C1:C100000") = MyArray2
    Debug.Print ("èIóπ")
End Sub
