Attribute VB_Name = "Módulo1"
Option Explicit
Option Base 0

Sub teste()
    Dim vvet As New Vandas, vaux As Variant
    
    vvet.data = Planilha1.Range("A1:E11").Value2

    Set vvet = vvet.column.push(Array(1, 2, 3, "4", "5"), 10)
    'vvet.printData
    For Each vaux In vvet.columns("Coluna A", "Coluna B")
        Debug.Print vaux
    Next vaux

End Sub

Sub routine_test()
    Dim vvet As New Vandas, vaux(-2 To 1, -5 To -3) As Variant, vitem As Variant
    
'    vaux(-2, -5) = "A"
'    vaux(-2, -4) = "B"
'    vaux(-2, -3) = "C"
'
'    vaux(-1, -5) = "D"
'    vaux(-1, -4) = "E"
'    vaux(-1, -3) = "F"
'
'    vaux(0, -5) = "G"
'    vaux(0, -4) = "H"
'    vaux(0, -3) = "I"
'
'    vaux(1, -5) = "J"
'    vaux(1, -4) = "K"
'    vaux(1, -3) = "L"
    
    vvet.data = Array("A", "B", "C", Array())
    vvet.printData
    Call vvet.column.push(Array(), 10)
    vvet.printData
    'For Each vitem In vvet.rows("A", "B")
    '    Debug.Print vitem
    'Next vitem
    'Set vvet = vvet.push(Array(1, 2, 3, "4", "5"), 100)
    'vvet.printData

End Sub
