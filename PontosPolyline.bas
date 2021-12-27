Attribute VB_Name = "Module1"
Sub PtsPoly()
'Declaração de variaveis
Dim XLnCADObject As AcadLWPolyline
Dim i As Integer
Dim x As Double, y As Double
Dim ConjCoords As Variant
Dim objEnt As AcadPoint
Dim Coords(0 To 2) As Double
Dim TabCoord() As Double
Dim contador As Integer
Dim Dx, Dy, D As Double

On Error GoTo Done
'Obtem as propriedades da polylinha
ThisDrawing.Utility.GetEntity XLnCADObject, basePnt, "Selecione a polylinha"
    
'Obtem o conjunto de coordenadas dos pontos da polyline
ConjCoords = XLnCADObject.Coordinates

contador = 0
For i = LBound(ConjCoords) To UBound(ConjCoords) Step 2
contador = contador + 1
Next i

ReDim TabCoord(contador - 1, 0 To 2)

'Coleta as coordenadas dos pontas da polyline
contador = 0
For i = LBound(ConjCoords) To UBound(ConjCoords) Step 2
    x = XLnCADObject.Coordinates(i)
    y = XLnCADObject.Coordinates(i + 1)
    If i > 1 Then
        For j = 0 To contador - 1
            Dx = (x - TabCoord(j, 0)) ^ 2
            Dy = (y - TabCoord(j, 1)) ^ 2
            D = Dx + Dy
            If D = 0 Then
                TabCoord(contador, 2) = 1
            End If
        Next j
    Else
        TabCoord(contador, 2) = 0
    End If
    TabCoord(contador, 0) = x
    TabCoord(contador, 1) = y
    contador = contador + 1
Next i

'Inseri pontos nas coordenadas dos pontos da polyline
ThisDrawing.SetVariable "PDMODE", 32 'Define o formato do ponto
For j = 0 To contador - 1
    If TabCoord(j, 2) = 0 Then
        Coords(0) = TabCoord(j, 0)
        Coords(1) = TabCoord(j, 1)
        Coords(2) = 0
        Set objEnt = ThisDrawing.ModelSpace.AddPoint(Coords)
    Else
    End If
    
Next j
'Pausa para continuar após o enter ser pressionado
ThisDrawing.Utility.GetString False, vbCr & "Pressione ENTER para continuar"
Done:

End Sub
