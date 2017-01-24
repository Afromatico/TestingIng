Imports AutocadManager2015
Imports Autodesk.AutoCAD.DatabaseServices
Imports Microsoft.Office.Interop.Excel

Public MustInherit Class Block

    Protected LayerTerreno As String = "PL-Terreno"
    Protected LayerPRC As String = "PL-PRC"
    Protected LayerTierra As String = "PL-Tierra"
    Protected LayerAcera As String = "PL-Acera"
    Protected LayerCalzada As String = "PL-Calzada"
    Protected LayerCiclovia As String = "PL-Ciclovia"
    Protected LayerAverde As String = "PL-AVerde"
    Protected LayerFFCC As String = "PL-FFCC"
    Protected LayerTextoTitulo As String = "TextoTitulo"
    Protected LayerTextoComentario As String = "TextoComentario"
    Protected LayerDimension As String = "Dimension"
    Protected LayerHatsh As String = "$Fondo"
    Protected LayerPostes As String = "POSTES"

    Protected celda1 As Object
    Protected celda2 As Object
    Protected celda3 As Object
    Protected celda4 As Object
    Protected celda5 As Object

    Protected drawDWG As DrawClass

    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        Me.celda1 = o1
        Me.celda2 = o2
        Me.celda3 = o3
        Me.celda4 = o4
        Me.celda5 = o5
        Me.drawDWG = drawDWG
    End Sub

    MustOverride Function calculeLenght() As Double

    MustOverride Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

End Class

Public MustInherit Class FinalBlock
    Inherits Block

    Private bool As Boolean

    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
        Me.bool = bool
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)


        If Me.calculeLenght > 0 Then
            drawDWG.addAtLastEntity(drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, Me.calculeLenght, 0, LayerPostes))
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1, tuple.Item2, 0.15, 0.2, celda1.ToString, LayerTextoComentario, 0, AttachmentPoint.BottomLeft))

            If bool Then
                drawEnd(tuple)
            Else
                drawEnd(New Tuple(Of Double, Double)(tuple.Item1 + Me.calculeLenght, tuple.Item2))
            End If

            Return New Tuple(Of Double, Double)(tuple.Item1 + Me.calculeLenght, tuple.Item2)
        End If

        drawEnd(tuple)

        Return tuple
    End Function

    MustOverride Sub drawEnd(tuple As Tuple(Of Double, Double))

    Sub drawTopComentary(tuple As Tuple(Of Double, Double))
        If Not String.IsNullOrEmpty(Me.celda3) Then
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1, tuple.Item2, 0, 0.2, Me.celda3.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
        End If
    End Sub

    Overrides Function calculeLenght() As Double

        If Not String.IsNullOrEmpty(Me.celda1) Then

            Return Me.celda1.ToString.Length * 0.2 + 0.3

        End If

        Return 0

    End Function

End Class

Public MustInherit Class ContructionBlock
    Inherits Block

    Protected h As Double
    Protected display As Display
    Protected mylayout As String
    Public drawColor As Boolean = True

    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
        Me.display = getDisplay()
        Me.h = h
    End Sub

    Overrides Function calculeLenght() As Double

        Return Me.celda4

    End Function

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        If String.IsNullOrEmpty(celda1) Then
            drawDWG.addAtFirst(drawDWG.drawDimension(tuple.Item1, tuple.Item2, tuple.Item1 + Me.calculeLenght, tuple.Item2, tuple.Item1 + Me.calculeLenght / 2, h, 0, LayerDimension))
        ElseIf celda1.Equals("VAR") Then
            Dim Ent As RotatedDimension = drawDWG.drawDimension(tuple.Item1, tuple.Item2, tuple.Item1 + Me.calculeLenght, tuple.Item2, tuple.Item1 + Me.calculeLenght / 2, h, 0, LayerDimension)
            Ent.DimensionText = "var"
            drawDWG.addAtFirst(Ent)
        ElseIf celda1.Equals("NM") Then

        End If

        If Not String.IsNullOrEmpty(celda3) Then
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1 + Me.calculeLenght / 2, tuple.Item2 - 0.8, celda3.ToString, LayerTextoComentario, 0, AttachmentPoint.BottomCenter))
        End If


        Return Me.display.draw(tuple)


    End Function

    Overridable Function getDisplay() As Display

        If String.IsNullOrEmpty(celda2) Then
            Return New Standart(celda4, drawDWG, Me)
        ElseIf celda2.Equals("R") Then
            Return New Reduced(celda4, drawDWG, Me)
        ElseIf celda2.Equals("NN") Then
            Return New NN(celda4, drawDWG, Me)
        ElseIf celda2.Equals("RIP") Then
            Return New Rip(celda4, drawDWG, Me)
        ElseIf celda2.ToString Like "IN/##*" Then
            Dim values As String() = celda2.Split("/")
            Return New Inclination(celda4, drawDWG, Me, values.GetValue(1))
        End If

        Return New Standart(celda4, drawDWG, Me)

    End Function

    Public Function getLayout() As String
        Return Me.mylayout
    End Function

End Class

Public Class FB
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Sub drawEnd(tuple As Tuple(Of Double, Double))

        drawDWG.addAtLastEntity(drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, 0, 1, LayerTerreno))

        Me.drawTopComentary(New Tuple(Of Double, Double)(tuple.Item1, tuple.Item2 + 1))

    End Sub

End Class

Public Class FS
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Sub drawEnd(tuple As Tuple(Of Double, Double))

        drawDWG.addAtLastEntity(drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, 0, 3, LayerTerreno))

        Me.drawTopComentary(New Tuple(Of Double, Double)(tuple.Item1, tuple.Item2 + 3))

    End Sub

End Class

Public Class F
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Sub drawEnd(tuple As Tuple(Of Double, Double))

        Me.drawTopComentary(tuple)

    End Sub

End Class

Public Class FPRC
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Sub drawEnd(tuple As Tuple(Of Double, Double))

        drawDWG.addAtLastEntity(drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, 0, 3, LayerPRC))
        drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1 - 0.2, tuple.Item2 + 3, "PRC", LayerPRC, Math.PI / 2, AttachmentPoint.BottomRight))


        Me.drawTopComentary(New Tuple(Of Double, Double)(tuple.Item1, tuple.Item2 + 3))

    End Sub

End Class

Public Class T
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerTierra
    End Sub
End Class

Public Class A
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerAcera
    End Sub
End Class

Public Class C
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerCalzada
    End Sub

    Overrides Function getDisplay() As Display

        If String.IsNullOrEmpty(celda2) Then
            Return New Reduced(celda4, drawDWG, Me)
        ElseIf celda2.Equals("P") Then
            Return New Standart(celda4, drawDWG, Me)
        ElseIf celda2.Equals("NN") Then
            Return New NN(celda4, drawDWG, Me)
        ElseIf celda2.Equals("Rip") Then
            Return New Rip(celda4, drawDWG, Me)
        ElseIf celda2.ToString Like "IN/##*" Then
            Dim values As String() = celda2.Split("/")
            Return New Inclination(celda4, drawDWG, Me, values.GetValue(1))
        End If

        Return New Reduced(celda4, drawDWG, Me)

    End Function
End Class

Public Class Ci
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerCiclovia
    End Sub
End Class

Public Class Av
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerAverde
    End Sub
End Class

Public Class FFCC
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerFFCC
    End Sub
End Class

Public Class L
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG, h)
        Me.mylayout = LayerPostes
        Me.drawColor = False
    End Sub
End Class

Public Class Com
    Inherits Block

    Protected h As Double
    Protected mylayout As String
    Dim main As MainClass
    Dim cj As Integer
    Dim xlSheet As Worksheet

    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass, h As Double, main As MainClass, ci As Integer, cj As Integer, xlSheet As Worksheet)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
        Me.mylayout = LayerPostes
        Me.main = main
        Me.cj = cj
        Me.xlSheet = xlSheet
        Me.h = h
    End Sub

    Overrides Function calculeLenght() As Double
        Return Me.celda4
    End Function

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)
        Dim lista As List(Of Block) = main.getNextCom(main.com, cj, xlSheet, h + 1)
        If String.IsNullOrEmpty(celda1) Then
            drawDWG.addAtFirst(drawDWG.drawDimension(tuple.Item1, tuple.Item2, tuple.Item1 + Me.calculeLenght, tuple.Item2, tuple.Item1 + Me.calculeLenght / 2, h, 0, LayerDimension))
        ElseIf celda2.Equals("VAR") Then
            Dim Ent As RotatedDimension = drawDWG.drawDimension(tuple.Item1, tuple.Item2, tuple.Item1 + Me.calculeLenght, tuple.Item2, tuple.Item1 + Me.calculeLenght / 2, h, 0, LayerDimension)
            Ent.DimensionText = "var"
            drawDWG.addAtFirst(Ent)
        End If

        Dim tuple2 As Tuple(Of Double, Double) = tuple
        For Each Block In lista
            tuple2 = Block.draw(tuple2)
        Next

        Return tuple2

    End Function

End Class


Public MustInherit Class Display

    Protected lenght As Double
    Protected drawDWG As DrawClass
    Protected myBlock As ContructionBlock

    Protected LayerTerreno As String = "PL-Terreno"
    Protected LayerPRC As String = "PL-PRC"
    Protected LayerTierra As String = "PL-Tierra"
    Protected LayerAcera As String = "PL-Acera"
    Protected LayerCalzada As String = "PL-Calzada"
    Protected LayerCiclovia As String = "PL-Ciclovia"
    Protected LayerAverde As String = "PL-AVerde"
    Protected LayerFFCC As String = "PL-FFCC"
    Protected LayerTextoTitulo As String = "TextoTitulo"
    Protected LayerTextoComentario As String = "TextoComentario"
    Protected LayerDimension As String = "Dimension"
    Protected LayerHatsh As String = "$Fondo"
    Protected LayerPostes As String = "POSTES"

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock)
        Me.lenght = lenght
        Me.drawDWG = drawDWG
        Me.myBlock = myBlock
    End Sub

    MustOverride Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

End Class

Public Class Standart
    Inherits Display

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock)
        MyBase.New(lenght, drawDWG, myBlock)
    End Sub

    Public Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, Me.lenght, 0, LayerTerreno))
        If myBlock.drawColor Then
            Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2 - 0.15, 0, 0, Me.lenght, 0, myBlock.getLayout))
        End If

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.lenght, tuple.Item2)
    End Function
End Class

Public Class Reduced
    Inherits Display

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock)
        MyBase.New(lenght, drawDWG, myBlock)
    End Sub

    Public Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, 0, -0.15, LayerTerreno))
        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1 + Me.lenght, tuple.Item2 - 0.15, 0, 0, 0, 0.15, LayerTerreno))

        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2 - 0.15, 0, 0, Me.lenght, 0, LayerTerreno))
        If myBlock.drawColor Then
            Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2 - 0.3, 0, 0, Me.lenght, 0, myBlock.getLayout))
        End If

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.lenght, tuple.Item2)
    End Function
End Class

Public Class Inclination
    Inherits Display

    Dim h As Double

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock, angle As Double)
        MyBase.New(lenght, drawDWG, myBlock)
        Me.h = Math.Tan(angle * Math.PI / 180) * Me.lenght
    End Sub

    Public Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, Me.lenght, h, LayerTerreno))
        If myBlock.drawColor Then
            Me.drawDWG.addAtLastEntity(Me.drawDWG.drawLine(tuple.Item1, tuple.Item2 - 0.15, 0, 0, Me.lenght, h, myBlock.getLayout))
        End If

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.lenght, tuple.Item2 + h)
    End Function
End Class

Public Class NN
    Inherits Display

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock)
        MyBase.New(lenght, drawDWG, myBlock)
    End Sub

    Public Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.lenght, tuple.Item2)
    End Function
End Class

Public Class Rip
    Inherits Display

    Sub New(lenght As Double, drawDWG As DrawClass, myBlock As ContructionBlock)
        MyBase.New(lenght, drawDWG, myBlock)
    End Sub


    Public Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Dim numberOfcuts As Integer = Convert.ToInt32(Math.Truncate(Me.lenght * 10))
        Dim cut As Integer = 1

        Dim bool As Boolean = False
        Dim lastTuple As Tuple(Of Double, Double) = tuple

        Dim lista As List(Of Tuple(Of Double, Double)) = New List(Of Tuple(Of Double, Double))
        lista.Add(tuple)
        lista.Add(New Tuple(Of Double, Double)(tuple.Item1 + lenght, tuple.Item2))

        While numberOfcuts > cut

            bool = False
            lastTuple = tuple

            Dim listaaux As List(Of Tuple(Of Integer, Tuple(Of Double, Double))) = New List(Of Tuple(Of Integer, Tuple(Of Double, Double)))
            Dim int As Integer = 0

            For Each tuple1 In lista
                If bool Then
                    Static Generator As System.Random = New System.Random()

                    Dim med As Double = (tuple1.Item2 + lastTuple.Item2) / 2

                    Dim rand As Double = Convert.ToDouble(Generator.Next(0, 100)) / 2000 - 0.025
                    While Math.Abs(tuple.Item2 - tuple1.Item2 - rand) > 0.2
                        rand = Convert.ToDouble(Generator.Next(0, 100)) / 2000 - 0.025
                    End While

                    listaaux.Add(New Tuple(Of Integer, Tuple(Of Double, Double))(lista.IndexOf(tuple1) + int, New Tuple(Of Double, Double)((tuple1.Item1 + lastTuple.Item1) / 2, (tuple1.Item2 + lastTuple.Item2) / 2 + rand)))
                    int += 1
                End If
                lastTuple = tuple1
                bool = True
            Next

            For Each tuple1 In listaaux
                lista.Insert(tuple1.Item1, tuple1.Item2)
            Next

            cut = cut * 2
        End While

        bool = False
        lastTuple = tuple

        Dim newpoly As Polyline = New Polyline()

        Me.drawDWG.addAtLastEntity(Me.drawDWG.drawPolyLine(0, 0, lista, LayerTerreno))
        If myBlock.drawColor Then
            Me.drawDWG.addAtLastEntity(Me.drawDWG.drawPolyLine(0, -0.15, lista, myBlock.getLayout))
        End If

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.lenght, tuple.Item2)
    End Function
End Class