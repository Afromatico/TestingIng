Imports AutocadManager2015
Imports Autodesk.AutoCAD.DatabaseServices

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

    Protected bool As Boolean

    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
        Me.bool = bool
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)


        If Me.calculeLenght > 0 Then
            drawDWG.addAtLastEntity(drawDWG.drawLine(tuple.Item1, tuple.Item2, 0, 0, Me.calculeLenght, 0, LayerPostes))
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1, tuple.Item2, 0.15, 0.2, celda4.ToString, LayerTextoComentario, 0, AttachmentPoint.BottomLeft))
        End If

        Return New Tuple(Of Double, Double)(tuple.Item1 + Me.calculeLenght, tuple.Item2)
    End Function

    Overrides Function calculeLenght() As Double

        If Not String.IsNullOrEmpty(Me.celda1) Then

            Return Me.celda1.ToString.Length * 0.2 + 0.3

        End If

        Return 0

    End Function

End Class

Public MustInherit Class ContructionBlock
    Inherits Block
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function calculeLenght() As Double

        Return Me.celda4

    End Function

End Class

Public Class FB
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        If bool Then
            drawDWG.addAtLastEntity(drawDWG.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1, tuple.Item2, 0, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
        Else
            drawDWG.addAtLastEntity(drawDWG.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))
            drawDWG.addAtFirst(drawDWG.drawText(tuple.Item1, tuple.Item2, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
        End If

        Return MyBase.draw(tuple)
    End Function

End Class

Public Class FS
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class F
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class FPRC
    Inherits FinalBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, bool As Boolean, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, bool, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class T
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class A
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class C
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class Ci
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class Av
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class FFCC
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class L
    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class

Public Class Com

    Inherits ContructionBlock
    Sub New(o1 As Object, o2 As Object, o3 As Object, o4 As Object, o5 As Object, drawDWG As DrawClass)
        MyBase.New(o1, o2, o3, o4, o5, drawDWG)
    End Sub

    Overrides Function draw(tuple As Tuple(Of Double, Double)) As Tuple(Of Double, Double)

        Return New Tuple(Of Double, Double)(0, 0)
    End Function

End Class
