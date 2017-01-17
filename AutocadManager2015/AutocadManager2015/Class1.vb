﻿Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Colors
Imports Autodesk.AutoCAD.Geometry


Public Class DrawClass

    Dim xOrigin As Double
    Dim yOrigin As Double

    Dim hScale As Double
    Dim vScale As Double

    Dim layerList As List(Of MyLayer)

    Dim entityList As List(Of Entity)

    Public Sub New()

        Me.New(0, 0, 1, 1)

    End Sub

    Public Sub New(x As Double, y As Double, h As Double, v As Double)

        MyBase.New()

        xOrigin = x
        yOrigin = y

        hScale = h
        vScale = v

        layerList = New List(Of MyLayer)
        entityList = New List(Of Entity)

    End Sub

    Public Sub draw()

        putEntietiesInDraw(entityList, Application.DocumentManager.MdiActiveDocument.Database)

        MsgBox("Se han dibujado " + entityList.Count.ToString + " Objetos.")

        Me.Finalize()

    End Sub

    Public Sub addLayer(name As String, color As Color, textHeight As Double, linewidth As Double)

        If Not layerList.Exists(Function(x) x.compareNameLayer(name)) Then
            layerList.Add(New MyLayer(name, color, textHeight, linewidth))
            CreateAndAssignALayer(name, color)
        Else
            MsgBox("Layer ya se a asignado previamente")
        End If

    End Sub

    Public Sub addAtFirst(ent As Entity)

        entityList.Add(ent)

    End Sub

    Public Sub addAtLastEntity(ent As Entity)

        entityList.Insert(0, ent)

    End Sub

    Public Function drawText(x0 As Double, y0 As Double, text As String, layer As String, rotation As Double, justify As AttachmentPoint) As Entity


        Dim asMtext As MText = New MText()
        asMtext.Attachment = justify
        asMtext.SetAttachmentMovingLocation(asMtext.Attachment)
        asMtext.Location = New Point3d(x0 * hScale + xOrigin, y0 * vScale + yOrigin, 0)
        asMtext.Width = 55
        asMtext.Contents = text
        asMtext.Rotation = rotation
        asMtext.Layer = layer
        asMtext.TextHeight = layerList.Find(Function(x) x.compareNameLayer(layer)).getTextSize

        Return asMtext

    End Function

    Public Function drawLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, layer As String) As Entity

        Dim newLine As New Autodesk.AutoCAD.DatabaseServices.Polyline
        Dim widht As Double = layerList.Find(Function(x) x.compareNameLayer(layer)).getWidthLine

        newLine.AddVertexAt(0, New Point2d(x0 * hScale + xOrigin, y0 * vScale + yOrigin), 0, widht, widht)
        newLine.AddVertexAt(0, New Point2d(x1 * hScale + xOrigin, y1 * vScale + yOrigin), 0, widht, widht)
        newLine.Layer = layer

        Return newLine

    End Function

    Private Sub CreateAndAssignALayer(layer1 As String, color As Color)
        '' Get the current document and database
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Layer table for read
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, _
                                         OpenMode.ForRead)

            If acLyrTbl.Has(layer1) = False Then
                Using acLyrTblRec As LayerTableRecord = New LayerTableRecord()

                    acLyrTblRec.Color = color
                    acLyrTblRec.Name = layer1

                    acLyrTbl.UpgradeOpen()

                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                End Using
            End If

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub

    Private Sub putEntietiesInDraw(ListEntity As List(Of Entity), db As Database)

        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim mSpace As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

            For Each Entity As Entity In ListEntity
                mSpace.AppendEntity(Entity)
                trans.AddNewlyCreatedDBObject(Entity, True)
            Next
            trans.Commit()

        End Using

    End Sub

    Private Class MyLayer

        Dim name As String
        Dim color As Color

        Dim textSize As Double

        Dim widthLine As Double

        Public Sub New(nameString As String, colorL As Color, textHeight As Double, linewidth As Double)

            name = nameString
            color = colorL
            textSize = textHeight
            widthLine = linewidth

        End Sub

        Public Function compareNameLayer(nameLayer As String) As Boolean

            If nameLayer.Equals(Me.name) Then
                Return True
            End If

            Return False

        End Function

        Public Function getTextSize() As Double
            Return Me.textSize
        End Function

        Public Function getWidthLine() As Double
            Return Me.widthLine
        End Function



    End Class

End Class