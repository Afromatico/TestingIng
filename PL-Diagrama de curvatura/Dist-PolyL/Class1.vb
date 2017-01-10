Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors

Public Class Class1




    <CommandMethod("ADDDISTANCE")> _
    Public Sub addDistance()
        selectSelecction()
        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Se ha ejecutado con exito el programa")
    End Sub

    Private Sub selectSelecction()

        Dim myPEO As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select BarMark:")
        Dim mydwg, mydb, myed, myPS, myPer, myent, mytrans, mytransman As Object
        mydwg = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        mydb = mydwg.Database
        myed = mydwg.Editor
        myPEO.SetRejectMessage("Porfavor, seleciona una Polylinea." & vbCrLf)
        myPEO.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), False)
        myPer = myed.GetEntity(myPEO)
        myPS = myPer.Status
        Select Case myPS
            Case Autodesk.AutoCAD.EditorInput.PromptStatus.OK
                Dim pline As Polyline
                mytransman = mydwg.TransactionManager
                mytrans = mytransman.StartTransaction
                myent = myPer.ObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                pline = CType(myent, Polyline)
                MsgBox(pline.Length.ToString)
                pline.List()

                For vertx As Integer = 0 To pline.NumberOfVertices - 1

                    MsgBox("Coordenadas :" + pline.GetPoint3dAt(vertx).X.ToString + ", " + pline.GetPoint3dAt(vertx).Y.ToString + ", " + pline.GetPoint3dAt(vertx).Z.ToString & vbCrLf)

                    MsgBox(pline.GetBulgeAt(vertx).ToString & vbCrLf)
                    If Not pline.GetBulgeAt(vertx) = 0 Then
                        MsgBox(pline.GetArcSegmentAt(vertx).Radius.ToString & vbCrLf)
                        MsgBox(pline.GetArcSegmentAt(vertx).StartAngle.ToString & vbCrLf)
                        MsgBox(pline.GetArcSegmentAt(vertx).EndAngle.ToString & vbCrLf)
                        MsgBox((pline.GetArcSegmentAt(vertx).Radius * (pline.GetArcSegmentAt(vertx).StartAngle - pline.GetArcSegmentAt(vertx).EndAngle)).ToString)
                    End If


                Next

                pline.GetPoint3dAt(0)

                Dim pPtRes As PromptPointResult = importPoint()

                If pPtRes.Status = PromptStatus.OK Then
                    Dim point2 As Point3d = pPtRes.Value
                    Dim point As Point3d = pline.GetClosestPointTo(point2, True)
                    MsgBox("Coordenadas :" + point.X.ToString + ", " + point.Y.ToString + ", " + point.Z.ToString & vbCrLf)
                    mytrans.Commit()
                    drawLine(point.X, point.Y, point2.X, point2.Y, True, Color.FromRgb(255, 0, 0), 0.5, mydb)
                Else
                    mytrans.Commit()
                End If
            Case Autodesk.AutoCAD.EditorInput.PromptStatus.Cancel
                MsgBox("You cancelled.")
                Exit Sub
            Case Autodesk.AutoCAD.EditorInput.PromptStatus.Error
                MsgBox("Error warning.")
                Exit Sub
            Case Else
                Exit Sub
        End Select

    End Sub

    Private Function importPoint() As PromptPointResult

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        pPtOpts.Message = vbLf & "Seleccione un punto :"
        Return acDoc.Editor.GetPoint(pPtOpts)

    End Function

    Private Sub drawLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, hasColor As Boolean, color As Color, widht As Double, db As Database)

        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim mSpace As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim newLine As New Autodesk.AutoCAD.DatabaseServices.Polyline
            newLine.AddVertexAt(0, New Point2d(x0, y0), 0, widht, widht)
            newLine.AddVertexAt(0, New Point2d(x1, y1), 0, widht, widht)
            If hasColor Then
                newLine.Color = color
            End If
            mSpace.AppendEntity(newLine)
            trans.AddNewlyCreatedDBObject(newLine, True)
            trans.Commit()

        End Using

    End Sub

End Class
