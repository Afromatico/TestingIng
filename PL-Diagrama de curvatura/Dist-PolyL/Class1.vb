Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors

Public Class Class1


    Dim tolerancia As Double = 10

    Dim H As Double = 500

    Dim V As Double = 100

    Dim Xorigin As Double = 0

    Dim Yorigin As Double = 0



    <CommandMethod("GENERARDIAGRAMADECURVATURA")> _
    Public Sub addDistance()

        Dim myDwg As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim myTransMan As Autodesk.AutoCAD.ApplicationServices.TransactionManager
        Dim mytrans As Transaction

        Dim result As Autodesk.AutoCAD.EditorInput.PromptEntityResult = getPolyline()
        Dim Status As Autodesk.AutoCAD.EditorInput.PromptStatus = result.Status

        Select Case Status
            Case Autodesk.AutoCAD.EditorInput.PromptStatus.OK

                Dim myent As DBObject
                Dim pline As Polyline

                myTransMan = myDwg.TransactionManager
                mytrans = myTransMan.StartTransaction
                myent = result.ObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                pline = CType(myent, Polyline)

                Dim pPtRes As PromptPointResult = importPoint()

                If pPtRes.Status = PromptStatus.OK Then

                    Dim point2 As Point3d = pPtRes.Value

                    Xorigin = point2.X
                    Yorigin = point2.Y

                    Dim analisis As ArrayList = parsePolyLine(pline)

                    Try
                        analisis = parsePolyLine(pline)
                    Catch ex As Exception
                        MsgBox("Error en el parser" & vbCrLf & ex.Message)
                        Return
                    End Try

                    Dim last As Tuple(Of Integer, String, Integer, Double, Double, Double) = New Tuple(Of Integer, String, Integer, Double, Double, Double)(0, "None", 0, 0, 0, 0)
                    Dim dis As Double = 0
                    Dim med As Double = 0

                    For Each tupla As Tuple(Of Integer, String, Integer, Double, Double, Double) In analisis

                        med = (dis * 2 + tupla.Item4) / 2

                        Select Case tupla.Item2
                            Case "Rect"
                                If last.Item2.Equals("None") Then

                                    drawLine(dis, 0, dis + tupla.Item4, 0, True, Color.FromRgb(255, 0, 0), 0, myDwg.Database)
                                    drawText(med, 0, "Recta en " + tupla.Item4.ToString("f") + " m.", True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)

                                Else
                                    Dim alpha As Double

                                    Dim point11 As Point3d = pline.GetLineSegmentAt(last.Item1).StartPoint
                                    Dim point22 As Point3d = pline.GetLineSegmentAt(last.Item1).EndPoint
                                    Dim point33 As Point3d = pline.GetLineSegmentAt(tupla.Item1).EndPoint

                                    Dim A As Double = point11.Y - point22.Y
                                    Dim B As Double = point11.X - point22.X

                                    Dim norm As Double = (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)))

                                    A = A / norm
                                    B = B / norm

                                    Dim A2 As Double = point22.Y - point33.Y
                                    Dim B2 As Double = point22.X - point33.X

                                    Dim norm2 As Double = (Math.Sqrt(Math.Pow(A2, 2) + Math.Pow(B2, 2)))

                                    A2 = A2 / norm
                                    B2 = B2 / norm

                                    Dim selec As Integer

                                    If (Math.Sqrt(Math.Pow(A + A2, 2) + Math.Pow(B + B2, 2))) > Math.Sqrt(2) Then
                                        If B > (B + B2) / (Math.Sqrt(Math.Pow(A + A2, 2) + Math.Pow(B + B2, 2))) Then
                                            If A > 0 Then
                                                selec = 3
                                            Else
                                                selec = 4
                                            End If
                                        Else
                                            If A > 0 Then
                                                selec = 4
                                            Else
                                                selec = 3
                                            End If
                                        End If

                                    Else
                                        If 0 > (B + B2) / (Math.Sqrt(Math.Pow(A + A2, 2) + Math.Pow(B + B2, 2))) Then
                                            If A > 0 Then
                                                selec = 2
                                            Else
                                                selec = 1
                                            End If
                                        Else
                                            If A > 0 Then
                                                selec = 1
                                            Else
                                                selec = 2
                                            End If
                                        End If
                                    End If

                                    Select Case selec
                                        Case 1
                                            alpha = 2 * Math.PI - Math.Acos(Math.Abs(A * A2 + B * B2) / (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)) * Math.Sqrt(Math.Pow(A2, 2) + Math.Pow(B2, 2))))
                                        Case 2
                                            alpha = Math.Acos(Math.Abs(A * A2 + B * B2) / (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)) * Math.Sqrt(Math.Pow(A2, 2) + Math.Pow(B2, 2))))
                                        Case 3
                                            alpha = Math.PI - Math.Acos(Math.Abs(A * A2 + B * B2) / (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)) * Math.Sqrt(Math.Pow(A2, 2) + Math.Pow(B2, 2))))
                                        Case 4
                                            alpha = Math.PI + Math.Acos(Math.Abs(A * A2 + B * B2) / (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)) * Math.Sqrt(Math.Pow(A2, 2) + Math.Pow(B2, 2))))

                                    End Select


                                    alpha = alpha * 200 / Math.PI

                                    drawText(med, 0, "Recta en " + tupla.Item4.ToString("f") + " m.", True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)
                                    drawText(dis, 0, "α=" + alpha.ToString("f"), True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)

                                    drawLine(dis, 0, dis + tupla.Item4, 0, True, Color.FromRgb(255, 0, 0), 0, myDwg.Database)

                                End If
                            Case "Arc"

                                Dim alpha2 As Double = 0

                                Dim alpha As Double = Math.Abs(tupla.Item5 * 200 / Math.PI - 200)
                                Dim T As Double = tupla.Item6 * Math.Tan(tupla.Item5 / 2)


                                If pline.GetBulgeAt(tupla.Item1) > 0 Then
                                    alpha2 = 750 / tupla.Item6
                                Else
                                    alpha2 = -750 / tupla.Item6
                                End If

                                drawText(med - 9 * H / 100, 6, "R=" + tupla.Item6.ToString("f"), True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)
                                drawText(med - 9 * H / 100, 1, "α=" + alpha.ToString("F4"), True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)
                                drawText(med + 9 * H / 100, 1, "T=" + T.ToString("F3"), True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)
                                drawText(med + 9 * H / 100, 6, "D=" + tupla.Item4.ToString("F3"), True, Color.FromRgb(0, 0, 0), 2.25, 0, AttachmentPoint.BottomCenter)

                                drawLine(dis, alpha2, dis + tupla.Item4, alpha2, True, Color.FromRgb(255, 0, 0), 0, myDwg.Database)
                            Case "Euler"

                            Case Else

                                MsgBox("Error en el dibujo" & vbCrLf)
                                Return

                        End Select

                        dis += tupla.Item4

                        last = tupla

                    Next

                    drawLine(0, 0, dis, 0, True, Color.FromRgb(0, 0, 0), 0, myDwg.Database)

                    mytrans.Commit()

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

    Private Function parsePolyLine(pline As Polyline) As ArrayList

        Dim parsePropperties As New ArrayList()

        Dim stype As SegmentType

        Dim countRectSegments As Integer = 0
        Dim countArcSegments As Integer = 0
        Dim countEulerSegments As Integer = 0

        Dim lastType As String = "None"

        For vertx As Integer = 0 To pline.NumberOfVertices - 1

            stype = pline.GetSegmentType(vertx)

            If stype = SegmentType.Line Then

                Dim lineseg As LineSegment2d = pline.GetLineSegment2dAt(vertx)

                If lineseg.Length < tolerancia Then
                    If lastType.Equals("None") Or lastType.Equals("Rect") Or lastType.Equals("Arc") Then
                        countEulerSegments += 1
                    ElseIf lastType.Equals("Euler") Then

                    Else
                        MsgBox("Error identificando polylinea")
                        Throw New System.Exception("An exception has occurred.")
                    End If

                    parsePropperties.Add(New Tuple(Of Integer, String, Integer, Double, Double, Double)(vertx, "Euler", countEulerSegments, lineseg.Length, 0, 0))
                    lastType = "Euler"

                Else
                    countRectSegments += 1
                    parsePropperties.Add(New Tuple(Of Integer, String, Integer, Double, Double, Double)(vertx, "Rect", countRectSegments, lineseg.Length, 0, 0))
                    lastType = "Rect"


                End If

            ElseIf stype = SegmentType.Arc Then

                Dim arcseg As CircularArc2d = pline.GetArcSegment2dAt(vertx)

                countArcSegments += 1
                parsePropperties.Add(New Tuple(Of Integer, String, Integer, Double, Double, Double)(vertx, "Arc", countArcSegments, arcseg.EndAngle * arcseg.Radius, arcseg.EndAngle, arcseg.Radius))
                lastType = "Arc"

            End If

        Next

        Return parsePropperties
    End Function

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
            newLine.AddVertexAt(0, New Point2d(x0 * 100 / H + Xorigin, y0 * 100 / V + Yorigin), 0, widht, widht)
            newLine.AddVertexAt(0, New Point2d(x1 * 100 / H + Xorigin, y1 * 100 / V + Yorigin), 0, widht, widht)
            If hasColor Then
                newLine.Color = color
            End If
            mSpace.AppendEntity(newLine)
            trans.AddNewlyCreatedDBObject(newLine, True)
            trans.Commit()

        End Using

    End Sub

    Private Function getPolyline() As PromptEntityResult

        Dim myPEO As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Select BarMark:")
        Dim mydwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim myed As Editor = mydwg.Editor

        myPEO.SetRejectMessage("Porfavor, seleciona una Polylinea." & vbCrLf)
        myPEO.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), False)
        Dim myPer As PromptEntityResult = myed.GetEntity(myPEO)

        Return myPer
    End Function

    Private Sub drawText(x0 As Double, y0 As Double, text As String, hasColor As Boolean, color As Color, height As Double, rotation As Double, justify As AttachmentPoint)

        Dim db As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim mSpace As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Using asMtext As MText = New MText()
                asMtext.Attachment = justify
                asMtext.SetAttachmentMovingLocation(asMtext.Attachment)
                asMtext.Location = New Point3d(x0 * 100 / H + Xorigin, y0 * 100 / V + Yorigin, 0)
                asMtext.Width = 55
                asMtext.Contents = text
                asMtext.Rotation = rotation
                If hasColor Then
                    asMtext.Color = color
                End If
                asMtext.TextHeight = height
                mSpace.AppendEntity(asMtext)
                trans.AddNewlyCreatedDBObject(asMtext, True)
            End Using
            trans.Commit()

        End Using

    End Sub

End Class
