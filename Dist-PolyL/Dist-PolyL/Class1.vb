Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors

Public Class Class1

    <CommandMethod("ADDDISTANCE")> _
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
                    Dim point As Point3d = pline.GetClosestPointTo(point2, False)

                    Dim param As Double = pline.GetParameterAtPoint(point)
                    Dim index As Integer = CInt(Math.Truncate(param))

                    Dim Lenght As Double = 0

                    Dim stype As SegmentType

                    For vertx As Integer = 0 To index - 1

                        stype = pline.GetSegmentType(vertx)

                        If stype = SegmentType.Line Then

                            Dim lineseg As LineSegment2d = pline.GetLineSegment2dAt(vertx)
                            Lenght += lineseg.Length

                        ElseIf stype = SegmentType.Arc Then

                            Dim lineseg As CircularArc2d = pline.GetArcSegment2dAt(vertx)
                            Dim Radio, beta As Double
                            Radio = lineseg.Radius
                            beta = lineseg.EndAngle
                            Lenght += Math.Abs(Radio * beta)

                        End If

                    Next

                    stype = pline.GetSegmentType(index)

                    If stype = SegmentType.Line Then

                        Dim lineseg As LineSegment2d = pline.GetLineSegment2dAt(index)


                        Lenght += Math.Abs(Math.Sqrt(Math.Pow(lineseg.StartPoint.X - point.X, 2) + Math.Pow(lineseg.StartPoint.Y - point.Y, 2)))



                    ElseIf stype = SegmentType.Arc Then

                        Dim lineseg As CircularArc2d = pline.GetArcSegment2dAt(index)
                        Dim Radio, alpha, beta As Double

                        Radio = lineseg.Radius

                        Dim dis1 As Double = (lineseg.StartPoint.Y - lineseg.Center.Y)

                        Dim dis2 As Double = (lineseg.StartPoint.X - lineseg.Center.X)

                        Dim dis3 As Double = (point.Y - lineseg.Center.Y)

                        Dim dis4 As Double = (point.X - lineseg.Center.X)

                        If dis1 >= 0 And dis2 >= 0 Then
                            alpha = Math.Atan(dis1 / dis2)
                        ElseIf dis1 < 0 And dis2 >= 0 Then
                            alpha = 2 * Math.PI - Math.Atan(-dis1 / dis2)
                        ElseIf dis1 >= 0 And dis2 < 0 Then
                            alpha = Math.PI - Math.Atan(-dis1 / dis2)
                        Else
                            alpha = Math.PI + Math.Atan(dis1 / dis2)
                        End If

                        If dis3 >= 0 And dis4 >= 0 Then
                            beta = Math.Atan(dis3 / dis4)
                        ElseIf dis3 < 0 And dis4 >= 0 Then
                            beta = 2 * Math.PI - Math.Atan(-dis3 / dis4)
                        ElseIf dis3 >= 0 And dis4 < 0 Then
                            beta = Math.PI - Math.Atan(-dis3 / dis4)
                        Else
                            beta = Math.PI + Math.Atan(dis3 / dis4)
                        End If


                        If pline.GetBulgeAt(index) >= 0 Then
                            If beta < alpha Then
                                beta += 2 * Math.PI
                            End If
                        Else
                            If beta > alpha Then
                                beta += -2 * Math.PI
                            End If
                        End If

                        MsgBox(alpha.ToString)
                        MsgBox(beta.ToString)

                        Lenght += Math.Abs(Radio * Math.Abs(alpha - beta))

                        End If

                        mytrans.Commit()

                        Dim m As Double = (point.Y - point2.Y) / (point.X - point2.X)
                        Dim c As Double = point.Y - m * point.X

                        drawLine(point.X, point.Y, point2.X, point2.Y, True, Color.FromRgb(255, 0, 0), 0, myDwg.Database)

                        drawLine(point.X, point.Y, point2.X + 5 * Math.Cos(Math.Atan(m)), point2.Y + 5 * Math.Sin(Math.Atan(m)), True, Color.FromRgb(255, 0, 0), 0, myDwg.Database)

                        drawText(point2.X + 5 * Math.Cos(Math.Atan(m)), point2.Y + 5 * Math.Sin(Math.Atan(m)), Lenght.ToString("f"), True, Color.FromRgb(255, 0, 0), 1.5, Math.Atan(m), AttachmentPoint.BottomCenter)

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
                asMtext.Location = New Point3d(x0, y0, 0)
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
