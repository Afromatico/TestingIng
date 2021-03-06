﻿Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Colors
Imports AutocadManager2015

Public Class Class1

    Dim draw As DrawClass

    Dim tolerancia As Double = 10

    Dim H As Double = 2

    Dim V As Double = 1

    Dim Xorigin As Double = 0

    Dim Yorigin As Double = 0

    Dim mm As Double = 750

    Dim Rmin As Double = 10000

    Dim heightText As Double = 1

    <CommandMethod("DIAG_CURVATURA")> _
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

                Dim listEntities As List(Of Entity) = New List(Of Entity)

                myTransMan = myDwg.TransactionManager
                mytrans = myTransMan.StartTransaction

                myent = result.ObjectId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                pline = CType(myent, Polyline)

                Dim pPtRes As PromptPointResult = importPoint()

                Try

                    If pPtRes.Status = PromptStatus.OK Then

                        Dim point2 As Point3d = pPtRes.Value

                        Dim layer1 As String = "TextLayer"
                        Dim layer2 As String = "RedLineLayer"
                        Dim layer3 As String = "BlackLineLayer"

                        draw = New DrawClass(point2.X, point2.Y, H, V)

                        draw.addLayer(layer1, Color.FromRgb(0, 0, 0), heightText, 0)
                        draw.addLayer(layer2, Color.FromRgb(255, 0, 0), heightText, 0)
                        draw.addLayer(layer3, Color.FromRgb(0, 0, 0), heightText, 0)

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

                        Dim lastNotEulesDis As Double = 0
                        Dim accEulerDis As Double = 0
                        Dim lastNotEuler As Tuple(Of Integer, String, Integer, Double, Double, Double) = New Tuple(Of Integer, String, Integer, Double, Double, Double)(0, "None", 0, 0, 0, 0)

                        Dim height As Double = 0
                        Dim lastHeight As Double = 0

                        For Each tupla As Tuple(Of Integer, String, Integer, Double, Double, Double) In analisis

                            med = (dis * 2 + tupla.Item4) / 2

                            Select Case tupla.Item2
                                Case "Rect"
                                    If Not last.Item2.Equals("Rect") Then

                                        draw.addAtFirst(draw.drawText(med, 1, "Recta en " + tupla.Item4.ToString("f") + " m.", layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtLastEntity(draw.drawLine(dis, 0, dis + tupla.Item4, 0, layer2))

                                        height = 0

                                    Else

                                        Dim alpha As Double

                                        Dim point11 As Point3d = pline.GetLineSegmentAt(last.Item1).StartPoint
                                        Dim point22 As Point3d = pline.GetLineSegmentAt(last.Item1).EndPoint
                                        Dim point33 As Point3d = pline.GetLineSegmentAt(tupla.Item1).EndPoint

                                        alpha = angle3Points(point11, point22, point33)

                                        draw.addAtFirst(draw.drawText(med, 1, "Recta en " + tupla.Item4.ToString("f") + " m.", layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(dis, 1, "α=" + alpha.ToString("f"), layer1, 0, AttachmentPoint.BottomCenter))

                                        draw.addAtLastEntity(draw.drawLine(dis, 0, dis + tupla.Item4, 0, layer2))

                                        height = 0

                                    End If
                                Case "Arc"

                                    Dim alpha2 As Double = 0

                                    Dim alpha As Double = Math.Abs(tupla.Item5 * 200 / Math.PI - 200)
                                    Dim T As Double = tupla.Item6 * Math.Tan(tupla.Item5 / 2)


                                    If pline.GetBulgeAt(tupla.Item1) > 0 Then
                                        alpha2 = mm / tupla.Item6

                                        draw.addAtFirst(draw.drawText(-9, 0, med, -3, "R=" + tupla.Item6.ToString("f"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(-9, 0, med, -1, "α=" + alpha.ToString("F4"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(9, 0, med, -1, "T=" + T.ToString("F3"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(9, 0, med, -3, "D=" + tupla.Item4.ToString("F3"), layer1, 0, AttachmentPoint.BottomCenter))

                                    Else
                                        alpha2 = -mm / tupla.Item6

                                        draw.addAtFirst(draw.drawText(-9, 0, med, +3, "R=" + tupla.Item6.ToString("f"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(-9, 0, med, +1, "α=" + alpha.ToString("F4"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(9, 0, med, +1, "T=" + T.ToString("F3"), layer1, 0, AttachmentPoint.BottomCenter))
                                        draw.addAtFirst(draw.drawText(9, 0, med, +3, "D=" + tupla.Item4.ToString("F3"), layer1, 0, AttachmentPoint.BottomCenter))

                                    End If

                                    draw.addAtLastEntity(draw.drawLine(dis, alpha2, dis + tupla.Item4, alpha2, layer2))

                                    height = alpha2

                                Case "Euler"

                                    If Not last.Item2.Equals("Euler") Then
                                        accEulerDis = tupla.Item4
                                        lastNotEuler = last
                                        lastNotEulesDis = dis
                                    Else
                                        accEulerDis += tupla.Item4
                                    End If

                                    If analisis.Count >= tupla.Item1 + 1 AndAlso (analisis.Count = tupla.Item1 + 1 OrElse Not analisis.Item(tupla.Item1 + 1).Item2.Equals("Euler")) Then

                                        Dim nextTuple As Tuple(Of Integer, String, Integer, Double, Double, Double)

                                        If Not analisis.Count = tupla.Item1 + 1 Then
                                            nextTuple = analisis.Item(tupla.Item1 + 1)
                                        Else
                                            nextTuple = New Tuple(Of Integer, String, Integer, Double, Double, Double)(0, "None", 0, 0, 0, 0)
                                        End If

                                        If lastNotEuler.Item2.Equals("Arc") And Not nextTuple.Item2.Equals("Arc") Then

                                            Dim alpha2 As Double = 0

                                            If pline.GetBulgeAt(lastNotEuler.Item1) > 0 Then
                                                alpha2 = mm / lastNotEuler.Item6
                                            Else
                                                alpha2 = -mm / lastNotEuler.Item6
                                            End If

                                            Dim cal As Double = Math.Sqrt(accEulerDis * lastNotEuler.Item6)

                                            draw.addAtFirst(draw.drawText(lastNotEulesDis + accEulerDis / 2, 1, "A=" + cal.ToString("f"), layer1, 0, AttachmentPoint.BottomCenter))

                                            draw.addAtLastEntity(draw.drawLine(lastNotEulesDis, alpha2, lastNotEulesDis + accEulerDis, 0, layer2))


                                        ElseIf Not lastNotEuler.Item2.Equals("Arc") And nextTuple.Item2.Equals("Arc") Then

                                            Dim alpha2 As Double = 0

                                            If pline.GetBulgeAt(nextTuple.Item1) > 0 Then
                                                alpha2 = mm / nextTuple.Item6
                                            Else
                                                alpha2 = -mm / nextTuple.Item6
                                            End If

                                            Dim cal As Double = Math.Sqrt(accEulerDis * nextTuple.Item6)

                                            draw.addAtFirst(draw.drawText(lastNotEulesDis + accEulerDis / 2, 1, "A=" + cal.ToString("f"), layer1, 0, AttachmentPoint.BottomCenter))

                                            draw.addAtLastEntity(draw.drawLine(lastNotEulesDis, 0, lastNotEulesDis + accEulerDis, alpha2, layer2))


                                        ElseIf lastNotEuler.Item2.Equals("Arc") And nextTuple.Item2.Equals("Arc") Then
                                            MsgBox("Se encontro clotoide, pero hay dos Arcos posibles, porfavor revisar a mano" & vbCrLf)
                                        Else
                                            MsgBox("Se encontro clotoide, pero no se encontro una curva a cual asociarla" & vbCrLf)
                                        End If

                                    ElseIf analisis.Count <= tupla.Item1 + 1 Then
                                        MsgBox("Se encontro clotoide, pero no se encontro una curva a cual asociarla" & vbCrLf)
                                    End If

                                Case Else

                                    MsgBox("Error en el dibujo" & vbCrLf)
                                    Return

                            End Select

                            If height <> lastHeight And Not last.Item2.Equals("None") And Not last.Item2.Equals("Euler") Then
                                draw.addAtLastEntity(draw.drawLine(dis, lastHeight, dis, height, layer2))
                            End If


                            dis += tupla.Item4

                            last = tupla

                            lastHeight = height

                        Next

                        draw.addAtLastEntity(draw.drawLine(0, 0, dis, 0, layer3))

                        draw.draw()

                        mytrans.Commit()

                    Else
                        mytrans.Commit()
                    End If

                Catch ex As Exception
                    mytrans.Commit()
                    MsgBox("Error en el dibujo" & vbCrLf & ex.Message)
                    Return
                End Try

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

    <CommandMethod("SET_PARAMETROS")> _
    Public Sub cambiarProporciones()
        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ingrese nueva Escala Horizontal (Default 1) 1:")

        pStrOpts.DefaultValue = 1
        pStrOpts.AllowNegative = False
        pStrOpts.AllowZero = False

        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)

        H = pStrRes.Value

        pStrOpts = New PromptDoubleOptions(vbLf & "Ingrese tamao del Texto (Default 1) : ")

        pStrOpts.DefaultValue = 1
        pStrOpts.AllowNegative = False
        pStrOpts.AllowZero = False

        pStrRes = acDoc.Editor.GetDouble(pStrOpts)

        heightText = pStrRes.Value

        pStrOpts = New PromptDoubleOptions(vbLf & "Ingrese nueva Escala de Curvatura (Default 750):")

        pStrOpts.DefaultValue = 750
        pStrOpts.AllowNegative = False
        pStrOpts.AllowZero = False

        pStrRes = acDoc.Editor.GetDouble(pStrOpts)

        mm = pStrRes.Value

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

                If (arcseg.Radius * arcseg.EndAngle) < tolerancia Then

                    If lastType.Equals("None") Or lastType.Equals("Rect") Or lastType.Equals("Arc") Then
                        countEulerSegments += 1
                    ElseIf lastType.Equals("Euler") Then

                    Else
                        MsgBox("Error identificando polylinea")
                        Throw New System.Exception("An exception has occurred.")
                    End If

                    parsePropperties.Add(New Tuple(Of Integer, String, Integer, Double, Double, Double)(vertx, "Euler", countEulerSegments, arcseg.Radius * arcseg.EndAngle, 0, 0))
                    lastType = "Euler"

                Else
                    countArcSegments += 1
                    parsePropperties.Add(New Tuple(Of Integer, String, Integer, Double, Double, Double)(vertx, "Arc", countArcSegments, arcseg.EndAngle * arcseg.Radius, arcseg.EndAngle, arcseg.Radius))
                    lastType = "Arc"

                End If

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

    Private Function getPolyline() As PromptEntityResult

        Dim myPEO As New Autodesk.AutoCAD.EditorInput.PromptEntityOptions(vbLf & "Seleccione Eje en planta:")
        Dim mydwg As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim myed As Editor = mydwg.Editor

        myPEO.SetRejectMessage("Porfavor, seleciona una Polylinea." & vbCrLf)
        myPEO.AddAllowedClass(GetType(Autodesk.AutoCAD.DatabaseServices.Polyline), False)
        Dim myPer As PromptEntityResult = myed.GetEntity(myPEO)

        Return myPer
    End Function

    Private Function angle3Points(startPoint As Point3d, commonPoint As Point3d, endPoint As Point3d) As Double

        Dim alpha As Double

        Dim A As Double = startPoint.Y - commonPoint.Y
        Dim B As Double = startPoint.X - commonPoint.X

        Dim norm As Double = (Math.Sqrt(Math.Pow(A, 2) + Math.Pow(B, 2)))

        A = A / norm
        B = B / norm

        Dim A2 As Double = commonPoint.Y - endPoint.Y
        Dim B2 As Double = commonPoint.X - endPoint.X

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


        Return alpha * 200 / Math.PI

    End Function

End Class