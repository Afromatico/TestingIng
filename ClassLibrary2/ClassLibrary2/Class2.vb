Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Autodesk.AutoCAD.Colors
Imports System.Globalization



Public Class Class1

    Dim hScale As Double = 1
    Dim vScale As Double = 10

    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim displayPoint As Point3d

    Dim LayerTuberia As String = "Tuberia"
    Dim LayerTerreno As String = "Terreno"
    Dim LayerDibujo As String = "Dibujo"
    Dim LayerCajon As String = "Cajon"
    Dim LayerDistancias As String = "Distancias"
    Dim LayerCotas As String = "Cotas"
    Dim LayerDatos As String = "Datos"

    Sub New()
        CreateAndAssignALayer("Tuberia", Color.FromRgb(0, 0, 255), LineWeight.LineWeight050)
        CreateAndAssignALayer("Terreno", Color.FromRgb(0, 0, 0), LineWeight.ByLineWeightDefault)
        CreateAndAssignALayer("Dibujo", Color.FromRgb(0, 0, 0), LineWeight.ByLineWeightDefault)
        CreateAndAssignALayer("Cajon", Color.FromRgb(255, 0, 0), LineWeight.LineWeight050)

        CreateAndAssignALayer("Distancias", Color.FromRgb(0, 0, 0), LineWeight.ByLineWeightDefault)
        CreateAndAssignALayer("Cotas", Color.FromRgb(0, 0, 0), LineWeight.ByLineWeightDefault)
        CreateAndAssignALayer("Datos", Color.FromRgb(0, 0, 0), LineWeight.ByLineWeightDefault)
    End Sub

    <CommandMethod("ImportGraphFromExcel")> _
    Public Sub importGraphFromExcel()

        Dim excelFile As String = Nothing
        Dim OpenFileDialog1 As OpenFileDialog = importExcel()
        Dim x As Double = 0

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            For Each file In OpenFileDialog1.FileNames
                Try

                    xlApp = New Microsoft.Office.Interop.Excel.Application
                    xlBook = xlApp.Workbooks.Open(file)
                    xlSheet = xlBook.Worksheets(1)

                    Dim pPtRes As PromptPointResult = importPoint()

                    If pPtRes.Status = PromptStatus.OK Then

                        displayPoint = pPtRes.Value

                        Dim referencia As Double = checkReference(xlSheet)

                        drawInitialState(xlSheet, referencia)

                        x = drawLines(xlSheet, referencia)

                        drawFinalState(xlSheet, referencia, x)

                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Se ha ejecutado con exito el programa, se a dibujado un perfil de " + x.ToString + " Metros.")

                    Else
                        Throw New Exception(ErrorStatus.InvalidDxf3dPoint, "Error al recivir un punto")
                    End If

                Catch ex As Exception
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message)
                    Return
                End Try
            Next
        End If
    End Sub

    Private Function importExcel() As OpenFileDialog

        Dim OpenFileDialog1 As New OpenFileDialog()

        OpenFileDialog1.Title = "Please select a excel file to open" 'User prompt 'for file open
        OpenFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx| Excel files (*.xls)|*.xls"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.RestoreDirectory = True

        Return OpenFileDialog1

    End Function

    Private Function checkReference(xlSheet As Worksheet) As Double

        Dim vertical As Integer = 5
        Dim horizontal As Integer = 3
        Dim referencia As Double = 0
        referencia = Math.Round(xlSheet.Cells(vertical, horizontal).Value() - 0.5) - 2

        While Not String.IsNullOrEmpty(xlSheet.Cells(vertical, horizontal).Value)
            If Math.Round(xlSheet.Cells(vertical, horizontal).Value() - 0.5) = 0 Then
                referencia = 0
                Exit While
            End If

            If xlSheet.Cells(vertical, horizontal).Value() - referencia - 2 < 0 Then
                referencia = Math.Round(xlSheet.Cells(vertical, horizontal).Value() - 0.5) - 2
                If referencia < 0 Then
                    referencia = 0
                    Exit While
                End If
            End If

            horizontal += 2

        End While

        Return referencia

    End Function

    Private Sub drawInitialState(xlSheet As Worksheet, reference As Double)

        drawLine(0, 0, 0, 60, LayerCajon, 0)
        drawLine(45, 0, 45, 60, LayerCajon, 0)

        drawText(28, 61, "REF. " + reference.ToString, LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 56, "DISTANCIAS PARCIALES", LayerDistancias, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 51, "DISTANC. ACUMULADAS", LayerDistancias, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 46, "COTAS TERRENO", LayerCotas, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 41, "COTAS RADIER", LayerCotas, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 31, "MATERIAL-DIAMETROS", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 26, "CAUDAL (l/s)", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 21, "PENDIENTES", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 16, "VOLUMEN EXCAV. 0-2 m", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(14.3, 11, "2-4 m", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(14.3, 6, "4-6 m", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)
        drawText(1.8, 1, "APOYO TIPO", LayerDatos, 2.5, 0, AttachmentPoint.BottomLeft)

        drawText(50, 51, xlSheet.Cells(2, 3).Value().ToString, LayerDistancias, 2.5, 0, AttachmentPoint.BottomCenter)
        drawText(50, 46, System.Convert.ToDouble(xlSheet.Cells(3, 3).Value()).ToString("F"), LayerCotas, 2.5, 0, AttachmentPoint.BottomCenter)
        drawText(50, 36, System.Convert.ToDouble(xlSheet.Cells(5, 3).Value()).ToString("F"), LayerCotas, 2.5, 0, AttachmentPoint.BottomCenter)

        drawText(49, 55 + (xlSheet.Cells(3, 3).Value() - reference) * vScale, System.Convert.ToDouble(Math.Round(xlSheet.Cells(3, 3).Value() - xlSheet.Cells(5, 3).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, 1.85, Math.PI / 2, AttachmentPoint.BottomCenter)

        drawLine(50, 60, 50, 60 + (xlSheet.Cells(3, 3).Value() - reference) * vScale, LayerDibujo, 0)
        drawLine(50, 60 + (xlSheet.Cells(5, 3).Value() - reference) * vScale, 50, 60 + (xlSheet.Cells(3, 3).Value() - reference) * vScale, LayerTuberia, 0.5)


    End Sub

    Private Sub drawLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, layer As String, widht As Double)

        Dim db As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim mSpace As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Dim newLine As New Autodesk.AutoCAD.DatabaseServices.Polyline
            newLine.AddVertexAt(0, New Point2d(x0 + displayPoint.X, y0 + displayPoint.Y), 0, widht, widht)
            newLine.AddVertexAt(0, New Point2d(x1 + displayPoint.X, y1 + displayPoint.Y), 0, widht, widht)
            newLine.Layer = layer
            mSpace.AppendEntity(newLine)
            trans.AddNewlyCreatedDBObject(newLine, True)
            trans.Commit()

        End Using

    End Sub

    Private Sub drawText(x0 As Double, y0 As Double, text As String, layer As String, height As Double, rotation As Double, justify As AttachmentPoint)

        Dim db As Database = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim mSpace As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)
            Using asMtext As MText = New MText()
                asMtext.Attachment = justify
                asMtext.SetAttachmentMovingLocation(asMtext.Attachment)
                asMtext.Location = New Point3d(x0 + displayPoint.X, y0 + displayPoint.Y, 0)
                asMtext.Width = 55
                asMtext.Contents = text
                asMtext.Rotation = rotation
                asMtext.Layer = layer
                asMtext.TextHeight = height
                mSpace.AppendEntity(asMtext)
                trans.AddNewlyCreatedDBObject(asMtext, True)
            End Using
            trans.Commit()

        End Using

    End Sub

    Private Function drawLines(xlSheet As Worksheet, reference As Double) As Double

        Dim x As Double = 0
        Dim y As Double = 0
        Dim h As Double = 0
        Dim j As Integer = 3
        Dim i As Integer = 2

        While Not String.IsNullOrEmpty(xlSheet.Cells(i, j).Value)

            If xlSheet.Cells(i, j).Value() = 0 Then
                y = xlSheet.Cells(i + 1, j).Value() - reference
                h = xlSheet.Cells(i + 3, j).Value() - reference
                j += 2
                Continue While
            End If

            Dim black As Color = Color.FromRgb(0, 0, 0)
            Dim blue As Color = Color.FromRgb(0, 0, 255)

            Dim diametro As Double = getDiameter(xlSheet.Cells(i + 5, j - 1).Value().ToString) / 1000


            drawLine(x * hScale + 50, y * vScale + 60, xlSheet.Cells(i, j).Value() * hScale + 50, (xlSheet.Cells(i + 1, j).Value() - reference) * vScale + 60, LayerTerreno, 0)
            drawLine(xlSheet.Cells(i, j).Value() * hScale + 50, 60, xlSheet.Cells(i, j).Value() * hScale + 50, (xlSheet.Cells(i + 1, j).Value() - reference) * vScale + 60, LayerDibujo, 0)
            drawLine(xlSheet.Cells(i, j).Value() * hScale + 50, 0, xlSheet.Cells(i, j).Value() * hScale + 50, 35, LayerDibujo, 0)

            drawLine(x * hScale + 50, h * vScale + 60, xlSheet.Cells(i, j).Value() * hScale + 50, (xlSheet.Cells(i + 2, j).Value() - reference) * vScale + 60, LayerTuberia, 0.5)
            drawLine(x * hScale + 50, (h + diametro) * vScale + 60, xlSheet.Cells(i, j).Value() * hScale + 50, ((xlSheet.Cells(i + 2, j).Value() - reference) + diametro) * vScale + 60, LayerTuberia, 0.5)

            drawLine(xlSheet.Cells(i, j).Value() * hScale + 50, (xlSheet.Cells(i + 3, j).Value() - reference) * vScale + 60, xlSheet.Cells(i, j).Value() * hScale + 50, (xlSheet.Cells(i + 1, j).Value() - reference) * vScale + 60, LayerTuberia, 0.5)

            drawText(xlSheet.Cells(i, j).Value() * hScale + 49, 55 + (xlSheet.Cells(i + 1, j).Value() - reference) * vScale, System.Convert.ToDouble(Math.Round(xlSheet.Cells(i + 1, j).Value() - xlSheet.Cells(i + 2, j).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, 1.85, Math.PI / 2, AttachmentPoint.BottomCenter)
            drawText(xlSheet.Cells(i, j).Value() * hScale + 51, 55 + (xlSheet.Cells(i + 1, j).Value() - reference) * vScale, System.Convert.ToDouble(Math.Round(xlSheet.Cells(i + 1, j).Value() - xlSheet.Cells(i + 3, j).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, 1.85, Math.PI / 2, AttachmentPoint.TopCenter)

            Dim actualX As Double = xlSheet.Cells(i, j).Value() * hScale + 50

            drawText(actualX, 51, xlSheet.Cells(i, j).Value().ToString, LayerDistancias, 2.5, 0, AttachmentPoint.BottomCenter)
            drawText(actualX, 46, System.Convert.ToDouble(xlSheet.Cells(i + 1, j).Value()).ToString("F"), LayerCotas, 2.5, 0, AttachmentPoint.BottomCenter)
            drawText(actualX, 41, System.Convert.ToDouble(xlSheet.Cells(i + 2, j).Value()).ToString("F"), LayerCotas, 2.5, 0, AttachmentPoint.BottomCenter)
            drawText(actualX, 36, System.Convert.ToDouble(xlSheet.Cells(i + 3, j).Value()).ToString("F"), LayerCotas, 2.5, 0, AttachmentPoint.BottomCenter)

            actualX = (xlSheet.Cells(i, j).Value() + x) * hScale / 2 + 50

            drawText(actualX, 56, xlSheet.Cells(i - 1, j - 1).Value().ToString, LayerDistancias, 2.5, 0, AttachmentPoint.BottomCenter)
            drawText(actualX, 31, xlSheet.Cells(i + 5, j - 1).Value().ToString, LayerDatos, 2.5, 0, AttachmentPoint.BottomCenter)
            drawText(actualX, 21, System.Convert.ToDouble(xlSheet.Cells(i + 7, j - 1).Value() * 100).ToString("F") + "%", LayerDatos, 2.5, 0, AttachmentPoint.BottomCenter)


            x = xlSheet.Cells(i, j).Value()
            y = xlSheet.Cells(i + 1, j).Value() - reference
            h = xlSheet.Cells(i + 3, j).Value() - reference

            j += 2

        End While

        Return x

    End Function

    <CommandMethod("SETHORIZONTALSCALE")> _
    Public Sub setHorizontalScale()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ingrese nueva Escala Horizontal: ")
        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)

        hScale = 100 / pStrRes.Value

        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("se ha cambiado escala horizontal a: " + hScale.ToString)
    End Sub

    <CommandMethod("SETVERTICALSCALE")> _
    Public Sub setVerticalScale()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ingrese nueva Escala Horizontal: ")
        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)

        vScale = 100 / pStrRes.Value

        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("se ha cambiado escala vertical a" + vScale.ToString)
    End Sub

    Private Function getDiameter(celda As String) As Double
        Dim values As String() = celda.Split("-")
        Return Double.Parse(values.GetValue(1))
    End Function

    Private Sub drawFinalState(xlSheet As Worksheet, referencia As Double, x As Double)


        drawLine(0, 0, x * hScale + 63, 0, LayerCajon, 0)
        drawLine(0, 60, x * hScale + 63, 60, LayerCajon, 0)
        drawLine(x * hScale + 63, 0, x * hScale + 63, 60, LayerCajon, 0)

        For index As Integer = 1 To 11
            drawLine(0, 60 - 5 * index, x * hScale + 63, 60 - 5 * index, LayerDibujo, 0)
        Next

    End Sub

    Private Function importPoint() As PromptPointResult

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        pPtOpts.Message = vbLf & "Enter the start point of the Draw: "
        Return acDoc.Editor.GetPoint(pPtOpts)

    End Function

    Public Sub CreateAndAssignALayer(layer1 As String, color As Color, linewidth As LineWeight)
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
                    acLyrTblRec.LineWeight = linewidth

                    acLyrTbl.UpgradeOpen()

                    acLyrTbl.Add(acLyrTblRec)
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, True)
                End Using
            End If

            '' Save the changes and dispose of the transaction
            acTrans.Commit()
        End Using
    End Sub
End Class
