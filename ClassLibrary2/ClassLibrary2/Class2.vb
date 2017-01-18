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
Imports AutocadManager2015



Public Class Class1

    Dim hScale As Double = 1
    Dim vScale As Double = 10

    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim draw As DrawClass

    Dim LayerTuberia As String = "Tuberia"
    Dim LayerTerreno As String = "Terreno"
    Dim LayerDibujo As String = "Dibujo"
    Dim LayerCajon As String = "Cajon"
    Dim LayerDistancias As String = "Distancias"
    Dim LayerCotas As String = "Cotas"
    Dim LayerDatos As String = "Datos"
    Dim LayerDatosCotas As String = "DatosCotas"

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

                        draw = New DrawClass(pPtRes.Value.X, pPtRes.Value.Y, hScale, vScale)

                        Dim referencia As Double = checkReference(xlSheet)

                        setLayout()

                        drawInitialState(xlSheet, referencia)

                        x = drawLines(xlSheet, referencia)

                        drawFinalState(xlSheet, referencia, x)

                        draw.draw()

                        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("Se ha ejecutado con exito el programa, se a dibujado un perfil de " + x.ToString + " Metros.")

                    Else
                        Throw New Exception(ErrorStatus.InvalidDxf3dPoint, "Error al recivir un punto")
                    End If

                    xlBook.Close()
                Catch ex As Exception
                    Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog(ex.Message)
                    xlBook.Close()
                    Return
                End Try
            Next
        End If
    End Sub

    Private Sub setLayout()

        draw.addLayer(LayerTuberia, Color.FromRgb(0, 0, 255), 1, 0.5)
        draw.addLayer(LayerTerreno, Color.FromRgb(0, 0, 0), 1, 0)
        draw.addLayer(LayerDibujo, Color.FromRgb(0, 0, 0), 1, 0)
        draw.addLayer(LayerCajon, Color.FromRgb(255, 0, 0), 1, 0.5)

        draw.addLayer(LayerDistancias, Color.FromRgb(0, 0, 0), 2.5, 0)
        draw.addLayer(LayerCotas, Color.FromRgb(0, 0, 0), 2.5, 0)
        draw.addLayer(LayerDatos, Color.FromRgb(0, 0, 0), 2.5, 0)
        draw.addLayer(LayerDatosCotas, Color.FromRgb(0, 0, 0), 1.85, 0)

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

        draw.addAtLastEntity(draw.drawLine(0, 0, 0, 60 / vScale, LayerCajon))
        draw.addAtLastEntity(draw.drawLine(45, 0, 0, 0, 0, 60 / vScale, LayerCajon))

        draw.addAtFirst(draw.drawText(28, 61, 0, 0, "REF. " + reference.ToString, LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 56, 0, 0, "DISTANCIAS PARCIALES", LayerDistancias, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 51, 0, 0, "DISTANC. ACUMULADAS", LayerDistancias, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 46, 0, 0, "COTAS TERRENO", LayerCotas, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 41, 0, 0, "COTAS RADIER", LayerCotas, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 31, 0, 0, "MATERIAL-DIAMETROS", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 26, 0, 0, "CAUDAL (l/s)", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 21, 0, 0, "PENDIENTES", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 16, 0, 0, "VOLUMEN EXCAV. 0-2 m", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(14.3, 11, 0, 0, "2-4 m", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(14.3, 6, 0, 0, "4-6 m", LayerDatos, 0, AttachmentPoint.BottomLeft))
        draw.addAtFirst(draw.drawText(1.8, 1, 0, 0, "APOYO TIPO", LayerDatos, 0, AttachmentPoint.BottomLeft))

        draw.addAtFirst(draw.drawText(50, 51, 0, 0, xlSheet.Cells(2, 3).Value().ToString, LayerDistancias, 0, AttachmentPoint.BottomCenter))
        draw.addAtFirst(draw.drawText(50, 46, 0, 0, System.Convert.ToDouble(xlSheet.Cells(3, 3).Value()).ToString("F"), LayerCotas, 0, AttachmentPoint.BottomCenter))
        draw.addAtFirst(draw.drawText(50, 36, 0, 0, System.Convert.ToDouble(xlSheet.Cells(5, 3).Value()).ToString("F"), LayerCotas, 0, AttachmentPoint.BottomCenter))

        draw.addAtFirst(draw.drawText(49, 55, 0, (xlSheet.Cells(3, 3).Value() - reference), System.Convert.ToDouble(Math.Round(xlSheet.Cells(3, 3).Value() - xlSheet.Cells(5, 3).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, Math.PI / 2, AttachmentPoint.BottomCenter))

        draw.addAtLastEntity(draw.drawLine(50, 60, 0, 0, 0, (xlSheet.Cells(3, 3).Value() - reference), LayerDibujo))
        draw.addAtLastEntity(draw.drawLine(50, 60, 0, (xlSheet.Cells(5, 3).Value() - reference), 0, (xlSheet.Cells(3, 3).Value() - reference), LayerTuberia))

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

            Dim diametro As Double = getDiameter(xlSheet.Cells(i + 5, j - 1).Value().ToString) / 1000


            draw.addAtLastEntity(draw.drawLine(50, 60, x, y, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 1, j).Value() - reference), LayerTerreno))
            draw.addAtLastEntity(draw.drawLine(50, 60, xlSheet.Cells(i, j).Value(), 0, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 1, j).Value() - reference), LayerDibujo))
            draw.addAtLastEntity(draw.drawLine(50, 0, xlSheet.Cells(i, j).Value(), 0, xlSheet.Cells(i, j).Value(), 35 / vScale, LayerDibujo))

            draw.addAtLastEntity(draw.drawLine(50, 60, x, h, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 2, j).Value() - reference), LayerTuberia))
            draw.addAtLastEntity(draw.drawLine(50, 60, x, (h + diametro), xlSheet.Cells(i, j).Value(), ((xlSheet.Cells(i + 2, j).Value() - reference) + diametro), LayerTuberia))

            draw.addAtLastEntity(draw.drawLine(50, 60, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 3, j).Value() - reference), xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 1, j).Value() - reference), LayerTuberia))

            draw.addAtFirst(draw.drawText(49, 55, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 1, j).Value() - reference), System.Convert.ToDouble(Math.Round(xlSheet.Cells(i + 1, j).Value() - xlSheet.Cells(i + 2, j).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, Math.PI / 2, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(51, 55, xlSheet.Cells(i, j).Value(), (xlSheet.Cells(i + 1, j).Value() - reference), System.Convert.ToDouble(Math.Round(xlSheet.Cells(i + 1, j).Value() - xlSheet.Cells(i + 3, j).Value(), 2, MidpointRounding.AwayFromZero)).ToString("F"), LayerDatos, Math.PI / 2, AttachmentPoint.TopCenter))

            Dim actualX As Double = xlSheet.Cells(i, j).Value()

            draw.addAtFirst(draw.drawText(50, 51, actualX, 0, xlSheet.Cells(i, j).Value().ToString, LayerDistancias, 0, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(50, 46, actualX, 0, System.Convert.ToDouble(xlSheet.Cells(i + 1, j).Value()).ToString("F"), LayerCotas, 0, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(50, 41, actualX, 0, System.Convert.ToDouble(xlSheet.Cells(i + 2, j).Value()).ToString("F"), LayerCotas, 0, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(50, 36, actualX, 0, System.Convert.ToDouble(xlSheet.Cells(i + 3, j).Value()).ToString("F"), LayerCotas, 0, AttachmentPoint.BottomCenter))

            actualX = (xlSheet.Cells(i, j).Value() + x) / 2

            draw.addAtFirst(draw.drawText(50, 56, actualX, 0, xlSheet.Cells(i - 1, j - 1).Value().ToString, LayerDistancias, 0, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(50, 31, actualX, 0, xlSheet.Cells(i + 5, j - 1).Value().ToString, LayerDatos, 0, AttachmentPoint.BottomCenter))
            draw.addAtFirst(draw.drawText(50, 21, actualX, 0, System.Convert.ToDouble(xlSheet.Cells(i + 7, j - 1).Value() * 100).ToString("F") + "%", LayerDatos, 0, AttachmentPoint.BottomCenter))


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

        hScale = pStrRes.Value

        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("se ha cambiado escala horizontal a: " + hScale.ToString)
    End Sub

    <CommandMethod("SETVERTICALSCALE")> _
    Public Sub setVerticalScale()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim pStrOpts As PromptDoubleOptions = New PromptDoubleOptions(vbLf & "Ingrese nueva Escala Horizontal: ")
        Dim pStrRes As PromptDoubleResult = acDoc.Editor.GetDouble(pStrOpts)

        vScale = pStrRes.Value

        Autodesk.AutoCAD.ApplicationServices.Application.ShowAlertDialog("se ha cambiado escala vertical a" + vScale.ToString)
    End Sub

    Private Function getDiameter(celda As String) As Double
        Dim values As String() = celda.Split("-")
        Return Double.Parse(values.GetValue(1))
    End Function

    Private Sub drawFinalState(xlSheet As Worksheet, referencia As Double, x As Double)

        draw.addAtLastEntity(draw.drawLine(0, 0, x + 63 / hScale, 0, LayerCajon))
        draw.addAtLastEntity(draw.drawLine(0, 60, 0, 0, x + 63 / hScale, 0, LayerCajon))
        draw.addAtLastEntity(draw.drawLine(63, 0, x, 0, x, 60 / vScale, LayerCajon))

        For index As Integer = 1 To 11
            draw.addAtLastEntity(draw.drawLine(0, 60 - 5 * index, 0, 0, x + 63 / hScale, 0, LayerDibujo))
        Next

    End Sub

    Private Function importPoint() As PromptPointResult

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        pPtOpts.Message = vbLf & "Enter the start point of the Draw: "
        Return acDoc.Editor.GetPoint(pPtOpts)

    End Function

End Class
