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
    Dim vScale As Double = 1

    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim draw As DrawClass

    Dim LayerTerreno As String = "PL-Terreno"
    Dim LayerPRC As String = "PL-PRC"
    Dim LayerTierra As String = "PL-Tierra"
    Dim LayerAcera As String = "PL-Acera"
    Dim LayerCalzada As String = "PL-Calzada"
    Dim LayerCiclovia As String = "PL-Ciclovia"
    Dim LayerAverde As String = "PL-AVerde"
    Dim LayerFFCC As String = "PL-FFCC"
    Dim LayerTextoTitulo As String = "TextoTitulo"
    Dim LayerTextoComentario As String = "TextoComentario"
    Dim LayerDimension As String = "Dimension"
    Dim LayerHatsh As String = "$Fondo"
    Dim LayerPostes As String = "POSTES"



    <CommandMethod("ImportPTFromExcel")> _
    Public Sub importGraphFromExcel()

        Dim excelFile As String = Nothing
        Dim OpenFileDialog1 As OpenFileDialog = importExcel()

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then

            For Each file In OpenFileDialog1.FileNames
                Try

                    xlApp = New Microsoft.Office.Interop.Excel.Application
                    xlBook = xlApp.Workbooks.Open(file)
                    xlSheet = xlBook.Worksheets(1)

                    Dim pPtRes As PromptPointResult = importPoint()

                    If pPtRes.Status = PromptStatus.OK Then

                        draw = New DrawClass(pPtRes.Value.X, pPtRes.Value.Y, hScale, vScale)

                        setLayout()

                        drawPT(1, 1, xlSheet, 0, 0)

                        draw.draw()

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

        draw.addLayer(LayerTerreno, Color.FromRgb(0, 0, 0), 1, 0.05)
        draw.addLayer(LayerPRC, Color.FromRgb(255, 0, 0), 0.3, 0.085)
        draw.addLayer(LayerTierra, Color.FromRgb(165, 124, 0), 1, 0.216)
        draw.addLayer(LayerAcera, Color.FromRgb(255, 191, 0), 1, 0.216)
        draw.addLayer(LayerCalzada, Color.FromRgb(128, 128, 128), 1, 0.216)
        draw.addLayer(LayerCiclovia, Color.FromRgb(127, 63, 63), 1, 0.216)
        draw.addLayer(LayerAverde, Color.FromRgb(0, 255, 0), 1, 0.216)
        draw.addLayer(LayerFFCC, Color.FromRgb(0, 255, 255), 1, 0.216)

        draw.addLayer(LayerTextoTitulo, Color.FromRgb(0, 0, 0), 1, 0)
        draw.addLayer(LayerTextoComentario, Color.FromRgb(0, 0, 0), 0.2, 0)
        draw.addLayer(LayerDimension, Color.FromRgb(0, 0, 0), 0.5, 0)

        draw.addLayer(LayerHatsh, Color.FromRgb(255, 255, 255), 1, 0)

        draw.addLayer(LayerPostes, Color.FromRgb(0, 0, 0), 1, 0)



    End Sub

    Private Function importExcel() As OpenFileDialog

        Dim OpenFileDialog1 As New OpenFileDialog()

        OpenFileDialog1.Title = "Please select a excel file to open" 'User prompt 'for file open
        OpenFileDialog1.Filter = "Excel Files (*.xlsx)|*.xlsx| Excel files (*.xls)|*.xls"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.RestoreDirectory = True

        Return OpenFileDialog1

    End Function

    Private Sub drawPT(ci As Integer, cj As Integer, xlSheet As Worksheet, x As Double, y As Double)

        Dim j As Integer = cj + 4
        Dim i As Integer = ci + 1

        Dim x0 As Double = 0
        Dim y0 As Double = 0

        Dim bool As Boolean = False
        Dim fail As Boolean = False

        Dim length As Double = getLength(ci, cj, xlSheet)

        draw.drawHatch(x - length / 2 - 5, y - 14.6296, 0, 0, length + 10, 14.6296, LayerHatsh, LayerPostes)


        While Not String.IsNullOrEmpty(xlSheet.Cells(j, i).Value)

            If fail Then
                MsgBox("valor de celda no esperado X:" + i.ToString + ", Y:" + j.ToString)
            End If

            Select Case xlSheet.Cells(j, i).Value

                Case "F"
                    If Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        x0 = xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3
                        y0 = 6.5
                        bool = True

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        y0 = 6.5
                        bool = True

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        fail = True

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        fail = True

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 0.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    End If
                Case "FB"
                    If Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        x0 = xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3
                        y0 = 6.5
                        bool = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        y0 = 6.5
                        bool = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        fail = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        fail = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 1, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    End If
                Case "FS"
                    If Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        x0 = xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3
                        y0 = 6.5
                        bool = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        y0 = 6.5
                        bool = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        fail = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        fail = True

                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerTerreno))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    End If
                Case "FPRC"
                    If Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        x0 = xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3
                        y0 = 6.5
                        bool = True

                        draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, x0 - 0.1, y0 + 3.0, "PRC", LayerPRC, Math.PI / 2, AttachmentPoint.BottomLeft))
                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerPRC))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And Not bool Then

                        y0 = 6.5
                        bool = True

                        draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, x0 - 0.1, y0 + 3.0, "PRC", LayerPRC, Math.PI / 2, AttachmentPoint.BottomLeft))
                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerPRC))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        drawComentaryFinal(length, x, y, j, i)

                        fail = True

                        draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, x0 - 0.1, y0 + 3.0, "PRC", LayerPRC, Math.PI / 2, AttachmentPoint.BottomLeft))
                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerPRC))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    ElseIf String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) And bool Then

                        fail = True

                        draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, x0 - 0.1, y0 + 3.0, "PRC", LayerPRC, Math.PI / 2, AttachmentPoint.BottomLeft))
                        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, x0, y0, x0, y0 + 3, LayerPRC))

                        If Not String.IsNullOrEmpty(xlSheet.Cells(j - 2, i).Value) Then
                            draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0, y0 + 1.2, xlSheet.Cells(j - 2, i).Value.ToString, LayerTextoComentario, Math.PI / 2, AttachmentPoint.MiddleLeft))
                        End If

                    End If

                Case "T"
                Case "A"
                Case "C"
                Case "Ci"
                Case "Av"
                Case "FFCC"
                Case "L"
                Case "Com"

            End Select

            i += 1
        End While

    End Sub

    Private Sub drawComentaryFinal(length As Double, x As Double, y As Double, j As Integer, i As Integer)

        draw.addAtLastEntity(draw.drawLine(x - length / 2, y - 14.6296, 0, 6.5, xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3, 6.5, LayerPostes))
        draw.addAtFirst(draw.drawText(x - length / 2, y - 14.6296, 0.1, 6.7, xlSheet.Cells(j - 4, i).Value.ToString, LayerTextoComentario, 0, AttachmentPoint.BottomLeft))

    End Sub

    Private Function getLength(ci As Integer, cj As Integer, xlSheet As Worksheet) As Double

        Dim j As Integer = cj + 4
        Dim i As Integer = ci + 1
        Dim length As Double = 0

        While Not String.IsNullOrEmpty(xlSheet.Cells(j, i).Value)

            Dim val As String = xlSheet.Cells(j, i).Value
            If val.Equals("F") Or val.Equals("FB") Or val.Equals("FB") Or val.Equals("FB") Then
                If Not String.IsNullOrEmpty(xlSheet.Cells(j - 4, i).Value) Then

                    length += xlSheet.Cells(j - 4, i).Value.ToString.Length * 0.2 + 0.3

                End If
            End If

            If Not String.IsNullOrEmpty(xlSheet.Cells(j - 1, i).Value) Then

                length += xlSheet.Cells(j - 1, i).Value

            End If

            i += 1
        End While

        Return length

    End Function

    Private Function getDiameter(celda As String) As Double
        Dim values As String() = celda.Split("-")
        Return Double.Parse(values.GetValue(1))
    End Function

    Private Function importPoint() As PromptPointResult

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

        pPtOpts.Message = vbLf & "Enter the start point of the Draw: "
        Return acDoc.Editor.GetPoint(pPtOpts)

    End Function

End Class