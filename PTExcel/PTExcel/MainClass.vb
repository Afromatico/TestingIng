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



Public Class MainClass

    Dim hScale As Double = 1
    Dim vScale As Double = 1

    Public com As Integer = 1

    Dim xlApp As Microsoft.Office.Interop.Excel.Application
    Dim xlBook As Microsoft.Office.Interop.Excel.Workbook
    Dim xlSheet As Microsoft.Office.Interop.Excel.Worksheet

    Dim draw As DrawClass

    Dim blockList As List(Of Block)

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

                        blockList = New List(Of Block)

                        createBlocks(1, 1, xlSheet, -11.4728)

                        setLayout()

                        drawPT(xlSheet, 0, 0)

                        Dim title As MText = draw.drawText(0, -1.3, xlSheet.Cells(1, 1).Value, LayerTextoTitulo, 0, AttachmentPoint.BottomCenter)
                        draw.addAtFirst(title)

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

    Private Sub createBlocks(ci As Integer, cj As Integer, xlSheet As Worksheet, h As Double)

        Dim j As Integer = cj + 4
        Dim i As Integer = ci + 1
        Dim first As Boolean = False

        While Not String.IsNullOrEmpty(xlSheet.Cells(j, i).Value)

            Select Case xlSheet.Cells(j, i).Value

                Case "F"
                    blockList.Add(New F(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FB"
                    blockList.Add(New FB(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FS"
                    blockList.Add(New FS(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FPRC"
                    blockList.Add(New FPRC(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "T"
                    blockList.Add(New T(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "A"
                    blockList.Add(New A(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "C"
                    blockList.Add(New C(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Ci"
                    blockList.Add(New Ci(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Av"
                    blockList.Add(New Av(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "FFCC"
                    blockList.Add(New FFCC(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "L"
                    blockList.Add(New L(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Com"
                    blockList.Add(New Com(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h, Me, ci, cj, xlSheet))
            End Select

            first = True
            i += 1
            com = i
        End While

    End Sub

    Private Sub drawPT(xlSheet As Worksheet, x As Double, y As Double)

        Dim tupla As Tuple(Of Double, Double)
        Dim length As Double = getLength()

        If length < 30 Then
            draw.drawHatch(x - 20, y - 14.6296, 0, 0, 40, 14.6296, LayerHatsh, LayerPostes)
        Else
            draw.drawHatch(x - length / 2 - 5, y - 14.6296, 0, 0, length + 10, 14.6296, LayerHatsh, LayerPostes)
        End If

        tupla = New Tuple(Of Double, Double)(x - length / 2, y - 8.1296)

        For Each Block In blockList
            tupla = Block.draw(tupla)
        Next

    End Sub

    Private Function getLength() As Double

        Dim length As Double = 0

        For Each Block In blockList
            length += Block.calculeLenght
        Next

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

    Public Function getNextCom(ci As Integer, cj As Integer, xlSheet As Worksheet, h As Double) As List(Of Block)

        Dim lista As List(Of Block) = New List(Of Block)

        Dim j As Integer = cj + 4
        Dim i As Integer = ci + 1
        Dim first As Boolean = False

        While Not String.IsNullOrEmpty(xlSheet.Cells(j, i).Value)

            Select Case xlSheet.Cells(j, i).Value

                Case "F"
                    lista.Add(New F(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FB"
                    lista.Add(New FB(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FS"
                    lista.Add(New FS(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "FPRC"
                    lista.Add(New FPRC(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, first, draw))
                Case "T"
                    lista.Add(New T(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "A"
                    lista.Add(New A(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "C"
                    lista.Add(New C(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Ci"
                    lista.Add(New Ci(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Av"
                    lista.Add(New Av(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "FFCC"
                    lista.Add(New FFCC(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "L"
                    lista.Add(New L(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h))
                Case "Com"
                    lista.Add(New Com(xlSheet.Cells(j - 4, i).Value, xlSheet.Cells(j - 3, i).Value, xlSheet.Cells(j - 2, i).Value, xlSheet.Cells(j - 1, i).Value, xlSheet.Cells(j, i).Value, draw, h, Me, ci, cj, xlSheet))
            End Select

            first = True
            i += 1
            com = i
        End While

        Return lista

    End Function

End Class