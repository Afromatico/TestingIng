Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Class1




    <CommandMethod("ADDDISTANCE")> _
    Public Sub addDistance()
        selectSelecction()
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

                MsgBox("Entity is on layer " & myent.Layer)
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


End Class
