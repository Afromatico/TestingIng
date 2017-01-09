Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry

Public Class Class1




    <CommandMethod("ADDDISTANCE")> _
    Public Sub addDistance()
        '' Get the current database
        Dim acCurDb As Object = HostApplicationServices.WorkingDatabase

        '' Create a dynamic reference to model or paper space
        Dim acSpace As Object = acCurDb.CurrentSpaceId

        '' Create a line that starts at 5,5 and ends at 12,3
        Dim acLine As Object = New Line(New Point3d(5, 5, 0),
                                        New Point3d(12, 3, 0))

        '' Add the new object to the current space
        acSpace.AppendEntity(acLine)
    End Sub


End Class
