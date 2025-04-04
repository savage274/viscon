Public Sub AddCenterConnectionPointToSelectedShapes()

    Dim vsoSelection As Visio.Selection
    Dim vsoShape As Visio.Shape
	
	Set vsoSelection = ActiveWindow.Selection
	
	    For Each vsoShape In vsoSelection

			Dim NewRow as Integer
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0.5"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0.5"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0.5"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*1"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0.5"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0.5"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0.5"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*1"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*1"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*1"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*1"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*1"
			NewRow = vsoShape.AddRow( visSectionConnectionPts , visRowLast, visTagDefault)
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visX).formula = "Width*0"
			vsoShape.CellsSRC( visSectionConnectionPts, NewRow, visY).formula = "Height*0"
			Next vsoShape
	
	
Exit Sub

End Sub
