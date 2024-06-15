Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Set obBHDoc = obWCAD.GetActiveBorehole()


Set obLogOne = obBHDoc.InsertNewLog(2)
obLogOne.Name = "Adjusted Fluid Cond"
obLogOne.Formula = "{Fluid Cond}/10000"

obBHDoc.RemoveLog "Fluid Cond"

Set obLogTwo = obBHDoc.InsertNewLog(2)
obLogTwo.Name = "Fluid Cond"
obLogTwo.Formula = "If({Adjusted Fluid Cond} < 0, 0, {Adjusted Fluid Cond})"
obLogTwo.LogUnit = "S/m"

obBHDoc.RemoveLog "Adjusted Fluid Cond"



Set obLogThree = obBHDoc.InsertNewLog(2)
obLogThree.Name = "Adjusted Fluid Cond 25C"
obLogThree.Formula = "{Fluid Cond 25C}/10000"

obBHDoc.RemoveLog "Fluid Cond 25C"

Set obLogFour = obBHDoc.InsertNewLog(2)
obLogFour.Name = "Fluid Cond 25C"
obLogFour.Formula = "If({Adjusted Fluid Cond 25C} < 0, 0, {Adjusted Fluid Cond 25C})"
obLogFour.LogUnit = "S/m"

obBHDoc.RemoveLog "Adjusted Fluid Cond 25C"