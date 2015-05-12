myrandom=RandomNumber.Value(1,5)
MsgBox myrandom
If myrandom=4 Then
	Reporter.ReportEvent micFail, "RandomFailure", "This test fails sometimes" 
Else
	Reporter.ReportEvent micPass, "MostlySuccess", "This test mostly succeeds" 
End If

