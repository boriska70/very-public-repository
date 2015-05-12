Dim fso, file, file_location, data, myrandom, newData
file_location = "state.txt"
Set fso = CreateObject("Scripting.FileSystemObject")

rem create file, if missing and put initial value
if not fso.FileExists(file_location) Then
Set file = fso.CreateTextFile(file_location, true)
file.Close
Set file= fso.OpenTextFile(file_location, 2, true)
file.Write("0")
file.Close
End If

rem get current value
Set file= fso.OpenTextFile(file_location, 1, true)
data = file.Read(1)
file.Close

rem increase value by 1 or start over
newData=CInt(data)+1
If newData>3 Then
newData=0
Reporter.ReportEvent micFail, "PeriodicalFailure", "This test fails once in 4 runs" 
Else
Reporter.ReportEvent micPass, "MostlySuccessful", "This test succeeds 3 times out of 4" 
End If

Set file= fso.OpenTextFile(file_location, 2, true)
file.Write(newData)
file.Close

Wait(5)


rem myrandom=RandomNumber.Value(1,4)
rem Reporter.ReportEvent micFail, "RandomFailure", "This test fails sometimes" 
