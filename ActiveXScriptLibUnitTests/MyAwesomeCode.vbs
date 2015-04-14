Option Explicit

Dim obj
Set obj = CreateObject("Blah")

obj.Echo "Echo from MyAwesomeCode.vbs"


Public Function TestCreateObject(progId)
   Set TestCreateObject = CreateObject(progId)
End Function
