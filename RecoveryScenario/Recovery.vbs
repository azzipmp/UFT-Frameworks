Function RecoveryFunction1(Object, Method, Arguments, retVal)
msgbox "recover fired"
Desktop.CaptureBitmap "C:\Users\amohamme\Desktop\UFT_Framework\LCAF\TestResults\ErrorSnapshot\recovery"&".png",true
'Environment("Recovery") = "True"
exitaction(0)
End Function
