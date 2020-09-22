Attribute VB_Name = "Functions"
Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", ".": End
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting Me.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(Me.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: End
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function

Sub StopService(ServiceName As String)
    a = """" & ServiceName & """"
    'uses the NET STOP function to stop the
    '     service
    Shell "net stop " & a, vbHide
End Sub


Sub StartService(ServiceName As String)
    a = """" & ServiceName & """"
    'uses the NET START function to start th
    '     e service
    Shell "net start " & a, vbHide
End Sub

