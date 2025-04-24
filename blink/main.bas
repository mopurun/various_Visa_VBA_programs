Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
Sub ボタン1_Click()
    Dim RM As New VisaComLib.ResourceManager
    Dim INST As New VisaComLib.FormattedIO488
    
    Const VisaResourceName As String = "USB0::~~::~~::~~::INSTR" 'Input VisaResourceName here
    
    Set INST.IO = RM.Open(VisaResourceName)
    
    INST.WriteString "*IDN?"
    Dim idn As String
    idn = INST.ReadString
    Cells(1, 1) = Replace(idn, vbLf, "")
    
    
    
    
    For i = 1 To 5
        Call Sleep(1000)
        INST.WriteString ":STOP"
        Call Sleep(1000)
        INST.WriteString ":RUN"
    Next

 
    
    
    
    
    INST.IO.Close
    Set INST = Nothing
    Set RM = Nothing
End Sub
