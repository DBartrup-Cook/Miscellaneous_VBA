Attribute VB_Name = "mdl_ProcessTimer"
Option Compare Database
Option Explicit

Sub CPU_Usage()

    Dim srvEx As Object 'SWbemServicesEx
    Dim xlProcSet As Object 'SWbemObjectSet
    Dim xlPerfSet As Object 'SWbemObjectSet
    Dim objEx As Object 'SWbemObjectEx
    
    Set srvEx = GetObject("winmgmts:root/CIMV2")
    Set xlProcSet = srvEx.ExecQuery("SELECT * FROM Win32_Process WHERE name = 'MSACCESS.EXE'") 'EXCEL.EXE
    Set xlPerfSet = srvEx.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfProc_Process WHERE NAME='MSACCESS'") 'EXCEL
    
    For Each objEx In xlProcSet
        Debug.Print objEx.Name & " RAM: " & objEx.WorkingSetSize / 1024 & "kb"
    Next
    
    For Each objEx In xlPerfSet
        Debug.Print objEx.Name & " CPU: " & objEx.PercentProcessorTime & "%"
    Next

End Sub


