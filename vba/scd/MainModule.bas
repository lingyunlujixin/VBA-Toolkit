Attribute VB_Name = "MainModule"
Sub °´Å¥1_Click()

    ' Initialize
    Dim scd As New SCDProducer
    
    ' Logsetting
    Logger.LogEnabled = True
    
    Logger.LogThreshold = 3  'info warn error

    Logger.LogCallback = Array("LogFile", "ImmediateLog")

    Logger.FileName = "d:\debug_test.log"
    
    ' start
    
    ' step 1: Set Target Excel
    'scd.setExcel Worksheets(1), 2, 1
    scd.setExcel Application.ActiveSheet, 2, 1
    ' step 2: Sql-Genaration
    scd.start ""
    
    'Logger.LogDebug scd.formatColumn("ID", "INT", True, "CMT", "-1")
    
    Logger.flush
    
End Sub
