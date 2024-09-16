Attribute VB_Name = "Xport"
Option Explicit

'VBACallGraphSetupProject.Xport

Public Sub ExportThisWorkbook()
    
    Call VBACallGraphSetup.ExportModules(ThisWorkbook)

End Sub
