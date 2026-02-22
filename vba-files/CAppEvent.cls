'@Folder("MORFunctions.Classes")
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAppEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents EventApp As Excel.Application
Attribute EventApp.VB_VarHelpID = -1

Private Sub EventApp_SheetActivate(ByVal Sh As Object)
    Set_All_Charts
End Sub
 
Private Sub EventApp_SheetDeactivate(ByVal Sh As Object)
    Reset_All_Charts
End Sub
 
Private Sub EventApp_WorkbookActivate(ByVal Wb As Workbook)
    Set_All_Charts
End Sub
 
Private Sub EventApp_WorkbookDeactivate(ByVal Wb As Workbook)
    Reset_All_Charts
End Sub
