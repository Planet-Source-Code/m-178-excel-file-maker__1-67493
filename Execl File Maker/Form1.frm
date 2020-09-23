VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim colu As Byte
Dim rw As Byte
'made by Metaferia Amhayesus metaferia@gmail.com
Dim ef1 As New ExcelFile

With ef1
    .OpenFile "vbtest.xls"
    .EWriteInteger 1, 1, 100
    .EWriteString 1, 2, "Test writing a string"
    .CloseFile
    MsgBox "Your XLS File Has Been Made"
End With

End
End Sub
