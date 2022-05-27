VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Развёртка"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "MainForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExecAllButton_Click()
  'вызываем отрисовку треугольника, скругляем, всё остальное
  
  ExecTriButton_Click
  
  ActiveDocument.BeginCommandGroup "Скругление"
  LtriOrig.Curve.Nodes.All.Fillet FILLETR
  ActiveDocument.EndCommandGroup
  
  executeRazv LtriOrig
  
End Sub

Private Sub ExecFinalButton_Click()
  Dim sr As ShapeRange
  Set sr = ActiveSelectionRange
  If sr.Count = 1 Then executeRazv sr(1)
End Sub

Private Sub ExecTriButton_Click()
  'рисуем треугольник
  ActiveDocument.BeginCommandGroup "Построение треугольника"
  Set LtriOrig = drawTriangle(CDbl(side_a), CDbl(side_b), CDbl(side_c), centerx, centery)
  ActiveDocument.EndCommandGroup
End Sub

Private Sub side_a_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbers KeyAscii
End Sub

Private Sub side_b_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbers KeyAscii
End Sub

Private Sub side_c_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyNumbers KeyAscii
End Sub

