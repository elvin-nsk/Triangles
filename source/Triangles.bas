Attribute VB_Name = "Triangles"
Option Explicit

Public Const FILLETR As Double = 40 'радиус скругления
Public Const CUT_TOP As Double = 31 'отступ от верха до первой распорки
Public Const CUT_MID As Double = 25 'отступ от центральной линии до срединной распорки
Public Const RASP_H As Double = 20 'высота распорки

Public centerx As Double, centery As Double
Public LtriOrig As Shape

  
Sub start()
  
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrCenter
  centerx = ActivePage.SizeWidth / 2
  centery = ActivePage.SizeHeight / 2
    
  MainForm.Show

End Sub

Sub executeRazv(tri As Shape)

  Dim shiftx As Double, triLen As Double
  Dim raspTopAx, raspTopAy, raspTopCx, raspTopCy
  Dim raspMidAx, raspMidAy, raspMidCx, raspMidCy
  Dim LtriTop As Shape, LtriMid As Shape
  Dim raspTop As Shape, raspMid As Shape
  Dim cline As Shape, tline As Shape, ts As Shape
  Dim Rtris As ShapeRange
  Dim txt As String, m As String
  
  ActiveDocument.BeginCommandGroup "Развёртка"
  
  If MainForm.OptionLeft Then tri.Flip cdrFlipHorizontal
  triLen = tri.Curve.Length
  
  'рисуем линию
  Set cline = drawLine(tri, (tri.SizeHeight / 2))
  cline.Outline.Color.CMYKAssign 0, 100, 100, 0
  
  'делаем дубль с отзеркаливанием
  shiftx = tri.SizeWidth * 2
  Set Rtris = tri.DuplicateAsRange(shiftx)
  Rtris.Add cline.Duplicate(shiftx)
  Rtris.Flip cdrFlipHorizontal
  
  Set tline = drawLine(tri, CUT_TOP) 'рисуем временную линию, чтобы взять координаты распорки
  Set LtriTop = sectTop(tri, CUT_TOP)
  tri.Move 0, -RASP_H
  cline.Move 0, -RASP_H
  Set raspTop = drawRasp(tline.LeftX, tline.TopY, tline.RightX, tline.TopY - RASP_H)
  tline.Delete
  
  Set tline = drawLine(tri, tri.TopY - cline.TopY + CUT_MID) 'проделываем этот трюк ещё раз
  Set LtriMid = sectTop(tri, tri.TopY - cline.TopY + CUT_MID)
  tri.Move 0, -RASP_H
  Set raspMid = drawRasp(tline.LeftX, tline.TopY, tline.RightX, tline.TopY - RASP_H)
  tline.Delete
  
  'формируем текст
  If MainForm.OptionRight Then m = "справа" Else m = "слева"
  txt = MainForm.clientName
  txt = txt + vbCrLf + "Цвет — " + MainForm.colorBox
  txt = txt + vbCrLf + "Ширина " + MainForm.volumeBox + " мм"
  txt = txt + vbCrLf + "Молния " + m
  txt = txt + vbCrLf + "Батан " + CStr(Round(triLen, 1)) + " x x " + CStr(Round((CDbl(MainForm.volumeBox) * 2 + CDbl(MainForm.zBox) * 4), 1)) + " мм (ширина слоя — " + CStr(Round((CDbl(MainForm.volumeBox) + CDbl(MainForm.zBox) * 2), 1)) + " мм)"
  txt = txt + vbCrLf + "До базлинии " + CStr(Round((CDbl(MainForm.volumeBox) / 2 - CDbl(MainForm.wBox) / 2 + CDbl(MainForm.zBox)), 1)) + " мм (при стропе " + MainForm.wBox + " мм)"
  txt = txt + vbCrLf + "Перегородка " + CStr(Round((cline.SizeWidth - 10), 1)) + " x " + MainForm.volumeBox + " мм"
  Set ts = ActiveLayer.CreateArtisticText(raspTop.LeftX + 10, raspTop.BottomY - 20, txt, cdrRussian, , , 24)
  ts.SetPosition raspTop.LeftX + raspTop.SizeWidth / 2, raspTop.BottomY - ((raspTop.BottomY - cline.TopY) / 2)
  
  ActiveDocument.EndCommandGroup

End Sub

Function drawTriangle(a, b, c, Cx, Cy) As Shape

  Dim spath As SubPath, crv As Curve
  Dim Ax, Ay, Bx, By, Adx, Ady, p
        
  'рассчитываем координаты вершин
  Bx = Cx - a
  By = Cy
  p = (a + b + c) / 2
  Ady = 2 / a * Sqr(p * (p - a) * (p - b) * (p - c))
  Adx = Sqr(b ^ 2 - Ady ^ 2)
  If a ^ 2 + b ^ 2 - c ^ 2 > 0 Then Ax = Cx - Adx Else Ax = Cx + Adx
  Ay = Cy - Ady
  
  'рисуем треугольник
  Set crv = Application.CreateCurve(ActiveDocument)
  Set spath = crv.CreateSubPath(Cx, Cy)
  spath.AppendLineSegment Bx, By
  spath.AppendLineSegment Ax, Ay
  spath.Closed = True
  
  Set drawTriangle = ActiveLayer.CreateCurve(crv)
  
End Function

Function drawLine(s As Shape, dy) As Shape

  Dim Ax, Bx, y
  Dim spath As SubPath, crv As Curve
  
  y = s.TopY - dy
  Ax = s.LeftX
  Bx = s.RightX
  
  Set crv = Application.CreateCurve(ActiveDocument)
  Set spath = crv.CreateSubPath(Ax, y)
  spath.AppendLineSegment Bx, y
  Set drawLine = s.Intersect(ActiveLayer.CreateCurve(crv), True, False)
  
End Function

Function drawRectangle(Ax, Ay, Cx, Cy) As Shape

  Dim spath As SubPath, crv As Curve
  
  Set crv = Application.CreateCurve(ActiveDocument)
  Set spath = crv.CreateSubPath(Ax, Ay)
  spath.AppendLineSegment Cx, Ay
  spath.AppendLineSegment Cx, Cy
  spath.AppendLineSegment Ax, Cy
  spath.Closed = True
  
  Set drawRectangle = ActiveLayer.CreateCurve(crv)

End Function

Function drawRasp(Ax, Ay, Cx, Cy) As Shape

  Dim spath As SubPath, crv As Curve
  Dim midy
  
  midy = Ay - (Ay - Cy) / 2
  
  Set crv = Application.CreateCurve(ActiveDocument)
  Set spath = crv.CreateSubPath(Ax, Ay)
  spath.AppendLineSegment Ax, Cy
  spath.AppendLineSegment Ax, midy
  spath.AppendLineSegment Cx, midy
  spath.AppendLineSegment Cx, Ay
  spath.AppendLineSegment Cx, Cy
  
  Set drawRasp = ActiveLayer.CreateCurve(crv)

End Function

Function sectTop(ByRef s As Shape, dy) As Shape
  
  Dim r As Shape, newsh As Shape
  
  Set r = drawRectangle(s.LeftX - 10, s.TopY + 10, s.RightX + 10, s.TopY - dy)
  Set newsh = r.Intersect(s, True, True)
  Set s = r.Trim(s, False, False)
  Set sectTop = newsh

End Function

Sub OnlyNumbers(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub




