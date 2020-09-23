Attribute VB_Name = "Module1"
Public Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long


Public Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal bRedraw As Boolean) As Long
    
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Plotting As Boolean

Public Sub DrawMandelBrautSet(TheForm As Object, _
                              Optional ByVal ColorMax As Long, _
                              Optional ByVal ColorStep As Long, _
                              Optional ByVal ComplexPlain_X1 As Currency, _
                              Optional ByVal ComplexPlain_Y1 As Currency, _
                              Optional ByVal ComplexPlain_X2 As Currency, _
                              Optional ByVal ComplexPlain_Y2 As Currency, _
                              Optional ByVal ColorIncrement As Integer, _
                              Optional ByVal ColorStart As Long)
                              
  'This sub draws a mandelbrot set on some object(a form or a picturebox works
  'best).  The optional values are there to change the appearance of the fractal.
  'If they are not passed or values outside of the acceptable range are passed
  'then the variables are set to the correct default value as determined by the
  'function of that variable.  This sub was picked up on the internet where I assume
  'it was posted by the author and http://www.planetsourcecode.com/.
  'I have modified the sub to make it more generalized
  'and to allow the possibility of changing the appearance of the fractal on the
  'fly.  The latest modifications were made by me, Ken Knutson(thevbman@earthlink.net)
  'on July 20, 2000 at 12:42 PM.  If you have questions or comments (no griping or
  'bellyaching) email me and I'll respond if I can.  Best of enjoyment.
  
  'These four variables define the rectangular region of the complex plain that
  'will be iterated.  Change the values to zoom in/out.  Check the incoming
  'values and set to default values if they are zero
  '   ComplexPlain_X1
  '   ComplexPlain_Y1
  '   ComplexPlain_X2
  '   ComplexPlain_Y2
  
  'These two variables are used to store the ScaleWidth and ScaleHeight values,for
  'faster access:
  '   ScreenWidth
  '   ScreenHeight
  
  'These two variables reflect the X and Y intervals of the loop that moves from
  '(ComplexPlain_X1,ComplexPlain_Y1) to
  '(ComplexPlain_X2,ComplexPlain_Y2) in the complex plain.
  '   StepX
  '   StepY
  
  'These two are used in the main loop.
  '   X
  '   Y
  
  'Cx and Cy are the real and imaginary part respectively of C,in the function
  ' Zv=Zv-1^2 + C
  
  'Zx and Zy are the real and imaginary part respectively of Z,in the function
  ' Zv=Zv-1^2 + C
  
  'This byte variable is assigned a number for each pixel in the form.
  '   Color
  
  'Used in the function that we iterate.
  '   TempX
  '   TempY
  
  Dim TempX As Currency
  Dim TempY As Currency
  Dim Color As Long
  Dim Zx As Currency
  Dim Zy As Currency
  Dim Cx As Currency
  Dim Cy As Currency
  Dim X As Currency
  Dim Y As Currency
  Dim StepX As Currency
  Dim StepY As Currency
  Dim ScreenWidth As Integer
  Dim ScreenHeight As Integer

                              
  'Check for empty and out of range values and correct them
  ColorMax = LimitVal(ColorMax, 255, 1, 255)
  ColorStep = LimitVal(ColorStep, 1677215, 1, 100)
  ColorIncrement = LimitVal(ColorIncrement, 32767, 0.1, 1)
  ComplexPlain_X1 = LimitVal(ComplexPlain_X1, 1677215, 0.1, -2)
  ComplexPlain_Y1 = LimitVal(ComplexPlain_Y1, 1677215, 0.1, 2)
  ComplexPlain_X2 = LimitVal(ComplexPlain_X2, 1677215, 0.1, 2)
  ComplexPlain_Y2 = LimitVal(ComplexPlain_Y2, 1677215, 0.1, -2)
  
  'set the object up so it will draw correctly
  TheForm.AutoRedraw = True
  TheForm.ScaleMode = 3
  
  'Set the values fo the draw area dimension variables for faster acess
  ScreenWidth = TheForm.ScaleWidth
  ScreenHeight = TheForm.ScaleHeight
  
  'Calculate the intervals of the loop.
  StepX = Abs(ComplexPlain_X2 - ComplexPlain_X1) / ScreenWidth
  StepY = Abs(ComplexPlain_Y2 - ComplexPlain_Y1) / ScreenHeight
  
  'Clear the object.
  TheForm.Cls
  
  'Set the plotting variable which indicates that we are currently doing the plot
  Plotting = True
  
  For X = 0 To ScreenWidth
    For Y = 0 To ScreenHeight
      Cx = ComplexPlain_X1 + X * StepX
      Cy = ComplexPlain_Y2 + Y * StepY
      Zx = 0
      Zy = 0
      Color = ColorStart
      While (Not (Zx * Zx + Zy * Zy > 4)) And Color < ColorMax And Plotting
        TempX = Zx
        TempY = Zy
        Zx = TempX * TempX - TempY * TempY + Cx
        Zy = 2 * TempX * TempY + Cy
        Color = Color + ColorIncrement 'was color + 1
      Wend
      If Not Plotting Then Exit Sub
      SetPixel TheForm.hdc, X, Y, Color * ColorStep
    Next Y
    TheForm.Refresh
    DoEvents
  Next X
  
  Plotting = False
End Sub

Public Function LimitVal(ByVal Val2Limit As Double, _
                         ByVal UpperLimit As Double, _
                         ByVal LowerLimit As Double, _
                         ByVal DefaultValue As Double) As Double
                         
  Dim SetToDefault As Boolean
  
  If Val2Limit < LowerLimit Then SetToDefault = True
  If Val2Limit > UpperLimit Then SetToDefault = True
  
  If SetToDefault Then
    LimitVal = DefaultValue
  Else
    LimitVal = Val2Limit
  End If
  
End Function
