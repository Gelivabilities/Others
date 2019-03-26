VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScrollbarDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'

Option Explicit
'自定义的滚动条
Private Const SBackH = 0
Private Const SThumbH = 1
Private Const STGripH = 2
Private Const SArrowH = 3
Private Const SArrowL = 4
Private Const SArrowR = 5
Private Const SBackV = 6
Private Const SThumbV = 7
Private Const STGripV = 8
Private Const SArrowV = 9
Private Const SArrowT = 10
Private Const SArrowB = 11

Private BarPics(11) As New pcMemDC

Private Sub Class_Initialize()
  BarPics(SBackH).CreateFromPicture LoadResPicture("SCRBACKH", vbResBitmap)
  BarPics(SThumbH).CreateFromPicture LoadResPicture("SCRTHUMBH", vbResBitmap)
  BarPics(STGripH).CreateFromPicture LoadResPicture("SCRGRIPH", vbResBitmap)
  BarPics(SArrowH).CreateFromPicture LoadResPicture("SCRARROWH", vbResBitmap)
  BarPics(SArrowL).CreateFromPicture LoadResPicture("SCRARROWHL", vbResBitmap)
  BarPics(SArrowR).CreateFromPicture LoadResPicture("SCRARROWHR", vbResBitmap)
  
  BarPics(SBackV).CreateFromPicture LoadResPicture("SCRBACKV", vbResBitmap)
  BarPics(SThumbV).CreateFromPicture LoadResPicture("SCRTHUMBV", vbResBitmap)
  BarPics(STGripV).CreateFromPicture LoadResPicture("SCRGRIPV", vbResBitmap)
  BarPics(SArrowV).CreateFromPicture LoadResPicture("SCRARROWV", vbResBitmap)
  BarPics(SArrowT).CreateFromPicture LoadResPicture("SCRARROWVT", vbResBitmap)
  BarPics(SArrowB).CreateFromPicture LoadResPicture("SCRARROWVB", vbResBitmap)
End Sub

Friend Sub DrawBack(lhDC As Long, mOrientation As Integer, mIndex As Integer, theRect As RECT)
    If mOrientation = 1 Then 'Horizontal
        BarPics(SBackH).StretchDraw lhDC, theRect.Left, theRect.Top, theRect.Right - theRect.Left, 3, mIndex * 2, 0, 2, 3
        BarPics(SBackH).StretchDraw lhDC, theRect.Left, theRect.Bottom - 3, theRect.Right - theRect.Left, 3, mIndex * 2, 17 - 3, 2, 3
        BarPics(SBackH).StretchDraw lhDC, theRect.Left, theRect.Top + 3, theRect.Right - theRect.Left, theRect.Bottom - theRect.Top - 6, mIndex * 2, 3, 2, 17 - 6
      Else
        BarPics(SBackV).StretchDraw lhDC, theRect.Left, theRect.Top, 3, theRect.Bottom - theRect.Top, 0, mIndex * 2, 3, 2
        BarPics(SBackV).StretchDraw lhDC, theRect.Right - 3, theRect.Top, 3, theRect.Bottom - theRect.Top, 17 - 3, mIndex * 2, 3, 2
        BarPics(SBackV).StretchDraw lhDC, theRect.Left + 3, theRect.Top, theRect.Right - theRect.Left - 6, theRect.Bottom - theRect.Top, 3, mIndex * 2, 17 - 6, 2
    End If
End Sub

Friend Sub DrawBar(lhDC As Long, mOrientation As Integer, mIndex As Integer, theRect As RECT)
    If mOrientation = 1 Then 'Horizontal
        BarPics(SThumbH).Draw lhDC, 0, mIndex * 17, 3, 3, theRect.Left, theRect.Top
        BarPics(SThumbH).Draw lhDC, 0, (mIndex * 17) + 17 - 3, 3, 3, theRect.Left, theRect.Bottom - 3
        BarPics(SThumbH).StretchDraw lhDC, theRect.Left, theRect.Top + 3, 3, theRect.Bottom - theRect.Top - 6, 0, (mIndex * 17) + 3, 3, 17 - 6
        
        BarPics(SThumbH).StretchDraw lhDC, theRect.Left + 3, theRect.Top, theRect.Right - theRect.Left - 6, 3, 3, mIndex * 17, 15 - 6, 3
        BarPics(SThumbH).StretchDraw lhDC, theRect.Left + 3, theRect.Bottom - 3, theRect.Right - theRect.Left - 6, 3, 3, (mIndex * 17) + 17 - 3, 15 - 6, 3
        BarPics(SThumbH).StretchDraw lhDC, theRect.Left + 3, theRect.Top + 3, theRect.Right - theRect.Left - 6, theRect.Bottom - theRect.Top - 6, 3, (mIndex * 17) + 3, 15 - 6, 17 - 6
      
        BarPics(SThumbH).Draw lhDC, 15 - 3, mIndex * 17, 3, 3, theRect.Right - 3, theRect.Top
        BarPics(SThumbH).Draw lhDC, 15 - 3, (mIndex * 17) + 17 - 3, 3, 3, theRect.Right - 3, theRect.Bottom - 3
        BarPics(SThumbH).StretchDraw lhDC, theRect.Right - 3, theRect.Top + 3, 3, theRect.Bottom - theRect.Top - 6, 15 - 3, (mIndex * 17) + 3, 3, 17 - 6
      
        If theRect.Right - theRect.Left > 15 Then
            BarPics(STGripH).Draw lhDC, 0, mIndex * 9, 10, 9, ((theRect.Right - theRect.Left - 10) / 2) + theRect.Left, (theRect.Bottom - theRect.Top - 9) / 2, True
        End If
      Else
        BarPics(SThumbV).Draw lhDC, mIndex * 17, 0, 3, 3, theRect.Left, theRect.Top
        BarPics(SThumbV).Draw lhDC, (mIndex * 17) + 17 - 3, 0, 3, 3, theRect.Right - 3, theRect.Top
        BarPics(SThumbV).StretchDraw lhDC, theRect.Left + 3, theRect.Top, theRect.Right - theRect.Left - 6, 3, (mIndex * 17) + 3, 0, 17 - 6, 3
        
        BarPics(SThumbV).StretchDraw lhDC, theRect.Left, theRect.Top + 3, 3, theRect.Bottom - theRect.Top - 6, mIndex * 17, 3, 3, 15 - 6
        BarPics(SThumbV).StretchDraw lhDC, theRect.Right - 3, theRect.Top + 3, 3, theRect.Bottom - theRect.Top - 6, (mIndex * 17) + 17 - 3, 3, 3, 15 - 6
        BarPics(SThumbV).StretchDraw lhDC, theRect.Left + 3, theRect.Top + 3, theRect.Right - theRect.Left - 6, theRect.Bottom - theRect.Top - 6, (mIndex * 17) + 3, 3, 17 - 6, 15 - 6
    
        BarPics(SThumbV).Draw lhDC, mIndex * 17, 15 - 3, 3, 3, theRect.Left, theRect.Bottom - 3
        BarPics(SThumbV).Draw lhDC, (mIndex * 17) + 17 - 3, 15 - 3, 3, 3, theRect.Right - 3, theRect.Bottom - 3
        BarPics(SThumbV).StretchDraw lhDC, theRect.Left + 3, theRect.Bottom - 3, theRect.Right - theRect.Left - 6, 3, (mIndex * 17) + 3, 15 - 3, 17 - 6, 3
    
        If theRect.Bottom - theRect.Top > 15 Then
            BarPics(STGripV).Draw lhDC, 0, mIndex * 10, 9, 10, (theRect.Right - theRect.Left - 9) / 2, ((theRect.Bottom - theRect.Top - 10) / 2) + theRect.Top, True
        End If
    End If
End Sub

Friend Sub DrawTopLeftButton(lhDC As Long, mOrientation As Integer, mIndex As Integer, theRect As RECT)
    If mOrientation = 1 Then 'Horizontal
        BarPics(SArrowH).Draw lhDC, mIndex * 18, 0, 18, 3, theRect.Left, theRect.Top
        BarPics(SArrowH).Draw lhDC, mIndex * 18, 17 - 3, 18, 3, theRect.Left, theRect.Bottom - 3
        BarPics(SArrowH).StretchDraw lhDC, theRect.Left, theRect.Top + 3, 18, theRect.Bottom - theRect.Top - 6, mIndex * 18, 3, 18, 17 - 6
        BarPics(SArrowL).Draw lhDC, mIndex * 18, 0, 18, 17, theRect.Left, (theRect.Bottom - theRect.Top - 17) / 2, True
      Else
        BarPics(SArrowV).Draw lhDC, 0, mIndex * 18, 3, 18, theRect.Left, theRect.Top
        BarPics(SArrowV).Draw lhDC, 17 - 3, mIndex * 18, 3, 18, theRect.Right - 3, theRect.Top
        BarPics(SArrowV).StretchDraw lhDC, theRect.Left + 3, theRect.Top, theRect.Right - theRect.Left - 6, 18, 3, mIndex * 18, 17 - 6, 18
        BarPics(SArrowT).Draw lhDC, 0, mIndex * 18, 17, 18, (theRect.Right - theRect.Left - 17) / 2, theRect.Left, True
    End If
End Sub

Friend Sub DrawBottomRightButton(lhDC As Long, mOrientation As Integer, mIndex As Integer, theRect As RECT)
    If mOrientation = 1 Then 'Horizontal
        BarPics(SArrowH).Draw lhDC, mIndex * 18, 0, 18, 3, theRect.Left, theRect.Top
        BarPics(SArrowH).Draw lhDC, mIndex * 18, 17 - 3, 18, 3, theRect.Left, theRect.Bottom - 3
        BarPics(SArrowH).StretchDraw lhDC, theRect.Left, theRect.Top + 3, 18, theRect.Bottom - theRect.Top - 6, mIndex * 18, 3, 18, 17 - 6
        BarPics(SArrowR).Draw lhDC, mIndex * 18, 0, 18, 17, theRect.Left, (theRect.Bottom - theRect.Top - 17) / 2, True
      Else
        BarPics(SArrowV).Draw lhDC, 0, mIndex * 18, 3, 18, theRect.Left, theRect.Top
        BarPics(SArrowV).Draw lhDC, 17 - 3, mIndex * 18, 3, 18, theRect.Right - 3, theRect.Top
        BarPics(SArrowV).StretchDraw lhDC, theRect.Left + 3, theRect.Top, theRect.Right - theRect.Left - 6, 18, 3, mIndex * 18, 17 - 6, 18
        BarPics(SArrowB).Draw lhDC, 0, mIndex * 18, 17, 18, (theRect.Right - theRect.Left - 17) / 2, theRect.Top, True
    End If
End Sub

