Attribute VB_Name = "Module1"
Option Explicit

Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long


Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)
Dim hrgn As Long, tRgn As Long
Dim X As Integer, Y As Integer, X0 As Integer
Dim hDC As Long, BM As BITMAP

    hDC = CreateCompatibleDC(0)
    If hDC Then
        SelectObject hDC, cPicture
        GetObject cPicture, Len(BM), BM
        hrgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
        For Y = 0 To BM.bmHeight
            For X = 0 To BM.bmWidth
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
                    X = X + 1
                Wend
                X0 = X
                While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
                    X = X + 1
                Wend
                If X0 < X Then
                    tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                    CombineRgn hrgn, hrgn, tRgn, 4
                    DeleteObject tRgn
                End If
            Next X
        Next Y
    
        GetBitmapRegion = hrgn
        DeleteObject SelectObject(hDC, cPicture)
    End If
       
    DeleteDC hDC
End Function

