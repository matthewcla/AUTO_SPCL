Attribute VB_Name = "modUserFormPosition"
'======== Module: modUserFormPosition ========
' Centers a UserForm on the monitor that contains the active Excel window.

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MonitorFromWindow Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function GetMonitorInfoW Lib "user32" ( _
        ByVal hMonitor As LongPtr, ByRef lpmi As MONITORINFO) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
#Else
    Private Declare Function MonitorFromWindow Lib "user32" ( _
        ByVal hWnd As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetMonitorInfoW Lib "user32" ( _
        ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
        ByVal hDC As Long, ByVal nIndex As Long) As Long
#End If

Private Const MONITOR_DEFAULTTONEAREST As Long = 2
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type MONITORINFO
    cbSize     As Long
    rcMonitor  As RECT
    rcWork     As RECT   ' work area, excludes taskbar/docked bars
    dwFlags    As Long
End Type

' --- Pixel -> Point conversion using system DPI (Excel is system-DPI aware) ---
Private Function PixelsToPointsX(ByVal px As Long) As Double
#If VBA7 Then
    Dim hDC As LongPtr
#Else
    Dim hDC As Long
#End If
    Dim dpi As Long
    hDC = GetDC(0&)
    dpi = GetDeviceCaps(hDC, LOGPIXELSX)
    Call ReleaseDC(0&, hDC)
    PixelsToPointsX = (px * 72#) / dpi
End Function

Private Function PixelsToPointsY(ByVal py As Long) As Double
#If VBA7 Then
    Dim hDC As LongPtr
#Else
    Dim hDC As Long
#End If
    Dim dpi As Long
    hDC = GetDC(0&)
    dpi = GetDeviceCaps(hDC, LOGPIXELSY)
    Call ReleaseDC(0&, hDC)
    PixelsToPointsY = (py * 72#) / dpi
End Function

' Public API: center a form on the monitor containing the active Excel window
Public Sub CenterUserFormOnActiveMonitor(ByVal frm As Object)
#If VBA7 Then
    Dim hWndXL As LongPtr
    Dim hMon As LongPtr
#Else
    Dim hWndXL As Long
    Dim hMon As Long
#End If
    Dim mi As MONITORINFO

    ' Identify the monitor that contains the Excel main window
    hWndXL = Application.hWnd
    hMon = MonitorFromWindow(hWndXL, MONITOR_DEFAULTTONEAREST)

    ' Populate monitor info (work area in pixels)
    mi.cbSize = Len(mi)
    If GetMonitorInfoW(hMon, mi) = 0 Then Exit Sub  ' fail-safe

    ' Convert work area to points once; then do all math in points
    Dim workLeftPts As Double, workTopPts As Double
    Dim workWidthPts As Double, workHeightPts As Double

    workLeftPts = PixelsToPointsX(mi.rcWork.Left)
    workTopPts = PixelsToPointsY(mi.rcWork.Top)
    workWidthPts = PixelsToPointsX(mi.rcWork.Right - mi.rcWork.Left)
    workHeightPts = PixelsToPointsY(mi.rcWork.Bottom - mi.rcWork.Top)

    ' Center the form in the work area
    frm.StartUpPosition = 0
    frm.Left = workLeftPts + (workWidthPts - frm.Width) / 2
    frm.Top = workTopPts + (workHeightPts - frm.Height) / 2
End Sub
'======== End Module ========


