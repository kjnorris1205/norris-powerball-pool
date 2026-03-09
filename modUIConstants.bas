Attribute VB_Name = "modUIConstants"
Option Compare Database
Option Explicit

' ======================================================================================
'  COLOR CONSTANTS
'  Long values computed as R + G*256 + B*65536
' ======================================================================================
Public Const UI_COLOR_PRIMARY As Long = 6697728         ' RGB(0, 51, 102) - Dark navy
Public Const UI_COLOR_ACCENT As Long = 10053171         ' RGB(51, 102, 153) - Medium blue
Public Const UI_COLOR_TEXT As Long = 0                   ' RGB(0, 0, 0) - Black
Public Const UI_COLOR_TEXT_LIGHT As Long = 8421504       ' RGB(128, 128, 128) - Gray
Public Const UI_COLOR_BACKGROUND As Long = 16777215      ' RGB(255, 255, 255) - White
Public Const UI_COLOR_HEADER_BG As Long = 15790320       ' RGB(240, 240, 240) - Light gray
Public Const UI_COLOR_BUTTON_BG As Long = 15849869       ' RGB(13, 110, 241) - System button
Public Const UI_COLOR_HIGHLIGHT As Long = 15853276       ' RGB(220, 230, 241) - Light blue

' ======================================================================================
'  FONT CONSTANTS
' ======================================================================================
Public Const UI_FONT_NAME As String = "Segoe UI"
Public Const UI_FONT_SIZE_TITLE As Integer = 18
Public Const UI_FONT_SIZE_HEADING As Integer = 12
Public Const UI_FONT_SIZE_BODY As Integer = 10

' ======================================================================================
'  LAYOUT CONSTANTS (twips: 1 inch = 1440 twips)
' ======================================================================================
Public Const UI_MARGIN As Long = 300                     ' ~0.2 inches
Public Const UI_CONTROL_SPACING As Long = 150            ' ~0.1 inches
Public Const UI_ROW_HEIGHT As Long = 450                 ' ~0.31 inches (row pitch)
Public Const UI_LABEL_WIDTH As Long = 2000               ' ~1.4 inches
Public Const UI_TEXTBOX_HEIGHT As Long = 360             ' ~0.25 inches
Public Const UI_BUTTON_WIDTH As Long = 2400              ' ~1.7 inches
Public Const UI_BUTTON_HEIGHT As Long = 420              ' ~0.29 inches
