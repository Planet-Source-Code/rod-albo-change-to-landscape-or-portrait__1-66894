VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Printer"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   1560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChangePrinterOrient 
      Caption         =   "Change Printer Orientation"
      Height          =   615
      Left            =   150
      TabIndex        =   2
      Top             =   1020
      Width           =   975
   End
   Begin VB.OptionButton optLand 
      Caption         =   "Landscape"
      Height          =   495
      Left            =   165
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton optPort 
      Caption         =   "Portrait"
      Height          =   495
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constants used in the PrinterMode structure
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

'Constants for NT security
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

'Constants used to make changes to the values contained in the PrinterMode
Private Const DM_MODIFY = 8
Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_COPY = 2
Private Const DM_OUT_BUFFER = DM_COPY
Private Const DM_DUPLEX = &H1000&
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3
Private Const DM_ORIENTATION = &H1&
Private PageDirection As Integer

Private Type PrinterMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmLogPixels As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long        ' // Windows 95 only
    dmICMIntent As Long        ' // Windows 95 only
    dmMediaType As Long        ' // Windows 95 only
    dmDitherType As Long       ' // Windows 95 only
    dmReserved1 As Long        ' // Windows 95 only
    dmReserved2 As Long        ' // Windows 95 only
End Type

Private Type PRINTER_DEFAULTS
'Note:
'  The definition of Printer_Defaults in the VB5 API viewer is incorrect.
'  Below, pPrinterMode has been corrected to LONG.
    pDatatype As String
    pPrinterMode As Long
    DesiredAccess As Long
End Type


'------DECLARATIONS

Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long
Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long

Private Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, ByVal pPrinterModeOutput As Any, ByVal pPrinterModeInput As Any, ByVal fMode As Long) As Long

Private Sub OrientationSettings(NewSetting As Long, chng As Integer, ByVal frm As Form)
    Dim PrinterHandle As Long
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim MyPrinterMode As PrinterMode
    Dim Result As Long
    Dim Needed As Long
    Dim pFullPrinterMode As Long
    Dim pi2_buffer() As Long
    
    PrinterName = Printer.DeviceName
    If PrinterName = "" Then
        Exit Sub
    End If
    
    pd.pDatatype = vbNullString
    pd.pPrinterMode = 0&
    'Printer_Access_All is required for NT security
    pd.DesiredAccess = PRINTER_ALL_ACCESS
    
    Result = OpenPrinter(PrinterName, PrinterHandle, pd)
    
    'The first call to GetPrinter gets the size, in bytes, of the buffer needed.
    'This value is divided by 4 since each element of pi2_buffer is a long.
    Result = GetPrinter(PrinterHandle, 2, ByVal 0&, 0, Needed)
    ReDim pi2_buffer((Needed \ 4))
    Result = GetPrinter(PrinterHandle, 2, pi2_buffer(0), Needed, Needed)
    
    'The seventh element of pi2_buffer is a Pointer to a block of memory
    '  which contains the full PrinterMode (including the PRIVATE portion).
    pFullPrinterMode = pi2_buffer(7)
    
    'Copy the Public portion of FullPrinterMode into our PrinterMode structure
    Call CopyMemory(MyPrinterMode, ByVal pFullPrinterMode, Len(MyPrinterMode))
    
    'Make desired changes
    MyPrinterMode.dmDuplex = NewSetting
    MyPrinterMode.dmFields = DM_DUPLEX Or DM_ORIENTATION
    MyPrinterMode.dmOrientation = chng
    
    'Copy our PrinterMode structure back into FullPrinterMode
    Call CopyMemory(ByVal pFullPrinterMode, MyPrinterMode, Len(MyPrinterMode))
    
    'Copy our changes to "the PUBLIC portion of the PrinterMode" into "the PRIVATE portion of the PrinterMode"
    Result = DocumentProperties(frm.hwnd, PrinterHandle, PrinterName, ByVal pFullPrinterMode, ByVal pFullPrinterMode, DM_IN_BUFFER Or DM_OUT_BUFFER)
    
    'Update the printer's default properties (to verify, go to the Printer folder
    '  and check the properties for the printer)
    Result = SetPrinter(PrinterHandle, 2, pi2_buffer(0), 0&)
    
    Call ClosePrinter(PrinterHandle)
    
    'Note: Once "Set Printer = " is executed, anywhere in the code, after that point
    '      changes made with SetPrinter will ONLY affect the system-wide printer  --
    '      -- the changes will NOT affect the VB printer object.
    '      Therefore, it may be necessary to reset the printer object's parameters to
    '      those chosen in the PrinterMode.
    Dim p As Printer
    For Each p In Printers
        If p.DeviceName = PrinterName Then
            Set Printer = p
            Exit For
        End If
    Next p
    Printer.Duplex = MyPrinterMode.dmDuplex
End Sub

Public Sub ChngPrinterOrientationLandscape(ByVal frm As Form)
    PageDirection = 2
    Call OrientationSettings(DMDUP_SIMPLEX, PageDirection, frm)
End Sub

Public Sub ResetPrinterOrientation(ByVal frm As Form)
 
    If PageDirection = 1 Then
        PageDirection = 2
    Else
        PageDirection = 1
    End If
    Call OrientationSettings(DMDUP_SIMPLEX, PageDirection, frm)
End Sub

Public Sub ChngPrinterOrientationPortrait(ByVal frm As Form)

    PageDirection = 1
    Call OrientationSettings(DMDUP_SIMPLEX, PageDirection, frm)
End Sub


Private Sub cmdChangePrinterOrient_Click()
      
    If optPort.Value = True Then
        ChngPrinterOrientationPortrait Me
    ElseIf optLand.Value = True Then
        ChngPrinterOrientationLandscape Me
    End If
        
End Sub

