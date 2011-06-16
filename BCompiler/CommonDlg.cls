VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' * * * * * * * * * * * * * *
' *                         *
' * Classe clsCommonDialogs *
' * ----------------------- *
' *                         *
' * Por Walter Staeblein    *
' * http://www.codex.com.br *
' *                         *
' * * * * * * * * * * * * * *

Option Explicit

'' Variaveis do programa
Dim mFILE As String
Dim mPATH As String
Dim mFULLPATH As String
Dim mCOLOR As Long
Dim mBOLD As Boolean
Dim mITALIC As Boolean
Dim mSIZE As Integer
Dim mFONT As String
Dim mHWND As Long
Dim mCANCEL As Boolean          '' True se Cancel for Pressionado no Dialog
Dim mFILTER As String
Dim mFromPAGE As Integer
Dim mToPAGE As Integer
Dim mINITDIR As String
Dim mFlags As OFNtype
Dim mPRTFlags As PrinterConstants
Dim mMultipleFiles As Boolean
Dim mFilterIndex As Integer
Dim mDEFEXT As String
Dim mFontFlags As CFType
Dim mUnderline As Boolean
Dim mStrikeThru As Boolean
Dim aCustColor(15) As Long


' Connect & Disconnect net drive
Private Const RESOURCETYPE_DISK = &H1

Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const DM_DUPLEX = &H1000&
Private Const DM_ORIENTATION = &H1&


Private Const FW_BOLD = 700
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_NONETWORKBUTTON = &H200000
Private Const PD_PRINTSETUP = &H40
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_NOWARNING = &H80
Private Const CFERR_CHOOSEFONTCODES = &H2000
Private Const CFERR_MAXLESSTHANMIN = &H2002
Private Const CFERR_NOFONTS = &H2001
Private Const CC_ANYCOLOR = &H100
Private Const CC_CHORD = 4
Private Const CC_CIRCLES = 1
Private Const CC_ELLIPSES = 8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_FULLOPEN = &H2
Private Const CC_INTERIORS = 128
Private Const CC_NONE = 0
Private Const CC_PIE = 2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_RGBINIT = &H1
Private Const CC_ROUNDRECT = 256 '
Private Const CC_SHOWHELP = &H8
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_STYLED = 32
Private Const CC_WIDE = 16
Private Const CC_WIDESTYLED = 64
Private Const CCERR_CHOOSECOLORCODES = &H5000
Private Const LOGPIXELSY = 90
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2
Private Const NOERROR = 0
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_PROGRAMS = &H2
Private Const CSIDL_CONTROLS = &H3
Private Const CSIDL_PRINTERS = &H4
Private Const CSIDL_PERSONAL = &H5
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_STARTUP = &H7
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_SENDTO = &H9
Private Const CSIDL_BITBUCKET = &HA
Private Const CSIDL_STARTMENU = &HB
Private Const CSIDL_DESKTOPDIRECTORY = &H10
Private Const CSIDL_DRIVES = &H11
Private Const CSIDL_NETWORK = &H12
Private Const CSIDL_NETHOOD = &H13
Private Const CSIDL_FONTS = &H14
Private Const CSIDL_TEMPLATES = &H15
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_WININICHANGE = &H1A

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_USENEWUI = &H40
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000

' API TYPES
Private Type PRINTDLG_TYPE
  lStructSize As Long
  hwndOwner As Long
  hDevMode As Long
  hDevNames As Long
  hdc As Long
  Flags As Long
  nFromPage As Integer
  nToPage As Integer
  nMinPage As Integer
  nMaxPage As Integer
  nCopies As Integer
  hInstance As Long
  lCustData As Long
  lpfnPrintHook As Long
  lpfnSetupHook As Long
  lpPrintTemplateName As String
  lpSetupTemplateName As String
  hPrintTemplate As Long
  hSetupTemplate As Long
End Type
Private Type DEVNAMES_TYPE
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra As String * 100
End Type
Private Type DEVMODE_TYPE
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
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

' PUBLIC ENUM
Public Enum PrinterConstants
  cdlPDAllPages = &H0
  cdlPDCollate = &H10
  cdlPDDisablePrintToFile = &H80000
  cdlPDHelpButton = &H800
  cdlPDHidePrintToFile = &H100000
  cdlPDNoPageNums = &H8
  cdlPDNoSelection = &H4
  cdlPDNoWarning = &H80
  cdlPdPageNums = &H2
  cdlPDPrintSetup = &H40
  cdlPDPrintToFile = &H20
  cdlPDReturnDC = &H100
  cdlPDReturnDefault = &H400
  cdlPDReturnIC = &H200
  cdlPDSelection = &H1
  cdlPDUseDevModeCopies = &H40000
End Enum
Public Enum ErrorConstants
  cdlCancel = 32755
End Enum

' Variáveis
Private intMinPage As Integer
Private intMaxPage As Integer
Private intFromPage As Integer
Private intToPage As Integer


' Estruturas
Public Enum DLGmode
    Complete = 2
    Partial = 1
    FixedPartial = 0
End Enum

Public Enum wModeType   ' Modo de abertura do FileDialog
    OpenFile = 1
    SaveFile = 2
End Enum

Private Type OPENFILENAME
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type COLORSTRUC
    lStructSize As Long
    hwnd As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 31
End Type

Private Type FONTSTRUC
    lStructSize As Long
    hwnd As Long
    hdc As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type

Private Type DevMode
    dmDeviceName As String * 32
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
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFreq As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private ppp As ITEMIDLIST

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Enum OFNtype
    FILEDEFAULT = -1
    ALLOWMULTISELECT = &H200
    CREATEPROMPT = &H2000
    ENABLEHOOK = &H20
    ENABLETEMPLATE = &H40
    ENABLETEMPLATEHANDLE = &H80
    EXPLORER = &H80000
    EXTENSIONDIFFERENT = &H400
    FILEMUSTEXIST = &H1000
    HIDEREADONLY = &H4
    LONGNAMES = &H200000
    NOCHANGEDIR = &H8
    NODEREFERENCELINKS = &H100000
    NOLONGNAMES = &H40000
    NONETWORKBUTTON = &H20000
    NOREADONLYRETURN = &H8000
    NOTESTFILECREATE = &H10000
    NOVALIDATE = &H100
    OVERWRITEPROMPT = &H2
    PATHMUSTEXIST = &H800
    ReadOnly = &H1
    SHAREAWARE = &H4000
    SHAREFALLTHROUGH = 2
    SHARENOWARN = 1
    SHAREWARN = 0
    SHOWHELP = &H10
End Enum

Public Enum CFType
    ANSIONLY = &H400&
    APPLY = &H200&
    BITMAP = 2
    PRINTERFONTS = &H2
    PRIVATEFIRST = &H200
    PRIVATELAST = &H2FF
    RIFF = 11
    SCALABLEONLY = &H20000
    SCREENFONTS = &H1
    BOTH = (&H1 Or &H2)
    DIB = 8
    DIF = 5
    DSPBITMAP = &H82
    DSPENHMETAFILE = &H8E
    DSPMETAFILEPICT = &H83
    DSPTEXT = &H81
    EFFECTS = &H100&
    ENABLEHOOK = &H8&
    ENABLETEMPLATE = &H10&
    ENABLETEMPLATEHANDLE = &H20&
    ENHMETAFILE = 14
    FIXEDPITCHONLY = &H4000&
    FORCEFONTEXIST = &H10000
    GDIOBJFIRST = &H300
    GDIOBJLAST = &H3FF
    INITTOLOGFONTSTRUCT = &H40&
    LIMITSIZE = &H2000&
    METAFILEPICT = 3
    NOFACESEL = &H80000
    NOVERTFONTS = &H1000000
    NOVECTORFONTS = &H800&
    NOOEMFONTS = &H800&
    NOSCRIPTSEL = &H800000
    NOSIMULATIONS = &H1000&
    NOSIZESEL = &H200000
    NOSTYLESEL = &H100000
    OEMTEXT = 7
    OWNERDISPLAY = &H80
    Palette = 9
    PENDATA = 10
    SCRIPTSONLY = &H400&
    SELECTSCRIPT = &H400000
    SHOWHELP = &H4&
    SYLK = 4
    Text = 1
    TIFF = 6
    TTONLY = &H40000
    UNICODETEXT = 13
    USESTYLE = &H80&
    WAVE = 12
    WYSIWYG = &H8000
End Enum

Public Enum FormatDriveCapEnum
    SHFD_CAPACITY_DEFAULT = 0       ' default drive capacity
    SHFD_CAPACITY_360 = 3           ' 360KB, applies to 5.25" drives only
    SHFD_CAPACITY_720 = 5           ' 720KB, applies to 3.5" drives only
End Enum

Public Enum FormatDriveFormatTypeEnum
    SHFD_FORMAT_QUICK = 0           ' quick format
    SHFD_FORMAT_FULL = 1            ' full format
    SHFD_FORMAT_SYSONLY = 2         ' copies system files only (Win95 Only!)
End Enum


'// Declares privados do módulo
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As COLORSTRUC) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long



' Declares públicos do módulo
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal szFilename As String, ByVal dwCommand As Long, ByVal dwData As Any) As Long

' Enum de flags para o diálogo de impressora
Enum PrintFLAGS
    PD_NOSELECTION = &H4
    PD_DISABLEPRINTTOFILE = &H80000
    PD_PRINTTOFILE = &H20
    PD_RETURNDC = &H100
    PD_RETURNDEFAULT = &H400
    PD_RETURNIC = &H200
    PD_SELECTION = &H1
    PD_SHOWHELP = &H800
    PD_NOPAGENUMS = &H8
    PD_PAGENUMS = &H2
    PD_ALLPAGES = &H0
    PD_COLLATE = &H10
    PD_HIDEPRINTTOFILE = &H100000
End Enum


' Enum do HTMLHelp
Enum HTMLhelpFLAGS
    HH_DISPLAY_TOPIC = &H0&
    HH_HELP_FINDER = &H0&
    HH_DISPLAY_TOC = &H1&    '// Currently Not Implemented
    HH_DISPLAY_INDEX = &H2&  '// Currently Not Implemented
    HH_DISPLAY_SEARCH = &H3& '// Currently Not Implemented
    HH_SET_WIN_TYPE = &H4&
    HH_GET_WIN_TYPE = &H5&
    HH_GET_WIN_HANDLE = &H6&
    HH_ENUM_INFO_TYPE = &H7&
    HH_SET_INFO_TYPE = &H8&
    HH_SYNC = &H9&
    HH_ADD_NAV_UI = &H10&     '// Currently Not Implemented
    HH_ADD_BUTTON = &H11&     '// Currently Not Implemented
    HH_GETBROWSER_APP = &H12& '// Currently Not Implemented
    HH_KEYWORD_LOOKUP = &H13&
    HH_DISPLAY_TEXT_POPUP = &H14&
    HH_HELP_CONTEXT = &H15&
    HH_TP_HELP_CONTEXTMENU = &H16&
    HH_TP_HELP_WM_HELP = &H17&
    HH_CLOSE_ALL = &H18&
    HH_ALINK_LOOKUP = &H19&
    HH_GET_LAST_ERROR = &H20&       '// Currently Not Implemented
    HH_ENUM_CATEGORY = &H21&
    HH_ENUM_CATEGORY_IT = &H22&
    HH_RESET_IT_FILTER = &H23&
    HH_SET_INCLUSIVE_FILTER = &H24&
    HH_SET_EXCLUSIVE_FILTER = &H25&
    HH_SET_GUID = &H26&
    HH_INTERNAL = &H255&
End Enum

' Enum do BrowseForFolders
Enum BFType
    Pastas = 0
    Pastas_e_Arquivos = 1
    Computadores = 2
    Impressoras = 3
End Enum


Public Function SetDefaultPrinter(objPrn As Printer) As Boolean

    Dim X As Long, szTmp As String

    szTmp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    X = WriteProfileString("windows", "device", szTmp)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")
    SetDefaultPrinter = (X <> 0)
    
End Function

Public Function GetPrinterFromDevice(sDeviceName As String) As Object

    Dim PRT As Printer
    
    For Each PRT In Printers
        If PRT.DeviceName = sDeviceName Then
            Set GetPrinterFromDevice = PRT
            Exit For
        End If
    Next
    
End Function


' retorna o devicename da impressora default via API
' READ ONLY

Public Property Get GetDefaultPrinter()

    Dim X As Long, szTmp As String, dwBuf As Long

    dwBuf = 1024
    szTmp = Space(dwBuf + 1)
    X = GetProfileString("windows", "device", "", szTmp, dwBuf)
    GetDefaultPrinter = Trim(Left(szTmp, X))

End Property

' Reseta a impressora default para o devicename passado via API
Public Sub MakeDefaultPrinter(PrinterDeviceName As String)

    Dim X As Long

    X = WriteProfileString("windows", "device", PrinterDeviceName)
    X = SendMessageByString(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")

End Sub

' Escolha de pasta
Public Sub BrowseFolder(szDialogTitle As String, BrowseForWhat As BFType, Optional NewInterface As Boolean = False)

    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String
    Dim wPos As Integer, R As Long, PDL As ITEMIDLIST, Flags As Long
    
    Select Case BrowseForWhat
    Case 1  ' Arquivos e Pastas
        R = SHGetSpecialFolderLocation(mHWND, CSIDL_DESKTOP, PDL)
        Flags = BIF_RETURNONLYFSDIRS + BIF_BROWSEINCLUDEFILES + BIF_VALIDATE
    Case 2  ' Computadores
        R = SHGetSpecialFolderLocation(mHWND, CSIDL_NETWORK, PDL)
        Flags = BIF_BROWSEFORCOMPUTER
    Case 3  ' Impressoras
        R = SHGetSpecialFolderLocation(mHWND, CSIDL_PRINTERS, PDL)
        Flags = BIF_BROWSEFORPRINTER
    Case Else   ' Pastas
        R = SHGetSpecialFolderLocation(mHWND, CSIDL_DESKTOP, PDL)
        Flags = BIF_RETURNONLYFSDIRS + BIF_VALIDATE
    End Select
    
    If NewInterface = True Then Flags = Flags + BIF_USENEWUI + BIF_EDITBOX
    
    BI.hOwner = mHWND
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = Flags
    BI.pidlRoot = PDL.mkid.cb
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    If X Then
        wPos = InStr(szPath, Chr(0))
        mPATH = Left$(szPath, wPos - 1)
        mCANCEL = False
    Else
        '' Cancela foi pressionado
        mPATH = ""
        mCANCEL = True
    End If

End Sub

' Permite conectar a impressoras locais ou na rede
Public Sub DialogConnectToPrinter()

    Shell "rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", vbNormalFocus

End Sub

Private Function ByteToString(aBytes() As Byte) As String

    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String

    dwBytePoint = LBound(aBytes)

    While dwBytePoint <= UBound(aBytes)

        dwByteVal = aBytes(dwBytePoint)

        If dwByteVal = 0 Then
            ByteToString = szOut
            Exit Function
        Else
            szOut = szOut & Chr$(dwByteVal)
        End If

        dwBytePoint = dwBytePoint + 1

    Wend

    ByteToString = szOut

End Function

' Diálogo de Cor
Public Sub DialogColor(Modo As DLGmode)

    Dim X As Long, CS As COLORSTRUC, CustColor(16) As Long, CSTNTs As Long
    
    Select Case Modo
    Case 0
        CSTNTs = CC_PREVENTFULLOPEN
    Case 1
        CSTNTs = 0
    Case Else
        CSTNTs = 3
        
    End Select
    
    CS.lStructSize = Len(CS)
    CS.hwnd = mHWND
    CS.hInstance = App.hInstance
    CS.Flags = CC_SOLIDCOLOR + CSTNTs + CC_SHOWHELP
    CS.lpCustColors = VarPtr(aCustColor(0))
    X = ChooseColor(CS)
    If X = 0 Then
        mCANCEL = True
    Else
        mCOLOR = CS.rgbResult
        mCANCEL = False
    End If

End Sub

' Dialogo de Abrir / Salvar Arquivo
Public Sub DialogFile(wMode As wModeType, Optional szDialogTitle As String = "")

    Dim X As Long, OFN As OPENFILENAME, szFile As String
    Dim szFileTitle As String, FilterBUF As String, I As Integer
    Dim A As Integer, FL As OFNtype, TMP As Variant, Findexes As Integer

    FilterBUF = mFILTER
    Findexes = 0
    Do
        X = InStr(1, FilterBUF, "|")
        If X Then Mid$(FilterBUF, X, 1) = Chr(0)
        Findexes = Findexes + 1
    Loop Until X = 0

    If szDialogTitle = "" Then
        If wMode = OpenFile Then
            szDialogTitle = "Abrir Arquivo"
        Else
            szDialogTitle = "Salvar Arquivo"
        End If
    End If
    
    OFN.lStructSize = Len(OFN)
    OFN.hwnd = mHWND
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = String$(250, 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = FilterBUF
    If Findexes < mFilterIndex Then mFilterIndex = 1
    OFN.nFilterIndex = mFilterIndex
    OFN.lpstrInitialDir = mPATH
    OFN.lpstrDefExt = mDEFEXT

    If wMode = 1 Then
        If mFlags = 0 Then
            FL = PATHMUSTEXIST Or FILEMUSTEXIST Or HIDEREADONLY
        Else
            FL = mFlags
        End If
        OFN.Flags = FL
        X = GetOpenFileName(OFN)
    Else
        If mFlags = 0 Then
            FL = OVERWRITEPROMPT Or PATHMUSTEXIST Or HIDEREADONLY
        Else
            FL = mFlags
        End If
        OFN.Flags = FL
        X = GetSaveFileName(OFN)
    End If
    
    If X <> 0 Then
        ' remove os char 0s extras
        TMP = Replace(OFN.lpstrFile, Chr(0) & Chr(0), "")
        If Right(TMP, 1) = Chr(0) Then TMP = Left(TMP, Len(TMP) - 1)
        ' Divide num array para obter path e filename
        TMP = Split(TMP, Chr(0))
        mFilterIndex = OFN.nFilterIndex
        ' Verifica se foram escolhidos 1 ou mais arquivos
        If UBound(TMP) = 0 Then
            mMultipleFiles = False
            mFULLPATH = TMP(0)
            mFILE = Replace(OFN.lpstrFileTitle, Chr(0), "")
            A = InStrRev(TMP(0), "\")
            If A Then
                mPATH = Left(TMP(0), A)
            Else
                mPATH = TMP(0)
            End If
        Else
            mMultipleFiles = True
            mFULLPATH = TMP(0) & TMP(1)
            mPATH = TMP(0)
            mFILE = ""
            For I = 1 To UBound(TMP)
                mFILE = mFILE & TMP(I) & "|"
            Next
            mFILE = Left(mFILE, Len(mFILE) - 1)
        End If
        mDEFEXT = OFN.lpstrDefExt
        mCANCEL = False
    Else
        mFULLPATH = ""
        mCANCEL = True
    End If

End Sub

' Dialogo de Fonte
Public Sub DialogFont()

    ' O tamanho não está entrando direito

    Dim cf As FONTSTRUC
    Dim lf As LOGFONT
    Dim I As Integer
    Dim hMem As Long, pMem As Long
    Dim retVal As Long
    
    cf.hwnd = mHWND

    lf.lfWeight = 400 + Abs((mBOLD = True)) * 500
    lf.lfFaceName = mFONT & vbNullChar
    lf.lfHeight = -(mSIZE * (20 / Screen.TwipsPerPixelY))
    lf.lfWidth = 0
    lf.lfCharSet = 1
    lf.lfEscapement = 0         ' angle between baseline and escapement vector
    lf.lfOrientation = 0        ' angle between baseline and orientation vector
    lf.lfOutPrecision = 0       ' default precision mapping
    lf.lfClipPrecision = 0      ' default clipping precision
    lf.lfQuality = 0            ' default quality setting
    lf.lfPitchAndFamily = 16    ' default pitch, proportional with serifs

    cf.rgbColors = 0
    cf.Flags = mFontFlags + &H40&
    cf.iPointSize = mSIZE * 10
    cf.hdc = Printer.hdc
    cf.hwnd = mHWND
    
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lf))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lf, Len(lf)
    
    cf.lpLogFont = pMem
    cf.lStructSize = Len(cf)
    
    If ChooseFont(cf) <> 1 Then
        mCANCEL = True
        Exit Sub
    End If
    mCANCEL = False

    CopyMemory lf, ByVal pMem, Len(lf)

    mSIZE = cf.iPointSize / 10
    mITALIC = lf.lfItalic
    mBOLD = IIf(lf.lfWeight > 400, 1, 0)

    mFONT = Left(lf.lfFaceName, InStr(1, lf.lfFaceName, Chr(0)) - 1)

    mCOLOR = cf.rgbColors
    mStrikeThru = lf.lfStrikeOut
    mUnderline = lf.lfUnderline
    
    retVal = GlobalUnlock(hMem)
    retVal = GlobalFree(hMem)

End Sub

' Dialogo de Impressão
Public Sub DialogPrintSetup()

    Dim X As Long, PD As PRINTDLG_TYPE

    If mHWND = 0 Then Exit Sub

    PD.lStructSize = Len(PD)
    PD.hwndOwner = mHWND
    PD.Flags = PD_PRINTSETUP
    X = PrintDialog(PD)

End Sub

' Flags do diálogo Abrir / Salvar
Public Property Let FileFlags(ByVal FFlgs As OFNtype)

    mFlags = FFlgs

End Property

Public Property Get FileFlags() As OFNtype

    FileFlags = mFlags

End Property

' Flags do diálogo de Fontes
Public Property Let FontFlags(ByVal FFlgs As CFType)

    mFontFlags = FFlgs

End Property

Public Property Get FontFlags() As CFType

    FontFlags = mFontFlags

End Property

' Flags do diálogo impressão
Public Property Let PrinterFlags(ByVal FFlgs As PrinterConstants)

    mPRTFlags = FFlgs

End Property

Public Property Get PrinterFlags() As PrinterConstants

    PrinterFlags = mPRTFlags

End Property

' Nome do arquivo (Abrir / Salvar)
Public Property Let FileName(ByVal fln As String)

    mFILE = fln

End Property

Public Property Get FileName() As String

    FileName = mFILE

End Property

' Caminho (Abrir / Salvar, BrowseForFolder)
Public Property Let Path(ByVal PTH As String)

    mPATH = PTH

End Property

Public Property Get Path() As String

    Path = mPATH

End Property

' Caminho Completo (Abrir / Salvar)
Public Property Get FullPath() As String
    '' Read only
    FullPath = mFULLPATH

End Property

' Cores do Sistema
Public Property Get CustomColor(ByVal Ind As Integer) As Long
    
    If Ind >= 0 And Ind <= 15 Then
        CustomColor = aCustColor(Ind)
    Else
        CustomColor = -1
    End If
    
End Property

Public Property Let CustomColor(ByVal Ind As Integer, COR As Long)
    
    If Ind >= 0 And Ind <= 15 Then
        aCustColor(Ind) = COR
    End If
    
End Property

' Cor (Cor, Fonte)
Public Property Let Color(ByVal COR As Long)

    mCOLOR = COR

End Property

Public Property Get Color() As Long

    Color = mCOLOR

End Property

' Nome da Fonte (Fonte)
Public Property Let Font(xxx As String)

    mFONT = xxx

End Property

Public Property Get Font() As String
    '' Read only
    Font = mFONT

End Property

' tamanho da Fonte (Fonte)
Public Property Let FontSize(ByVal xxx As Integer)

    mSIZE = xxx

End Property

Public Property Get FontSize() As Integer
    
    FontSize = mSIZE

End Property

' Negrito (fonte)
Public Property Let FontBold(ByVal bld As Boolean)

    mBOLD = bld

End Property

Public Property Get FontBold() As Boolean

    FontBold = mBOLD

End Property

' Italico  (fonte)
Public Property Let FontItalic(ByVal ita As Boolean)

    mITALIC = ita

End Property

Public Property Get FontItalic() As Boolean

    FontItalic = mITALIC

End Property

' Sublinhado (fonte) Read Only
Public Property Get FontUnderline() As Boolean

    FontUnderline = mUnderline

End Property

' Riscado (fonte) Read Only
Public Property Get FontStrikeThru() As Boolean

    FontStrikeThru = mStrikeThru

End Property

' Handle do form pai
Public Property Let hwnd(ByVal hw As Long)

    mHWND = hw

End Property

Public Property Get hwnd() As Long

    hwnd = mHWND

End Property

' Filtro de extensões (Abrir / Salvar) Read Only
Public Property Let Filter(fil As String)
    
    mFILTER = fil

End Property

' Extensão Default (Abrir / Salvar)
Public Property Let DefaultEXT(ByVal xxx As String)

    mDEFEXT = xxx

End Property

Public Property Get DefaultEXT() As String

    DefaultEXT = mDEFEXT

End Property

' Indice do Filtro a ser feito default qdo o diálogo abrir (Abrir / Salvar)
Public Property Let FilterIndex(ByVal xxx As String)

    mFilterIndex = xxx

End Property

Public Property Get FilterIndex() As String

    FilterIndex = mFilterIndex

End Property

' Fica True quando cancel for pressionado em algum diálogo (Varios)
Public Property Get Cancel()

    Cancel = mCANCEL

End Property

' retorna se foram escolhidos mais de um arquivo no diálogo de abrir arquivo (read Only)
Public Property Get MultipleFiles() As Boolean

    MultipleFiles = mMultipleFiles
    
End Property

Private Sub Class_Initialize()

    Dim I As Integer
    
    mFILE = ""
    mPATH = "c:\"
    mFULLPATH = "c:\"
    mCOLOR = 0
    mBOLD = False
    mITALIC = False
    mSIZE = 8
    mFONT = "Arial"
    mHWND = 0
    mCANCEL = False
    mFILTER = ""
    mFromPAGE = 1
    mToPAGE = 1
    mINITDIR = ""
    mDEFEXT = ""
    mFlags = 0
    mPRTFlags = cdlPDAllPages
    intMinPage = 0
    intMaxPage = 0
    intFromPage = 0
    intToPage = 0
    mFontFlags = BOTH
    mMultipleFiles = False
    
    ' Inicializa as cores customizáveis
    For I = 0 To 15
        aCustColor(I) = GetSysColor(I)
    Next
    
End Sub

' Diálogo de Formatação de disquetes
Public Sub FormatFloppy(ByVal DriveLetter As String, Capacity As FormatDriveCapEnum, Format As FormatDriveFormatTypeEnum)
    
    ' Formata um disquete
    
    Dim DRNum As Long
    
    DRNum = Asc(Left(DriveLetter, 1)) - 65
    SHFormatDrive Me.hwnd, DRNum, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_QUICK

End Sub

' Diálogo de conexão de drives de rede (compartilhamentos)
Public Sub ConnectNetDrive()

    Dim Resp As Long
    
    Resp = WNetConnectionDialog(mHWND, RESOURCETYPE_DISK)

End Sub

' Diálogo de desconexão de drives de rede (compartilhamentos)
Public Sub DisconnectNetDrive()

    Dim Resp As Long
    
    Resp = WNetDisconnectDialog(mHWND, RESOURCETYPE_DISK)

End Sub

' Copia Disquetes
Public Sub CopyDisk()

    Shell "rundll32 diskcopy,DiskCopyRunDll", vbNormalFocus

End Sub

' Menor Página (Impressão)
Property Get Min() As Integer

  Min = intMinPage
  
End Property

Property Let Min(ByVal intNewValue As Integer)

  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMinPage = intNewValue
  If intNewValue > intFromPage Then intFromPage = intNewValue
  If intNewValue > intToPage Then intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
  
End Property

' Página Inicial (Impressão)
Property Get FromPage() As Integer

  FromPage = intFromPage
  
End Property

Property Let FromPage(ByVal intNewValue As Integer)

  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intFromPage = intNewValue
  If intNewValue > intToPage Then intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
  
End Property

 ' Página Final (Impressão)
Property Get ToPage() As Integer
  
  ToPage = intToPage

End Property

Property Let ToPage(ByVal intNewValue As Integer)

  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intToPage = intNewValue
  If intNewValue > intMaxPage Then intMaxPage = intNewValue
  If intNewValue < intFromPage Then intFromPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
  
End Property

' Maior Página  (Impressão)
Property Get Max() As Integer

  Max = intMaxPage
  
End Property

Property Let Max(ByVal intNewValue As Integer)

  intNewValue = IIf(intNewValue > 0, intNewValue, 0)
  intMaxPage = intNewValue
  If intNewValue < intToPage Then intToPage = intNewValue
  If intNewValue < intFromPage Then intFromPage = intNewValue
  If intNewValue < intMinPage Then intMinPage = intNewValue
  
End Property

' Diálogo de Impressão
Public Sub DialogPrint()

Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE
Dim lpDevMode As Long, lpDevName As Long
Dim intReturn As Integer
Dim objPrinter As Printer
Dim strNewPrinterName As String

  mCANCEL = False
  
  With PrintDlg
    .lStructSize = Len(PrintDlg)
    .hwndOwner = mHWND
    .Flags = mPRTFlags
    .nMinPage = intMinPage
    .nFromPage = intFromPage
    .nToPage = intToPage
    .nMaxPage = intMaxPage
  End With
  
  'Setar orientação e duplex
  DevMode.dmDeviceName = Printer.DeviceName
  DevMode.dmSize = Len(DevMode)
  DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
  DevMode.dmOrientation = Printer.Orientation
  
  On Error Resume Next
  DevMode.dmDuplex = Printer.Duplex
  On Error GoTo 0
  
  PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
  lpDevMode = GlobalLock(PrintDlg.hDevMode)
  If lpDevMode > 0 Then
    CopyMemory ByVal lpDevMode, DevMode, Len(DevMode)
    intReturn = GlobalUnlock(lpDevMode)
  End If
  
  With DevName
    .wDriverOffset = 8
    .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
    .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
    .wDefault = 0
  End With
  
  With Printer
    DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
  End With
  
  PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
  lpDevName = GlobalLock(PrintDlg.hDevNames)
  If lpDevName > 0 Then
    CopyMemory ByVal lpDevName, DevName, Len(DevName)
    intReturn = GlobalUnlock(lpDevName)
  End If
  
  ' Chama diálogo
  If PrintDialog(PrintDlg) Then
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    CopyMemory DevName, ByVal lpDevName, 45
    intReturn = GlobalUnlock(lpDevName)
    
    With PrintDlg
      mPRTFlags = .Flags
      intFromPage = .nFromPage
      intToPage = .nToPage
    End With
     GlobalFree PrintDlg.hDevNames
    
    ' Marca as propriedades retornadas pela estrutura
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    CopyMemory DevMode, ByVal lpDevMode, Len(DevMode)
    intReturn = GlobalUnlock(PrintDlg.hDevMode)
    GlobalFree PrintDlg.hDevMode
    strNewPrinterName = UCase$(Left(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, Chr$(0)) - 1))
    
    If Printer.DeviceName <> strNewPrinterName Then
      For Each objPrinter In Printers
        If UCase$(objPrinter.DeviceName) = strNewPrinterName Then Set Printer = objPrinter
      Next
    End If
    
    On Error Resume Next
    
    With Printer
      .Copies = DevMode.dmCopies
      .Duplex = DevMode.dmDuplex
      .Orientation = DevMode.dmOrientation
    End With
    On Error GoTo 0
    mCANCEL = False
  Else
    GlobalFree PrintDlg.hDevMode
    GlobalFree PrintDlg.hDevNames
    mCANCEL = True
    End If
    
End Sub

' Carrega um listbox ou combobox com a lista de impressoras disponíveis na coleção printers
Public Function LoadPrinters(LST As Control, Optional MarkDefault As Boolean = False) As Integer

    Dim PRT As Printer, BUF As String, I As Integer
    
    LST.Clear
    For Each PRT In Printers
        BUF = PRT.DeviceName
        If Printer.DeviceName = BUF Then
            LoadPrinters = I
            If MarkDefault = True Then
                BUF = BUF & " (Default)"
            End If
        End If
        LST.AddItem BUF
    Next
    
End Function

' Mostra HTMLHelp
Public Function ShowHTMLHelp(strFilename As String) As Long
    
    Dim hwndHelp As Long
    
    'The return value is the window handle of the created help window.
    ShowHTMLHelp = HtmlHelp(mHWND, strFilename, HH_DISPLAY_TOPIC, 0&)
    
End Function

