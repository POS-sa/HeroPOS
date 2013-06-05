Attribute VB_Name = "SRP770_API"
' Return Types
Public Const SEM_SUCCESS As Long = 0
Public Const SEM_ERR_NOPRINTER As Long = -10    'Specified printer driver does not exist
Public Const SEM_ERR_NOTSUPPOER As Long = -20   'Specified printer or port are not supported
Public Const SEM_ERR_OPEN As Long = -30         'Cannot open printer port
Public Const SEM_ERR_WRITE As Long = -40        'Write Error
Public Const SEM_ERR_READ As Long = -50         'Read Error
Public Const SEM_ERR_TIMEOUT As Long = -60      'Timeout Error
Public Const SEM_ERR_PARAM As Long = -70        'Function Parameter Error

' DLL Function prototypes
Declare Function DirectWrite Lib "SRP770_API.dll" _
    (ByVal pazPrinterName As String, ByRef byWrite As Byte, ByVal dwWrite As Long) As Long



