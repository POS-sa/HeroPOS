Attribute VB_Name = "modCryptoAPI"
Option Explicit
'dss
Public sGetKeyName As String

'
' Enumeration for error codes.
'
Public Enum eERROR_CODE
    eProviderUnavailable = 101
    eGeneratingKeyPair = 102
    eExportingPublicKey = 103
    eExportingPrivateKey = 104
    eGeneratingHash = 105
    eGeneratingSignature = 106
    eWritingSignatureFile = 107
    eImportingPublicKey = 108
    eImportingPrivateKey = 109
    eReadingSignatureFile = 110
    eKeyAlreadyExists = 111
End Enum
'
' Key constants
'
Public Const RSA1 As Long = &H31415352
Public Const RSA2 As Long = &H32415352
'
' Algorithm id constants
'
Public Const ALG_CLASS_HASH As Long = (4 * 2 ^ 13)
Public Const ALG_SID_MD5 As Long = 3
Public Const CALG_MD5 As Long = ALG_CLASS_HASH + ALG_SID_MD5
'
' Acquire constants
'
Public Const PROV_RSA_FULL As Long = 1
Public Const MS_DEF_PROV = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_ENHANCED_PROV = "Microsoft Enhanced Cryptographic Provider v1.0"
'
' dwFlags definitions for CryptAcquireContext
'
Public Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Public Const CRYPT_NEWKEYSET As Long = &H8
Public Const CRYPT_DELETEKEYSET As Long = &H10
Public Const CRYPT_MACHINE_KEYSET As Long = &H20
'
' dwFlag definitions for CryptGenKey
'
Public Const CRYPT_EXPORTABLE As Long = &H1
Public Const CRYPT_USER_PROTECTED As Long = &H2
Public Const CRYPT_CREATE_SALT As Long = &H4
Public Const CRYPT_UPDATE_KEY As Long = &H8
Public Const CRYPT_NO_SALT As Long = &H10
Public Const CRYPT_PREGEN As Long = &H40
Public Const CRYPT_RECIPIENT As Long = &H10
Public Const CRYPT_INITIATOR As Long = &H40
Public Const CRYPT_ONLINE As Long = &H80
Public Const CRYPT_SF As Long = &H100
Public Const CRYPT_CREATE_IV As Long = &H200
Public Const CRYPT_KEK As Long = &H400
Public Const CRYPT_DATA_KEY As Long = &H800
'
' CryptSetProvParam
'
Public Const PP_CLIENT_HWND As Long = 1
Public Const PP_CONTEXT_INFO As Long = 11
Public Const PP_KEYEXCHANGE_KEYSIZE As Long = 12
Public Const PP_SIGNATURE_KEYSIZE As Long = 13
Public Const PP_KEYEXCHANGE_ALG As Long = 14
Public Const PP_SIGNATURE_ALG As Long = 15
Public Const PP_DELETEKEY As Long = 24
'
' exported key blob definitions
'
Public Const SIMPLEBLOB As Long = &H1
Public Const PUBLICKEYBLOB As Long = &H6
Public Const PRIVATEKEYBLOB As Long = &H7
Public Const PLAINTEXTKEYBLOB As Long = &H8
'
' nte errors
'
Public Const NTE_BAD_UID As Long = &H80090001             ' Bad UID
Public Const NTE_BAD_HASH As Long = &H80090002            ' Bad Hash
Public Const NTE_BAD_KEY As Long = &H80090003             ' Bad Key
Public Const NTE_BAD_LEN As Long = &H80090004             ' Bad Length
Public Const NTE_BAD_DATA As Long = &H80090005            ' Bad Data
Public Const NTE_BAD_SIGNATURE As Long = &H80090006       ' Bad Signature
Public Const NTE_BAD_VER As Long = &H80090007             ' Bad Version of provider
Public Const NTE_BAD_ALGID As Long = &H80090008           ' Invalid algorithm specified
Public Const NTE_BAD_FLAGS As Long = &H80090009           ' Invalid flags specified
Public Const NTE_BAD_TYPE As Long = &H8009000A            ' Invalid type specified
Public Const NTE_BAD_KEY_STATE As Long = &H8009000B       ' Key not valid for use in specified state
Public Const NTE_BAD_HASH_STATE As Long = &H8009000C      ' Hash not valid for use in specified state
Public Const NTE_NO_KEY As Long = &H8009000D              ' Key does not exist
Public Const NTE_NO_MEMORY As Long = &H8009000E           ' Insufficient memory available for the operation
Public Const NTE_EXISTS As Long = &H8009000F              ' Object already exists
Public Const NTE_PERM As Long = &H80090010                ' Access denied
Public Const NTE_NOT_FOUND As Long = &H80090011           ' Object was not found
Public Const NTE_DOUBLE_ENCRYPT As Long = &H80090012      ' Data already encrypted
Public Const NTE_BAD_PROVIDER As Long = &H80090013        ' Invalid provider specified
Public Const NTE_BAD_PROV_TYPE As Long = &H80090014       ' Invalid provider type specified
Public Const NTE_BAD_PUBLIC_KEY As Long = &H80090015      ' Provider's public key is invalid
Public Const NTE_BAD_KEYSET As Long = &H80090016          ' Keyset does not exist
Public Const NTE_PROV_TYPE_NOT_DEF As Long = &H80090017   ' Provider type not defined
Public Const NTE_PROV_TYPE_ENTRY_BAD As Long = &H80090018 ' Provider type as registered is invalid
Public Const NTE_KEYSET_NOT_DEF As Long = &H80090019      ' The keyset is not defined
Public Const NTE_KEYSET_ENTRY_BAD As Long = &H8009001A    ' Keyset as registered is invalid
Public Const NTE_PROV_TYPE_NO_MATCH As Long = &H8009001B  ' Provider type does not match registered value
Public Const NTE_SIGNATURE_FILE_BAD As Long = &H8009001C  ' The digital signature file is corrupt
Public Const NTE_PROVIDER_DLL_FAIL As Long = &H8009001D   ' Provider DLL failed to initialize correctly
Public Const NTE_PROV_DLL_NOT_FOUND As Long = &H8009001E  ' Provider DLL could not be found
Public Const NTE_BAD_KEYSET_PARAM As Long = &H8009001F    ' The Keyset parameter is invalid
Public Const NTE_FAIL As Long = &H80090020                ' An internal error occurred
Public Const NTE_SYS_ERR As Long = &H80090021             ' A base error occurred
'
' Generate constants
'
Public Const AT_SIGNATURE As Long = 2
'
' Length of types in bytes (for lset)
'
Public Const T_PUBLICKEYBLOBLEN = 84
Public Const T_PRIVATEKEYBLOBLEN = 308

Public Type T_EXP_PUBLICKEYBLOB
    bPublicKey(1 To T_PUBLICKEYBLOBLEN) As Byte
End Type

Public Type T_EXP_PRIVATEKEYBLOB
    bPrivateKey(1 To T_PRIVATEKEYBLOBLEN) As Byte
End Type

Public Type T_PUBLICKEYBLOB
    bType    As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 64) As Byte
End Type

Public Type T_PRIVATEKEYBLOB
    bType    As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
    magic    As Long
    bitlen   As Long
    pubexp   As Long
    modulus(1 To 64)  As Byte
    prime1(1 To 32)   As Byte
    prime2(1 To 32)   As Byte
    exponent1(1 To 32)       As Byte
    exponent2(1 To 32)       As Byte
    coefficient(1 To 32)     As Byte
    privateExponent(1 To 64) As Byte
End Type

Public Declare Function CryptAcquireContext Lib "advapi32.dll" _
    Alias "CryptAcquireContextA" (ByRef hCryptProv As Long, _
    ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    
Public Declare Function CryptReleaseContext Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Public Declare Function CryptGenKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long
     
Public Declare Function CryptExportKey Lib "advapi32.dll" _
    (ByVal hKey As Long, ByVal hExpKey As Long, ByVal dwBlobType As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, _
    ByRef pdwDataLen As Long) As Long
     
Public Declare Function CryptCreateHash Lib "advapi32.dll" _
    (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
    ByVal dwFlags As Long, ByRef phHash As Long) As Long
    
Public Declare Function CryptHashData Lib "advapi32.dll" _
    (ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal dwFlags As Long) As Long
    
Public Declare Function CryptSignHash Lib "advapi32.dll" _
    Alias "CryptSignHashA" (ByVal hHash As Long, ByVal dwKeySpec As Long, _
    ByVal sDescription As String, ByVal dwFlags As Long, _
    ByRef pbSignature As Any, ByRef pdwSigLen As Long) As Long
      
Public Declare Function CryptVerifySignature Lib "advapi32.dll" _
    Alias "CryptVerifySignatureA" (ByVal hHash As Long, _
    ByRef pbSignature As Any, ByVal dwSigLen As Long, _
    ByVal hPubKey As Long, ByVal sDescription As String, _
    ByVal dwFlags As Long) As Long

Public Declare Function CryptDestroyHash Lib "advapi32.dll" _
    (ByVal hHash As Long) As Long
   
Public Declare Function CryptImportKey Lib "advapi32.dll" _
    (ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, _
    ByRef phKey As Long) As Long

Public Declare Function CryptDestroyKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Private Function fGetErrorMsg(code As eERROR_CODE) As String
'
' Decode the selected error code and return a message.
'
Select Case code
    Case eProviderUnavailable
        fGetErrorMsg = "Error accessing security module"
    Case eGeneratingKeyPair
        fGetErrorMsg = "Error generating key pair"
    Case eExportingPublicKey
        fGetErrorMsg = "Error exporting public key"
    Case eExportingPrivateKey
        fGetErrorMsg = "Error exporting private key"
    Case eGeneratingHash
        fGetErrorMsg = "Error generating hash"
    Case eGeneratingSignature
        fGetErrorMsg = "Error signing file"
    Case eWritingSignatureFile
        fGetErrorMsg = "Error writing signature file"
    Case eImportingPublicKey
        fGetErrorMsg = "Error importing public key"
    Case eImportingPrivateKey
        fGetErrorMsg = "Error importing private key"
    Case eReadingSignatureFile
        fGetErrorMsg = "Error reading signature file"
    Case eKeyAlreadyExists
        fGetErrorMsg = "Error creating key, key already exists"
End Select
End Function


Public Function fInsertArg(sText As String, lPosition As Long, _
    vArgument As Variant) As String
'
' Insert a string into another string.
'
fInsertArg = Left$(sText, lPosition - 1) & vArgument & Mid$(sText, lPosition + 2)
End Function

Public Function fGetErrorString(code As eERROR_CODE, ParamArray Arguments() As Variant) As String
Dim sError     As String
Dim lArgIndex  As Long
Dim lStrArgPos As Long
Dim lNumArgPos As Long

sError = fGetErrorMsg(code)
lArgIndex = 0

While InStr(sError, "%s") Or InStr(sError, "%n")
    lStrArgPos = InStr(sError, "%s")
    lNumArgPos = InStr(sError, "%n")
    
    If lStrArgPos <> 0 And lNumArgPos <> 0 Then
        If lStrArgPos < lNumArgPos Then
            sError = fInsertArg(sError, lStrArgPos, Arguments(lArgIndex))
            sError = fInsertArg(sError, lNumArgPos, Arguments(lArgIndex + 1))
        Else
            sError = fInsertArg(sError, lNumArgPos, Arguments(lArgIndex))
            sError = fInsertArg(sError, lStrArgPos, Arguments(lArgIndex + 1))
        End If
    ElseIf lStrArgPos <> 0 Then
        sError = fInsertArg(sError, lStrArgPos, Arguments(lArgIndex))
        lArgIndex = lArgIndex + 1
    ElseIf lNumArgPos <> 0 Then
        sError = fInsertArg(sError, lNumArgPos, Arguments(lArgIndex))
        lArgIndex = lArgIndex + 1
    End If
Wend

fGetErrorString = sError
End Function

'Public Sub pRaiseError(sSource As String, sRoutine As String, _
'    code As eERROR_CODE, Optional ByVal sArgument As String = "")
''
'' Raise an error.
''
'err.Raise vbObjectError + code, sSource & "." & sRoutine, _
'        fGetErrorMsg(code) & " " & sArgument
'End Sub

