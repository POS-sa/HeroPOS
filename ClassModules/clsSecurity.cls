VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
' Digital signing and encryption functions.
'
Const Source = "clsSecurity"
'
' Chunk size to use to hash files.
'
Private Const C_CHUNK_SIZE As Long = 512

Public Enum E_KEY_TYPE
    eKeyNone = 0
    eKeyPublic = 1
    eKeyPrivate = 2
End Enum

Private m_hProvider  As Long        ' handle to crypto provider
Private m_hKeyPair   As Long        ' handle to key pair
Private m_eKeyStatus As E_KEY_TYPE  ' flag indicating current key type

Public Sub CreateKeyPair(KeyName As String)
Dim lResult    As Long
Dim PublicKey  As T_EXP_PUBLICKEYBLOB
Dim PrivateKey As T_EXP_PRIVATEKEYBLOB
'
' Creates a new key pair and writes out
' a file for each part called "KeyName.opgp"
' and "KeyName.opkp" respectively in the
' application path
'
Const Routine = "CreateKeyPair"
'
' See if either key already exists.
'
If Dir(App.Path & "\" & KeyName & ".opkp") <> "" And _
   Dir(App.Path & "\" & KeyName & ".opgp") <> "" Then
    'Call pRaiseError(Source, Routine, eKeyAlreadyExists)
End If
'
' See if provider is obtained.
'
If Not (IsProviderOpen) Then
    '
    ' Get the provider.
    '
    If Not (AcquireCryptoProvider) Then
        'Call pRaiseError(Source, Routine, eProviderUnavailable)
        Exit Sub
    End If
End If
'
' Create the new key pair.
' The use of &H2000000 guarantees a 512 bit key.
'
lResult = CryptGenKey(m_hProvider, _
            AT_SIGNATURE, &H2000000 Or CRYPT_EXPORTABLE, m_hKeyPair)
                      
If lResult = 0 Then
    m_eKeyStatus = eKeyNone
    'Call pRaiseError(Source, Routine, eGeneratingKeyPair)
    Exit Sub
Else
    '
    ' New stored key type.
    '
    m_eKeyStatus = eKeyPrivate
End If
'
' Copy the keys to variables.
'
If Not (ExportPublicKey(PublicKey)) Then
    'Call pRaiseError(Source, Routine, eExportingPublicKey)
    Exit Sub
End If
If Not (ExportPrivateKey(PrivateKey)) Then
    'Call pRaiseError(Source, Routine, eExportingPrivateKey)
    Exit Sub
End If
'
' Create the key files.
If Not (WriteFile(App.Path & "\" & KeyName & ".opgp", PublicKey.bPublicKey)) Then
    'Call pRaiseError(Source, Routine, eExportingPublicKey)
End If
If Not (WriteFile(App.Path & "\" & KeyName & ".opkp", PrivateKey.bPrivateKey)) Then
    'Call pRaiseError(Source, Routine, eExportingPrivateKey)
End If
'
' Destroy the key.
'
CryptDestroyKey m_hKeyPair

m_eKeyStatus = eKeyNone
End Sub

Public Sub SignFile(sKeyName As String, sFileName As String)
Dim lResult       As Long
Dim hHash         As Long
Dim lSigLen       As Long
Dim bSig(1 To 64) As Byte
Dim PrivateKey    As T_EXP_PRIVATEKEYBLOB
'
' Sign a file with the specified private key.
'
Const Routine = "SignFile"
'
' See if the provider is obtained.
'
If Not (IsProviderOpen) Then
    '
    ' Get the provider.
    '
    If Not (AcquireCryptoProvider) Then
'        Call pRaiseError(Source, Routine, eProviderUnavailable)
        'Exit Sub
    End If
End If
'
' Read the private key file.
'
If Not (ReadFile(App.Path & "\" & sKeyName & ".opkp", PrivateKey.bPrivateKey)) Then
'    Call pRaiseError(Source, Routine, eImportingPrivateKey)
    'Exit Sub
End If
'
' Import the key blob into the key container.
'
If Not (ImportPrivateKey(PrivateKey)) Then
'    Call pRaiseError(Source, Routine, eImportingPrivateKey)
    'Exit Sub
End If
'
' Hash the file.
'
If Not (HashFile(sFileName, hHash)) Then
'    Call pRaiseError(Source, Routine, eGeneratingHash)
    'Exit Sub
End If
'
' Create the signature.
'
lSigLen = 64
lResult = CryptSignHash(hHash, AT_SIGNATURE, vbNullString, 0, bSig(1), lSigLen)
If lResult = 0 Then
'    Call pRaiseError(Source, Routine, eGeneratingSignature)
    Exit Sub
End If
'
' Destroy the hash.
'
lResult = CryptDestroyHash(hHash)
'
' Write the signature file.
If Not (WriteFile(sFileName & ".sgn", bSig)) Then
'    Call pRaiseError(Source, Routine, eWritingSignatureFile)
    Exit Sub
End If
End Sub

Public Function VerifyFile(sKeyName As String, sFileName As String) As Boolean
Dim lResult       As Long
Dim lSigLen       As Long
Dim hHash         As Long
Dim bSig(1 To 64) As Byte
Dim PublicKey     As T_EXP_PUBLICKEYBLOB
'
' Verify a file with the specified public key.
'
Const Routine = "VerifyFile"
'
' See if the provider is obtained.
'
If Not (IsProviderOpen) Then
    '
    ' Get the provider.
    '
    If Not (AcquireCryptoProvider) Then
'        Call pRaiseError(Source, Routine, eProviderUnavailable)
        Exit Function
    End If
End If
'
' Read the public key file.
'
If Not (ReadFile(App.Path & "\" & sKeyName & ".opgp", PublicKey.bPublicKey)) Then
'    Call pRaiseError(Source, Routine, eImportingPublicKey)
    Exit Function
End If
'
' Import the key blob into key container.
'
If Not (ImportPublicKey(PublicKey)) Then
'    Call pRaiseError(Source, Routine, eImportingPublicKey)
    Exit Function
End If
'
' Read the signature file.
'
lSigLen = 64
If Not (ReadFile(sFileName & ".sgn", bSig)) Then
'    Call pRaiseError(Source, Routine, eReadingSignatureFile)
    Exit Function
End If
'
' Hash the file.
'
If Not (HashFile(sFileName, hHash)) Then
'    Call pRaiseError(Source, Routine, eGeneratingHash)
    Exit Function
End If

lResult = CryptVerifySignature(hHash, bSig(1), lSigLen, m_hKeyPair, vbNullString, 0)
If lResult = 0 Then
    VerifyFile = False
Else
    VerifyFile = True
End If
'
' Destroy the hash.
'
lResult = CryptDestroyHash(hHash)
End Function

Private Function HashFile(sFileName As String, lHash As Long) As Boolean
Dim hHash           As Long
Dim lResult         As Long
Dim lFileHandle     As Long
Dim lBytesRemaining As Long
Dim abFile()        As Byte
'
' Hash a file in preparation for signing or verifying.
'
lResult = CryptCreateHash(m_hProvider, CALG_MD5, 0, 0, hHash)

If lResult = 0 Then
    HashFile = False
    Exit Function
End If

lFileHandle = FreeFile

On Error GoTo ErrorHandler
'
' Open the file.
'
Open sFileName For Binary As lFileHandle
'
' Progressively hash the file in chunks.
'
ReDim abFile(1 To C_CHUNK_SIZE)
'
' Handle whole chunks.
'
While (LOF(lFileHandle) - Seek(lFileHandle) >= C_CHUNK_SIZE)
    Get lFileHandle, , abFile
    lResult = CryptHashData(hHash, abFile(1), C_CHUNK_SIZE, 0)
    If lResult = 0 Then
        HashFile = False
        Exit Function
    End If
Wend
'
' Handle any leftover bytes.
'
lBytesRemaining = (LOF(lFileHandle) - Seek(lFileHandle) + 1)

If lBytesRemaining > 0 Then
    ReDim abFile(1 To lBytesRemaining)
    Get lFileHandle, , abFile
    lResult = CryptHashData(hHash, abFile(1), lBytesRemaining, 0)
    If lResult = 0 Then
        HashFile = False
        Exit Function
    End If
End If

Close lFileHandle

HashFile = True
'
' Return a handle to the hash object.
'
lHash = hHash
Exit Function
    
ErrorHandler:
    HashFile = False
End Function

Private Function AcquireCryptoProvider() As Boolean
Dim lResult As Long
'
' Get a handle to the base cryptographic provider.
'
'
' There is no good way to know if a key container
' exists other than to assume it does and test the
' return code.
'
lResult = CryptAcquireContext(m_hProvider, vbNullString, _
                 vbNullString, PROV_RSA_FULL, 0)
If lResult = 0 Then
    lResult = CryptAcquireContext(m_hProvider, App.EXEName, _
                     vbNullString, PROV_RSA_FULL, 0)
                     
End If

'
' If unsuccessful, we need to create a container.
'
If lResult = 0 Then
    If err.LastDllError = NTE_EXISTS Or err.LastDllError = NTE_BAD_KEYSET Then
        lResult = CryptAcquireContext(m_hProvider, App.EXEName, _
                 vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
    End If
End If

If lResult = 0 Then
    AcquireCryptoProvider = False
    '
    ' Clear the provider handle.
    '
    m_hProvider = 0
Else
    AcquireCryptoProvider = True
End If
End Function

Private Function IsProviderOpen() As Boolean
'
' See if a valid crypto provider handle is held.
'
IsProviderOpen = (m_hProvider <> 0)
End Function

Private Function ExportPublicKey(ByRef PublicKey As T_EXP_PUBLICKEYBLOB) As Boolean
Dim lLenKey As Long
Dim lResult As Long
'
' Export the public part of a key
' into the PublicKey variable.
'
If m_eKeyStatus = eKeyNone Or m_hKeyPair = 0 Then
    ExportPublicKey = False
    Exit Function
End If

lLenKey = LenB(PublicKey)
lResult = CryptExportKey(m_hKeyPair, 0, PUBLICKEYBLOB, 0, PublicKey, lLenKey)

If lResult = 0 Then
    ExportPublicKey = False
Else
    ExportPublicKey = True
End If
End Function

Private Function ExportPrivateKey(ByRef PrivateKey As T_EXP_PRIVATEKEYBLOB) As Boolean
Dim lLenKey As Long
Dim lResult As Long
'
' Export the private part of a key
' into the PrivateKey variable
'
If m_eKeyStatus <> eKeyPrivate Or m_hKeyPair = 0 Then
    ExportPrivateKey = False
    Exit Function
End If

lLenKey = LenB(PrivateKey)
lResult = CryptExportKey(m_hKeyPair, 0, PRIVATEKEYBLOB, 0, PrivateKey, lLenKey)
    
If lResult = 0 Then
    ExportPrivateKey = False
Else
    ExportPrivateKey = True
End If
End Function

Private Function ImportPrivateKey(ByRef PrivateKey As T_EXP_PRIVATEKEYBLOB) As Boolean
Dim lLenKey As Long
Dim lResult As Long
'
' Import a private key into the key container.
'
lLenKey = LenB(PrivateKey)
lResult = CryptImportKey(m_hProvider, PrivateKey, lLenKey, 0, 0, m_hKeyPair)
 
If lResult = 0 Then
    ImportPrivateKey = False
Else
    ImportPrivateKey = True
End If
End Function

Private Function ImportPublicKey(ByRef PublicKey As T_EXP_PUBLICKEYBLOB) As Boolean
Dim lLenKey As Long
Dim lResult As Long
'
' Import a public key into the key container.
'
lLenKey = LenB(PublicKey)
lResult = CryptImportKey(m_hProvider, PublicKey, lLenKey, 0, 0, m_hKeyPair)
ImportPublicKey = (lResult <> 0)
End Function

Private Function WriteFile(sFileName As String, Data() As Byte) As Boolean
Dim lFileHandle As Long
'
' Writes a byte array out to a file.
'
On Error GoTo ErrorHandler

lFileHandle = FreeFile

Open sFileName For Binary As lFileHandle
Put lFileHandle, , Data
Close lFileHandle

WriteFile = True
Exit Function

ErrorHandler:
    WriteFile = False
End Function

Private Function ReadFile(sFileName As String, Data() As Byte) As Boolean
Dim lFileHandle As Long
'
' Reads a byte array from a file.
' If Data is fixed, it must be of the expected size.
'
On Error GoTo ErrorHandler

lFileHandle = FreeFile

Open sFileName For Binary As lFileHandle

ReDim bData(1 To LOF(lFileHandle))

Get lFileHandle, , Data
Close lFileHandle

ReadFile = True
Exit Function

ErrorHandler:
    ReadFile = False
End Function

