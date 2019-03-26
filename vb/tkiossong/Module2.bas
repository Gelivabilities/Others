Attribute VB_Name = "ini"

'**************************************
'Windows API/Global Declarations for :EZ
'     - .ini
'**************************************
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'**************************************
' Name: EZ - .ini
' Description:Access .ini files in the b
'     link of an eye. Use one line of your inp
'     ut to quickly retrive .ini values. With
'     the same one line of code write to your
'     .ini file. If you have any improvements
'     on this code, E-Mail me at "karatebob@ho
'     tmail.com".
' By: Frank Joseph Mattia
'
'
' Inputs:None
'
' Returns:Returns the value of a string
'     of an .ini file.
'
'Assumes:When you call this in your code
'     , this is the syntax you will need to us
'     e.
'
'Side Effects:None.
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.1382/lngWId.1/qx/
'     vb/scripts/ShowCode.htm
'for details.
'**************************************
Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
    '***************************************
    '     ****************************************
    '     ***************
    ' DESCRIPTION:Reads from an *.INI file s
    '     trFileName (full path & file name)
    ' RETURNS:The string stored in [strSecti
    '     onHeader], line beginning
    ' strVariableName=
    '***************************************
    '     ****************************************
    '     ***************
    ' Initialise variable
    Dim strReturn As String
    ' Blank the return string
    strReturn = String(255, Chr(0))
    'Get requested information, trimming the
    '     returned
    ' string
    mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
    '***************************************
    '     ****************************************
    '     **********************
    ' DESCRIPTION:Writes to an *.INI file ca
    '     lled strFileName (fullpath & file name)
    ' RETURNS:Integer indicating failure (0)
    '     or success (other)to write
    '***************************************
    '     ****************************************
    '     **********************
    mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function

