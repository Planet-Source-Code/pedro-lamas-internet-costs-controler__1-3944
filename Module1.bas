Attribute VB_Name = "Module1"
'********************************************************
'
'This Module is all you need to check (via RAS) if you
'are currently Online
'
'IsRASConnected() returns True if you are connected.
'Otherwise, it returns False.
'
'********************************************************

Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasConn As Long, lpRASCONNSTATUS As Any) As Long

Private Const RAS95_MaxEntryName = 256
Private Const RAS_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 128
Private Const RASCS_DONE = &H2000&

Type RASCONN95
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public Function IsRASConnected() As Boolean
    Dim TRasCon(255) As RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim lReturn As Long
    Dim Tstatus As RASCONNSTATUS95
    
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize
    lReturn = RasEnumConnections(TRasCon(0), lg, lpcon)
    
    If lReturn Then
        MsgBox "O módulo de detecção de ligação não está a funcionar." & vbCrLf & "Contacte o seu administrador do sistema", vbApplicationModal + vbCritical, "Erro"
        Exit Function
    End If
    
    Tstatus.dwSize = 160
    lReturn = RasGetConnectStatus(TRasCon(0).hRasConn, Tstatus)
    
    If Tstatus.RasConnState = RASCS_DONE Then
        IsRASConnected = True
    Else
        IsRASConnected = False
    End If
End Function
