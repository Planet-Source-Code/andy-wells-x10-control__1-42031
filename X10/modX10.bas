Attribute VB_Name = "modX10"
Option Explicit

' CM17 only supports these commands
Public Const C_ON = &H2
Public Const C_OFF = &H3
Public Const C_DIM = &H4
Public Const C_BRIGHT = &H5

' possible future support
Public Const ALL_UNITS_OFF = &H0
Public Const ALL_LIGHTS_ON = &H1
Public Const ALL_LIGHTS_OFF = &H6
Public Const C_EXTENDED = &H7
Public Const C_HAIL_REQ = &H8
Public Const C_HAIL_ACK = &H9
Public Const C_PRE_SET_DIM1 = &HA
Public Const C_PRE_SET_DIM2 = &HB
Public Const C_EXTENDED_DATA_TRANSFER = &HC
Public Const C_STATAUS_ON = &HD
Public Const C_STATUS_OFF = &HE
Public Const C_STATUS_REQUEST = &HF
Public Const C_CLEAR_MEM = &H10            ' private command
