VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTimeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private bVerbose As Boolean
Private sEmail As String
Private dDate As Date

Private Property Get Verbose() As Boolean
    Verbose = bVerbose
End Property

Private Property Let Verbose(Value As Boolean)
    bVerbose = Value
End Property

Private Property Get Email() As String
    Email = sEmail
End Property
Private Property Let Email(Value As String)
    sEmail = Value
End Property
