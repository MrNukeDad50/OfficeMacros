VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDistList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' requres the following references:
'1. Microsoft Outlook 16.0 Object Library for the Exchange Distribution List.
'2. clsDistList
'3. vbaList

Private myNames As vbaList
Private myIDs As vbaList


Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
    Set myNames = Nothing
    Set myIDs = Nothing
End Sub

Private Sub AddMember(ByRef olEntry As Outlook.AddressEntry)
    Dim olEntriesL As Outlook.AddressEntries
    Dim olEntryL As Outlook.AddressEntry
    Dim olDLL As Outlook.ExchangeDistributionList
    Dim myID As String
    
    If olEntry.AddressEntryUserType = olExchangeDistributionListAddressEntry Then
        Set olDLL = olEntry.GetExchangeDistributionList
        Set olEntriesL = olDLL.GetExchangeDistributionListMembers
        For Each olEntryL In olEntriesL
            AddMember olEntryL
        Next olEntryL
    Else
        If Not (myNames.Contains(olEntry.name)) Then myNames.Add olEntry.name
        myID = Right(olEntry.Address, (Len(olEntry.Address) - InStr(olEntry.Address, "-")))
        Debug.Print olEntry.Address
        If myID = "X2DDARR" Then myID = "DDARR"
        If myID = "X2PGIBBS" Then myID = "PGIBBS"
        If Not (myIDs.Contains(myID)) Then myIDs.Add (myID)
    End If
    
End Sub

Public Function NameListFrom(distListName As String) As vbaList
    GetListOfNamesIn (distListName)
    Set NameListFrom = myNames
    
End Function
Public Function IDListFrom(distListName As String) As vbaList
    GetListOfNamesIn (distListName)
    Set IDListFrom = myIDs
    
End Function

Private Sub GetListOfNamesIn(ListName As String)
    Dim olApp As Outlook.Application
    Dim olEntry As Outlook.AddressEntry
    Dim olEntryL As Outlook.AddressEntry
    Dim olDL As Outlook.ExchangeDistributionList
    Dim olDLL As Outlook.ExchangeDistributionList
    Dim olEntries As Outlook.AddressEntries
    Dim olEntriesL As Outlook.AddressEntries
    Dim myID As String
    
    Set myNames = New vbaList
    Set myIDs = New vbaList
    Set olApp = New Outlook.Application
    Set olEntry = olApp.Session.GetGlobalAddressList.AddressEntries(ListName)
    Set olDL = olEntry.GetExchangeDistributionList
    Set olEntry = Nothing
    Set olEntries = olDL.GetExchangeDistributionListMembers
    For Each olEntry In olEntries
        AddMember olEntry
    Next olEntry

End Sub
