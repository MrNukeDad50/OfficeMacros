Attribute VB_Name = "modMain"
Public myDistList As String

Public Sub BigSheetUpdate()
Attribute BigSheetUpdate.VB_ProcData.VB_Invoke_Func = "m\n14"
    ' This is a form driven program to update the bigsheet
    ' The form creates a bigsheet object, populates it with the needed parameters
    ' The bigsheet object does all of the logic and updates.
    ' Set default values for the update in the form code.
    frmSelectDistList.Show
End Sub
Public Sub BigSheetExport()
    Dim myBS As New clsBSExport
    
    myBS.GroomBacklog
    myBS.GroomWorklist
    myBS.ExportBS
    
End Sub



