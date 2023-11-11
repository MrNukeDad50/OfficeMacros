Attribute VB_Name = "modListTest"
Sub test()
    Dim list As New vbaList
    Set list = list.CreateInstance

    list.Add 1
    list.Add 9
    list.Add 6
    list.Add 13
    list.Add 2
    list.Add 6
    list.Add 4, 3
    list.Remove 13
    list.RemoveAtIndex 2
    list.Add "Test 1"
    list.Add "Test 2"
    list.Add 6
 
    Dim listCopy As New vbaList
    Set listCopy = list.Copy

    Dim i As Long

    Debug.Print "========================================"
    Debug.Print "IndexOf Pos: " & list.IndexOf(6)
    Debug.Print "LastIndexOf Pos: " & list.LastIndexOf(6)
    Debug.Print "Find Test 1 @ Pos: " & list.Find("Test 1")
    Debug.Print "[Test 1] exists: " & list.Exists("Test 1")
    Debug.Print "[Test 3] exists: " & list.Exists("Test 3")
    Debug.Print "Count: " & list.Count
    list.Clear
    Debug.Print "Clear() Count: " & list.Count
    list.Dispose
    Debug.Print "Disposed: " & list.Disposed
    
    Debug.Print ""
    For i = 0 To listCopy.Count - 1
        Debug.Print "Default - Pos " & i & ": " & listCopy.Items(i)
    Next i
    listCopy.Reverse
    For i = 0 To listCopy.Count - 1
        Debug.Print "Reverse - Pos " & i & ": " & listCopy.Items(i)
    Next i
End Sub
