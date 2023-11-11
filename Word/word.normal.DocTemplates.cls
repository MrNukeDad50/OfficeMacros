VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Worklog_New(myDoc As Document)
'Purpose: Take the existing document and create a working folder and MS Word worklog
' The working folder should be put in "C:\Users\ddarr\Documents\My Work Log\"
' The working folder should follow the format "<ISO date> - <task name>"
' The worklog title should follow the format "<ISO Date> - Worklog - <task name>.docx"
' The title of the doc should be repeated at the top of the worklog document page as "<ISO Date> - <task name>"

    Dim sWorkLogTitle As String
    Dim sDate As String
    Dim fobj As Object
    
    sWorkLogTitle = InputBox("Enter the task Name for the worklog", "Enter Name")
    sDate = Format(Now(), "YYYY-MM-DD")
    
    myDoc.Activate
    Selection.TypeText sDate & " - Worklog - " & sWorkLogTitle
    Selection.Style = ActiveDocument.Styles("Title")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph

    Set fobj = CreateObject("Scripting.FIleSystemObject")
    If Not fobj.folderexists("C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle) Then
        fobj.createfolder ("C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle)
    End If
    myDoc.SaveAs2 "C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle & "\" & sDate & " - Worklog - " & sWorkLogTitle

    Set fobj = Nothing
End Sub

Public Sub DocReview_new(myDoc As Document)
'Purpose: Take the existing document and create a working folder and MS Word worklog
' The working folder should be put in "C:\Users\ddarr\Documents\My Work Log\"
' The working folder should follow the format "<ISO date> - <task name>"
' The worklog title should follow the format "<ISO Date> - Worklog - <task name>.docx"
' The title of the doc should be repeated at the top of the worklog document page as "<ISO Date> - Worklog - <task name>"
' Append the document with the project directed template information to ensure the checklist is consistent.

    Dim sWorkLogTitle As String
    Dim sDate As String
    Dim fobj As Object
    Dim myDocHelper As New DocHelpers
    
'Get basic doc info needed
    sWorkLogTitle = InputBox("Enter the number and rev of the document to be reviewed <DocNo> Rev<RevNo>", "Enter DocNo and Rev")
    sWorkLogTitle = "Review of " & sWorkLogTitle
    sDate = Format(Now(), "YYYY-MM-DD")
    
'Generate Body of checklist / worklog
'Title
    myDoc.Activate
    Selection.TypeText sDate & " - Worklog - " & sWorkLogTitle
    Selection.Style = ActiveDocument.Styles("Title")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
'Purpose
    Selection.TypeText "Purpose:"
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    
        Selection.TypeText "Provide a technical review of " & sWorkLogTitle & " per TP-ENG-PROC-0022"
        Selection.Style = ActiveDocument.Styles("Normal")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
'Method
    Selection.TypeText "Method:"
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph

        Selection.TypeText "Perform a technical review of " & sWorkLogTitle & " using the requirements and guidance in:"
        Selection.Style = ActiveDocument.Styles("Normal")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
            'List formatting ...
            With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
                .NumberFormat = "%1."
                .TrailingCharacter = wdTrailingTab
                .NumberStyle = wdListNumberStyleArabic
                .NumberPosition = InchesToPoints(0.25)
                .Alignment = wdListLevelAlignLeft
                .TextPosition = InchesToPoints(0.5)
                .TabPosition = wdUndefined
                .ResetOnHigher = 0
                .StartAt = 1
                With .Font
                    .Bold = wdUndefined
                    .Italic = wdUndefined
                    .StrikeThrough = wdUndefined
                    .Subscript = wdUndefined
                    .Superscript = wdUndefined
                    .Shadow = wdUndefined
                    .Outline = wdUndefined
                    .Emboss = wdUndefined
                    .Engrave = wdUndefined
                    .AllCaps = wdUndefined
                    .Hidden = wdUndefined
                    .Underline = wdUndefined
                    .Color = wdUndefined
                    .Size = wdUndefined
                    .Animation = wdUndefined
                    .DoubleStrikeThrough = wdUndefined
                    .Name = ""
                End With
                .LinkedStyle = ""
            End With
            ListGalleries(wdNumberGallery).ListTemplates(1).Name = ""
            Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
                False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                wdWord10ListBehavior
            'list contents
            Selection.TypeText Text:="TP-ENG-PROC-0022, Engineering Review Process and Independent Design Verification" & vbCrLf
            Selection.TypeText Text:="MCRE Project Release Process " & vbCrLf
            Selection.TypeText Text:="Informal Peer Check " & vbCrLf
            Selection.TypeText Text:="Update Managed Task " & vbCrLf
            Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
'Checklist
    Selection.TypeText "Checklist"
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph
    'link to TP-ENG-PROC-0022
        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
            "https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D6708421%26tabid%3D0%26", _
            SubAddress:="", _
            ScreenTip:="", _
            TextToDisplay:="TP-ENG-PROC-0022"
        Selection.Style = ActiveDocument.Styles("Heading 2")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
                'List Formatting
                With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
                    .NumberFormat = ChrW(9744)
                    .TrailingCharacter = wdTrailingTab
                    .NumberStyle = wdListNumberStyleBullet
                    .NumberPosition = InchesToPoints(0.25)
                    .Alignment = wdListLevelAlignLeft
                    .TextPosition = InchesToPoints(0.5)
                    .TabPosition = wdUndefined
                    .ResetOnHigher = 0
                    .StartAt = 1
                    With .Font
                        .Bold = wdUndefined
                        .Italic = wdUndefined
                        .StrikeThrough = wdUndefined
                        .Subscript = wdUndefined
                        .Superscript = wdUndefined
                        .Shadow = wdUndefined
                        .Outline = wdUndefined
                        .Emboss = wdUndefined
                        .Engrave = wdUndefined
                        .AllCaps = wdUndefined
                        .Hidden = wdUndefined
                        .Underline = wdUndefined
                        .Color = wdUndefined
                        .Size = wdUndefined
                        .Animation = wdUndefined
                        .DoubleStrikeThrough = wdUndefined
                        .Name = "Segoe UI Symbol"
                    End With
                    .LinkedStyle = ""
                End With
                ListGalleries(wdBulletGallery).ListTemplates(1).Name = ""
                Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                    ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
                    False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                    wdWord10ListBehavior
                'List Contents
                Selection.TypeText Text:="Review Type and Scope Determination (5.1)" & vbCrLf
                Selection.TypeText Text:="General Information and Tools (5.2)" & vbCrLf
                    Selection.Range.ListFormat.ListIndent
                    Selection.TypeText Text:="Three Pass Review - Appendix 9.5" & vbCrLf
                        Selection.Range.ListFormat.ListIndent
                        Selection.TypeText Text:="Overview Pass" & vbCrLf
                        Selection.TypeText Text:="Details Pass" & vbCrLf
                        Selection.TypeText Text:="Challenge Pass" & vbCrLf
                        Selection.Range.ListFormat.ListOutdent
                    Selection.TypeText Text:="Highlight Review - Appendix 9.6" & vbCrLf
                    Selection.TypeText Text:="Independent Review - Appendix 9.7" & vbCrLf
                    Selection.Range.ListFormat.ListOutdent
                Selection.TypeText Text:="Technical Reviews (Section 5.3)" & vbCrLf
                    Selection.Range.ListFormat.ListIndent
                    Selection.TypeText Text:="Input Review" & vbCrLf
                    Selection.TypeText Text:="Assumption Review" & vbCrLf
                    Selection.TypeText Text:="Methodology Review" & vbCrLf
                    Selection.TypeText Text:="Results Review" & vbCrLf
                    Selection.TypeText Text:="Conclusion Review" & vbCrLf
                    Selection.TypeText Text:="Administrative Review" & vbCrLf
                    Selection.TypeText Text:="Review Scope Coverage Check" & vbCrLf
                    Selection.Range.ListFormat.ListOutdent
                Selection.TypeText Text:="Checking Reviews (Section 5.4)" & vbCrLf
                Selection.TypeText Text:="Performing IDVs (Section 5.5)" & vbCrLf
                Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph
        
        'link to MCRE Release Process
        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
            "https://collab.terrapower.com/x/wwo1Bg", _
            SubAddress:="", _
            ScreenTip:="", _
            TextToDisplay:="MCRE Release Process"
        Selection.Style = ActiveDocument.Styles("Heading 2")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
            'List Formatting
            With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
                .NumberFormat = ChrW(9744)
                .TrailingCharacter = wdTrailingTab
                .NumberStyle = wdListNumberStyleBullet
                .NumberPosition = InchesToPoints(0.25)
                .Alignment = wdListLevelAlignLeft
                .TextPosition = InchesToPoints(0.5)
                .TabPosition = wdUndefined
                .ResetOnHigher = 0
                .StartAt = 1
                With .Font
                    .Bold = wdUndefined
                    .Italic = wdUndefined
                    .StrikeThrough = wdUndefined
                    .Subscript = wdUndefined
                    .Superscript = wdUndefined
                    .Shadow = wdUndefined
                    .Outline = wdUndefined
                    .Emboss = wdUndefined
                    .Engrave = wdUndefined
                    .AllCaps = wdUndefined
                    .Hidden = wdUndefined
                    .Underline = wdUndefined
                    .Color = wdUndefined
                    .Size = wdUndefined
                    .Animation = wdUndefined
                    .DoubleStrikeThrough = wdUndefined
                    .Name = "Segoe UI Symbol"
                End With
                .LinkedStyle = ""
            End With
            ListGalleries(wdBulletGallery).ListTemplates(1).Name = ""
            Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
                ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
                False, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
                wdWord10ListBehavior
            'list contents
            Selection.TypeText Text:="All Types" & vbCrLf
                Selection.Range.ListFormat.ListIndent
                Selection.TypeText Text:="The correct, most-current template is being used" & vbCrLf
                Selection.TypeText Text:="Quality (spelling, formatting, legibility) is acceptable; use MS Word Grammar and Spelling" & vbCrLf
                Selection.TypeText Text:="Changes are clearly described or denoted" & vbCrLf
                Selection.TypeText Text:="Revision History table is filled out and includes JIRA ticket link." & vbCrLf
                Selection.TypeText Text:="Watermark or other such draft identification is present throughout." & vbCrLf
                Selection.TypeText Text:="Appropriate document markings are affixed:" & vbCrLf
                    Selection.Range.ListFormat.ListIndent
                    Selection.TypeText Text:="Proprietary & Confidential: "
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D6076353%26tabid%3D0%26", _
                            SubAddress:="", _
                            ScreenTip:="", _
                            TextToDisplay:="TP-LGL-POL-0001"
                    Selection.TypeText Text:=", TerraPower Communications Policy" & vbCrLf
                    Selection.TypeText Text:="Export Controlled: "
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D6034277%26tabid%3D13%26", _
                            SubAddress:="", _
                            ScreenTip:="", _
                            TextToDisplay:="TP-COM-POL-0001"
                    Selection.TypeText Text:=", Compliance with U.S. Export Control Laws & Regulations for TerraPower. "
                    Selection.TypeText Text:="Note: "
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D8510142%26tabid%3D%26", _
                            SubAddress:="", _
                            ScreenTip:="", _
                            TextToDisplay:="TP-CORP-PRSNT-0017"
                    Selection.TypeText Text:=" is a good basis for public information" & vbCrLf
                    Selection.TypeText Text:="Applied Technology: "
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D6034302%26tabid%3D0%26", _
                            SubAddress:="", _
                            ScreenTip:="", _
                            TextToDisplay:="TP-COM-PROC-0005"
                    Selection.TypeText Text:=", Applied Technology Management" & vbCrLf
                    Selection.TypeText Text:="Quality Level: "
                    ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, _
                            Address:="https://agile.terrapower.com/Agile/PLMServlet?fromPCClient=true&module=ItemHandler&requestUrl=module%3DItemHandler%26opcode%3DdisplayObject%26classid%3D9000%26objid%3D7119392%26tabid%3D13%26", _
                            SubAddress:="", _
                            ScreenTip:="", _
                            TextToDisplay:="TP-ENG-PROC-0043"
                    Selection.TypeText Text:=", Quality Level Marking" & vbCrLf
                    Selection.Range.ListFormat.ListOutdent
                Selection.Range.ListFormat.ListOutdent
            Selection.TypeText Text:="Documents" & vbCrLf
                Selection.Range.ListFormat.ListIndent
                Selection.TypeText Text:=" Tables that split across pages have a header row that repeats on each page. No rows are split across two pages." & vbCrLf
                Selection.Range.ListFormat.ListOutdent
            Selection.TypeText Text:="Drawings" & vbCrLf
                Selection.Range.ListFormat.ListIndent
                Selection.TypeText Text:="See guidance in the MCRE Release Process" & vbCrLf
                Selection.Range.ListFormat.ListOutdent
            Selection.Range.ListFormat.RemoveNumbers NumberType:=wdNumberParagraph

        'Peer Check
        Selection.TypeText Text:="Peer Check"
        Selection.Style = ActiveDocument.Styles("Heading 2")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph

        'link to Jira Ticket info
        ActiveDocument.Hyperlinks.Add Anchor:=Selection.Range, Address:= _
            "https://issuetrack.terrapower.com/browse/MCRE-2001?src=confmacro", _
            SubAddress:="", _
            ScreenTip:="", _
            TextToDisplay:="MCRE Document Release Managed Jira Ticket for Cloning"
        Selection.Style = ActiveDocument.Styles("Heading 2")
        Selection.EndKey Unit:=wdLine
        Selection.TypeParagraph
'worklog
    Selection.TypeText "Worklog:"
    Selection.Style = ActiveDocument.Styles("Heading 1")
    Selection.EndKey Unit:=wdLine
    Selection.TypeParagraph

    myDocHelper.DateTimeStamp Selection

    
    Set fobj = CreateObject("Scripting.FIleSystemObject")
    If Not fobj.folderexists("C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle) Then
        fobj.createfolder ("C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle)
    End If
    myDoc.SaveAs2 "C:\Users\ddarr\Documents\My Work Log\" & sDate & " " & sWorkLogTitle & "\" & sDate & " - Worklog - " & sWorkLogTitle
    
    
    
    Set fobj = Nothing
    


Worklog:

End Sub
