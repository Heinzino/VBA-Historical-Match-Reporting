Option Explicit
'Globals
Dim File_Found As Boolean
Dim Source_Name As String
Dim Generator_Cell As Range 'Loop through teams in G_Sheet


Sub Run_Process() 'master Macro

If MsgBox("Run report process?", vbYesNo, "Confirmation") = _
vbNo Then Exit Sub


Call Find_Data_File

If File_Found = False Then Exit Sub

Call Delete_Report_Sheets
Call Copy_Template_Sheet_Teams
Call Find_Fixture_In_Data_File

End Sub

Sub Find_Data_File()


'a=b
Source_Name = ThisWorkbook.Sheets("Generator"). _
Range("G_Source_Data").Value

Dim Chris_WB As Workbook


File_Found = False
For Each Chris_WB In Workbooks

    If Chris_WB.Name = Source_Name Then File_Found = True
    
Next Chris_WB

If File_Found = False Then
MsgBox "Data File not found", 0, "Check"
End If

End Sub

Sub Find_Fixture_In_Data_File()

Dim Sheet_Number As Integer
Dim Data_Cell As Range
Dim Target_Team As String
Dim Home_Mode As Boolean
Dim Target_Col As String
Dim Team_Offset As Integer

Application.ScreenUpdating = False

For Each Generator_Cell In ThisWorkbook.Sheets("Generator") _
                           .Range("G_Teams")
                           
Target_Team = Generator_Cell

'Default to Away
Home_Mode = False
Target_Col = "I5:I384"
Team_Offset = -1

If Generator_Cell.Offset(0, 1).Value = "HOME" Then
    Home_Mode = True
    Target_Col = "H5:H384"
    Team_Offset = 1
End If
    
    For Sheet_Number = 1 To 3
    
        With Workbooks(Source_Name).Sheets(Sheet_Number)
        
        
             
            For Each Data_Cell In .Range(Target_Col)
                
                If Data_Cell.Value = "Flatcoat Retriever" And _
                                     Data_Cell.Offset(0, Team_Offset).Value = Target_Team _
                                     Then
                    ThisWorkbook.Sheets(Target_Team).Range("D11").Offset(1 - Sheet_Number, 0) _
                                                    = .Cells(Data_Cell.Row, 13) 'Match day col 13
                    ThisWorkbook.Sheets(Target_Team).Range("E11").Offset(1 - Sheet_Number, 0) _
                                                    = .Cells(Data_Cell.Row, 14) 'Result col 13
                End If
                
            Next Data_Cell
        
        
        End With
    
    Next Sheet_Number

Call Create_PDF
Next Generator_Cell

Call Get_Email_Address

Application.ScreenUpdating = True
ThisWorkbook.Activate
Sheets("Generator").Select
MsgBox "Task complete", 0, "Complete" 'Specify exactly what tasks happened

End Sub



Sub Copy_Template_Sheet_Teams()

Dim Team_Name As Range

For Each Team_Name In ThisWorkbook.Sheets("Generator"). _
                      Range("G_Teams")

ThisWorkbook.Sheets("Template").Copy After:=Sheets(Sheets.Count)
ActiveSheet.Name = Team_Name.Value
ActiveSheet.Range("C6") = Team_Name.Value & " - " & _
                          Team_Name.Offset(0, 1).Value

Next Team_Name


End Sub

Sub Delete_Report_Sheets()

Dim Chris_Sheet As Worksheet

Application.DisplayAlerts = False
For Each Chris_Sheet In ThisWorkbook.Sheets

If Chris_Sheet.Index > 2 Then Chris_Sheet.Delete

Next Chris_Sheet
Application.DisplayAlerts = True

End Sub

Sub Create_PDF()

Dim Save_Name As String 'Join together file path and file name for saving the pdf
Dim Sheet_To_Print As String

Sheet_To_Print = Generator_Cell

ThisWorkbook.Activate
ThisWorkbook.Sheets(Sheet_To_Print).Select

Save_Name = ThisWorkbook.Sheets("Generator").Range("G_PDF_Save_Location")
Save_Name = Save_Name & ActiveSheet.Name & ".pdf"

ActiveSheet.PageSetup.PrintArea = "$A$1:$F$13"
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
Save_Name, Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=True

End Sub

Sub Get_Email_Address()


Dim Counter As Integer
Dim Email_Address As String


Counter = 1

Do

    With ThisWorkbook.Sheets("Generator")
        
        Email_Address = .Range("G_Email_Start").Offset(Counter, 0)
        
        If .Range("G_Email_Start").Offset(Counter, 1) = "Y" Then
            Call Send_Email(Email_Address)
        End If
        Counter = Counter + 1
        
    End With

Loop Until Counter = 4


End Sub

Sub Send_Email(Email_Address As String)

MsgBox "Email will send to " & Email_Address
Dim OutApp As Object
Dim OutMail As Object

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

With OutMail

    '.CC = ""
    '.BCC = ""
    .Subject = "Luna's Pre-match Analysis"
    .Body = "The pre-match analysis is now ready"
    .To = Email_Address
    .Send
    
End With

Set OutMail = Nothing
Set OutApp = Nothing

End Sub
