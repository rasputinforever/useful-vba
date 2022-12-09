' This will iterate through Sheet 1 column A through Row 10,000. 
' If any cell fails to provide a working filepath it will skip over that cell
' When it tries to iterate through a blank cell it kills the sub
' This example copy-pastes formatting. This section can be replaced with your desired actions

Sub processWorkbooks()

Dim tWb As Workbook
Dim sWb As Workbook
Dim tSht As Worksheet
Dim sSht As Worksheet
Dim tRng As Range
Dim sRng As Range

' Speeds Up Process
Application.ScreenUpdating = False

' Iterate through Range
For Each cell In Sheets("Name of Sheet where your range is located").Range("Range of Filepaths")
    
    On Error GoTo SkipCell
    
    ' exit at end of list. This helps prevent infinite loops
    If cell.Value2 = "" Then
        GoTo EndSub        
        Exit Sub
    Else: End If
    
    ' open workbook    
    Debug.Print "Processing this file: " & cell.Value2
    Set tWb = Workbooks.Open(cell.Value2) ' This is the target sheet filepath
    Set sWb = Workbooks.Open("Your Source Here") ' Alternatively, you can just set this to tWb
    Set tSht = tWb.Sheets("Target Sheet Name Here")
    Set sSht = sWb.Sheets("Source Sheet Here")
    Set tRng = tSht.Range("Target Range Here")
    Set sRng = sSht.Range("Source Range Here")
    
    ' unprotect if necessary
    tSht.Unprotect "password"
      
      ' ===YOUR DESIRED ACTIONS START HERE===
      
        ' In this example we are simply copy-pasting formatting
          sRng.Copy
          tRng.PasteSpecial Paste:=xlPasteFormats
        
      ' ===YOUR DESIRED ACTIONS END HERE===
      
    ' reprotect if necessary
    tSht.Protect "password"
   
   ' close and save
    tWb.Close SaveChanges:=True
    ' close but don't save. Remove this line if sWb = tWb. Move it to EndSub if using single point of source
    sWb.Close SaveChanges:=False

SkipCell:

Next cell


EndSub:
Debug.Print "Complete"
' Reinstate
Application.ScreenUpdating = True

End Sub
