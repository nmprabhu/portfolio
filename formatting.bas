Attribute VB_Name = "SupportFunctions"
Option Explicit

Sub Oops()

    'restores real estate data
    Sheets("Backup").Range("A1:D118").Copy
    Sheets("Real Estate").Range("A2").Select
    ActiveSheet.Paste
    
    'restores statistics table
    Sheets("Backup").Range("F3:I12").Copy
    Sheets("Real Estate").Range("F4").Select
    ActiveSheet.Paste
    
End Sub

Sub FormatData()
Attribute FormatData.VB_ProcData.VB_Invoke_Func = " \n14"

    'first row formatting
    With Range("A2:D2")
        .HorizontalAlignment = xlCenter
        .Interior.Color = 49407
        .Font.Bold = True
    End With
    
    'first column formatting
    With Range("A3:A119")
        .HorizontalAlignment = xlCenter
        .Interior.Color = 49407
    End With

    'inner data formatting
    With Range("B3:D119")
        .Interior.Color = 12566463
    End With
    
    'borders formatting
    With Range("A2:D119")
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
    With Range("A2:D2")
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With

End Sub
