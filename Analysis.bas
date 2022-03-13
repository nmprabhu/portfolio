Attribute VB_Name = "ParseDataAndSummarize"
Option Explicit

Public Sub ParseInfo()

    'defining variables
    Dim maxPages As String
    Dim cancel As String
    Dim cell As Range
    Dim ownedByLocation As Double
    Dim ownedByString As String
    Dim ownedBySpace As Double
    Dim name As String
    Dim pagesLocation As Double
    Dim pagesString As String
    Dim pagesPeriod As Double
    Dim pages As Double
    Dim row As Range
    
    'naming range messages
    Worksheets("OfficePrinter").Range(Worksheets("OfficePrinter").Range("B2"), Worksheets("OfficePrinter").Range("B2").End(xlDown)).name = "Messages"
    
    'asking user for max amount of pages
    maxPages = InputBox("Enter in the max amount of pages printed, as a number greater than 0.", "Max Pages")
    
    'if user clicks cancel
    If maxPages = "" Then
        cancel = MsgBox("Would you like to quit?", vbYesNoCancel, "Quit")
        If cancel = vbYes Or cancel = vbCancel Then
            Call MsgBox("Goodbye!", , "Goodbye")
            Exit Sub
        End If
        If cancel = vbNo Then
            Call ParseInfo
        End If
    End If
    
    'error if they enter a non-number or greater than 0
    If maxPages <= 0 Or IsNumeric(maxPages) = False Then
        Worksheets("Results").Range("F1").Value = " "
        Call ParseInfo
    End If
    
    'enter threshold in F1
    Worksheets("Results").Range("F1").Value = maxPages
    
    'writing in labels
    With Worksheets("Results").Range("A1")
        .Value = "Person"
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    
    With Worksheets("Results").Range("B1")
        .Value = "Number of Print Jobs"
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    
    With Worksheets("Results").Range("C1")
        .Value = "Total Pages"
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    
    With Worksheets("Results").Range("D1")
        .Value = "Average Number Pages"
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    
    With Worksheets("Results").Range("E1")
        .Value = "Number of Times Above Threshold"
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    
    For Each cell In Range("Messages")
        
        'parse out number of pages
        pagesLocation = InStr(cell, "printed: ") + 9
        pagesString = Mid(cell, pagesLocation)
        pagesPeriod = InStr(pagesString, ".")
        pages = Mid(pagesString, 1, pagesPeriod - 1)
        
        'put number of pages in results sheet
        With Worksheets("Results").Range("C1000").End(xlUp).Offset(1, 0)
            .Value = pages
            
            'highlight if above threshold
            If pages > maxPages Then
                .Interior.ColorIndex = 27
            End If
        End With
        
        'parse out person's name
        ownedByLocation = InStr(cell, "owned by ") + 9
        ownedByString = Mid(cell, ownedByLocation)
        ownedBySpace = InStr(ownedByString, " ")
        name = Mid(ownedByString, 1, ownedBySpace - 1)
        
        'put name in results sheet
        With Worksheets("Results").Range("A1000").End(xlUp).Offset(1, 0)
            .Value = name
            
            'highlight if above threshold
            If pages > maxPages Then
                .Interior.ColorIndex = 27
            End If
        End With
        
    Next cell

End Sub

Public Sub ClearResults()
    
    'defining variables
    Dim clear As String
    
    'msgbox with if user wants to clear
    clear = MsgBox("Do you really want to clear the data?", vbYesNo, "Clear Data")
    
    'if no then cancel procedure
    If clear = vbNo Then
        Exit Sub
    End If
    
    'if yes then clear
    If clear = vbYes Then
        With Worksheets("Results").Range(Worksheets("Results").Range("A2"), Worksheets("Results").Range("A2").SpecialCells(xlLastCell))
            .ClearContents
            .ClearFormats
        End With
    End If

End Sub

Public Sub Summarize()

    'defining variables
    Dim maxPages As Double
    Dim ownedByLocation As Double
    Dim ownedByString As String
    Dim ownedBySpace As Double
    Dim name As String
    
    Dim pagesLocation As Double
    Dim pagesString As String
    Dim pagesPeriod As Double
    Dim pages As Double
    
    Dim cell As Range
    Dim firstTime As Double
    Dim cell1 As Range
    Dim cell2 As Range
    Dim isFound As Boolean
    
    'retrieving threshold from F1
    maxPages = Worksheets("Results").Range("F1").Value
    
    'call clear results sub
    Call ClearResults
    
    'set firstTime as 1 to show its the first time
    firstTime = 1
    
    For Each cell In Range("Messages")
        
        'parse out person's name
        ownedByLocation = InStr(cell, "owned by ") + 9
        ownedByString = Mid(cell, ownedByLocation)
        ownedBySpace = InStr(ownedByString, " ")
        name = Mid(ownedByString, 1, ownedBySpace - 1)
        
        'parse out pages
        pagesLocation = InStr(cell, "printed: ") + 9
        pagesString = Mid(cell, pagesLocation)
        pagesPeriod = InStr(pagesString, ".")
        pages = Mid(pagesString, 1, pagesPeriod - 1)
        
        isFound = False
        
        'if its the first time then name in first row
        If firstTime = 1 Then
            Worksheets("Results").Range("A2").Value = name
            firstTime = 0
        
        'do this for loop for every other time
        Else: For Each cell1 In Worksheets("Results").Range(Worksheets("Results").Range("A1"), Worksheets("Results").Range("A1").End(xlDown))
            
            'if name is in list already
            If cell1.Value = name Then
                isFound = True
                
                'add one print to number of prints
                cell1.Offset(0, 1) = cell1.Offset(0, 1) + 1
                
                'add pages to number of pages
                cell1.Offset(0, 2) = cell1.Offset(0, 2) + pages
                
                Exit For
            End If

            Next cell1
            
            'if name isnt found
            If isFound = False Then
            
                'add name to next row
                Worksheets("Results").Range("A1000").End(xlUp).Offset(1, 0).Value = name
                
                'add one print to new row
                Worksheets("Results").Range("B1000").End(xlUp).Offset(1, 0).Value = 1
                
                'add pages to next row
                Worksheets("Results").Range("C1000").End(xlUp).Offset(1, 0).Value = pages
                
                'add 0 to new row for average
                Worksheets("Results").Range("D1000").End(xlUp).Offset(1, 0).Value = 0
                
            End If
            
        End If
        
    Next cell
    
    For Each cell2 In Worksheets("Results").Range(Worksheets("Results").Range("C2"), Worksheets("Results").Range("C2").End(xlDown))
        
        'average pages
        cell2.Offset(0, 1).Value = (cell2.Value) / (cell2.Offset(0, -1).Value)
        
        'times above threshold
        If cell2.Value >= maxPages Then
        
            'over max page amount
            cell2.Offset(0, 2).Value = (cell2.Value) - maxPages
            
            'under max page amount
        Else: cell2.Offset(0, 2).Value = "Under Threshold"
        End If
        
    Next cell2
    
End Sub

Public Sub ParseInfoAndSummarize()

    'call parse info
    Call ParseInfo
    
    'call summarize
    Call Summarize
    
End Sub
