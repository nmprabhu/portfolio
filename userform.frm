VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTCMembership 
   Caption         =   "TC^2 Membership"
   ClientHeight    =   6400
   ClientLeft      =   -10
   ClientTop       =   0
   ClientWidth     =   11240
   OleObjectBlob   =   "userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTCMembership"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim comp As Range 'cell of current company
Dim cell As Range 'only used to search
Dim numberCompanies As Double 'keeps count of companies
Dim currentCompany As Double 'number of current company
Dim dt As Date
Dim quit As String

Private Sub cmdDate_Click() 'used date picker
    dt = Date
    dt = GetDateTimePicker(dt)
    Me.lblDate.Caption = dt
End Sub

Private Sub cmdFirst_Click()
    Set comp = Worksheets("Companies").Range("A2")
    Call Setup
End Sub

Private Sub cmdGoToCompany_Click() 'error checking for numeric value, integer, within bounds
    If IsNumeric(Me.txtGoTo.Text) = False Or Int(CDbl(Me.txtGoTo.Text)) <> CDbl(Me.txtGoTo.Text) Or CDbl(Me.txtGoTo.Text) > numberCompanies Or CDbl(Me.txtGoTo.Text) <= 0 Then
        Exit Sub
    Else: Set comp = Worksheets("Companies").Range("A2").Offset(CDbl(Me.txtGoTo.Text) - 1, 0)
        Call Setup
    End If
End Sub

Private Sub cmdLast_Click()
    Set comp = Range("CompanyIDs").End(xlDown)
    Call Setup
End Sub

Private Sub cmdNext_Click()

    'hits the end
    If comp.Offset(1, 0).Value = "" Then
        Exit Sub
        
    'goes to next
    Else: Set comp = comp.Offset(1, 0)
        Call Setup
    End If
    
End Sub

Private Sub cmdPrevious_Click()

    'hits the beginning
    If comp.Offset(-1, 0).Value = "CompanyID" Then
        Exit Sub
        
    'goes to previous
    Else: Set comp = comp.Offset(-1, 0)
        Call Setup
    End If

End Sub

Private Sub cmdResetInfo_Click()
    Call Setup
End Sub

Private Sub cmdSaveCompany_Click()
    
    'checking if company websites are valid
    If InStr(1, Me.txtCompanyWebsite.Value, ".") = 0 And Me.txtCompanyWebsite.Text = "" Then
        Call MsgBox("Please enter a valid company website before saving!", vbOKOnly, Error)
        Exit Sub
    End If
    If Me.cmbCountry.Text = "USA" And InStr(1, Me.txtCompanyWebsite.Text, "N/A") = 0 Then
        If InStr(1, Me.txtCompanyWebsite.Text, ".com") = 0 And InStr(1, Me.txtCompanyWebsite.Text, ".net") = 0 And InStr(1, Me.txtCompanyWebsite.Text, ".biz") = 0 And InStr(1, Me.txtCompanyWebsite.Text, ".edu") = 0 And InStr(1, Me.txtCompanyWebsite.Text, ".org") = 0 And InStr(1, Me.txtCompanyWebsite.Text, ".gov") = 0 Then
            Call MsgBox("Please enter a valid company website before saving!", vbOKOnly, "Error")
            Exit Sub
        End If
    End If
    
    'checking if all fields are filled
    If Me.txtCompanyName.Text = "" Then
        Call MsgBox("Please enter a company name before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.lblDate.Caption = "" Then
        Call MsgBox("Please enter a membership date before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.txtStreetAddress.Text = "" Or Me.txtCity.Text = "" Or Me.cmbStateProvince.Text = "" Or Me.txtZipcode.Text = "" Or Me.cmbCountry.Text = "" Then
        Call MsgBox("Please enter a complete location before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.txtAnnualSales.Text = "" Then
        Call MsgBox("Please enter annual sales before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.txtNumberOfEmployees.Text = "" Then
        Call MsgBox("Please enter number of employees before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.cmbEndMarket.Text = "" Then
        Call MsgBox("Please select an end market before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.cmbProductType.Text = "" Then
        Call MsgBox("Please select a product type before saving!", vbOKOnly, "Error")
        Exit Sub
    End If
    If Me.txtComments.Text = "" Then
        Call MsgBox("Please enter a comment before saving! This can be N/A.", vbOKOnly, "Error")
        Exit Sub
    End If

    'call save sub
    Call Save

End Sub

Private Sub cmdSearchCompany_Click()
    
    'searches for company id
    If Me.cmbSearch.Text = "Company ID" Then
        If IsNumeric(Me.txtSearch.Text) = False Or Int(CDbl(Me.txtSearch.Text)) <> CDbl(Me.txtSearch.Text) Or CDbl(Me.txtSearch.Text) > numberCompanies Or CDbl(Me.txtSearch.Text) <= 0 Then
            Exit Sub
        Else: For Each cell In Range("CompanyIDs")
                If CDbl(cell.Value) = CDbl(Me.txtSearch.Text) Then
                    Set comp = cell
                    Call Setup
                End If
            Next cell
        End If
    End If
    
    'searches for company name
    If Me.cmbSearch.Text = "Company Name" Then
        For Each cell In Range("Companies")
            If InStr(1, UCase(cell.Value), UCase(Me.txtSearch.Text)) > 0 Then
                Set comp = cell.Offset(0, -1)
                Call Setup
            End If
        Next cell
    End If

End Sub

Private Sub UserForm_Initialize()
    
    'naming ranges
    Worksheets("Companies").Range(Worksheets("Companies").Range("A2"), Worksheets("Companies").Range("A2").End(xlDown)).Name = "CompanyIDs"
    Worksheets("Companies").Range(Worksheets("Companies").Range("B2"), Worksheets("Companies").Range("B2").End(xlDown)).Name = "Companies"
    Worksheets("ProductTypes").Range(Worksheets("ProductTypes").Range("A2"), Worksheets("ProductTypes").Range("A2").End(xlDown)).Name = "ProductCapabilityID"
    Worksheets("ProductTypes").Range(Worksheets("ProductTypes").Range("B2"), Worksheets("ProductTypes").Range("B2").End(xlToRight).End(xlDown)).Name = "ProductTypes"
    Worksheets("EndMarkets").Range(Worksheets("EndMarkets").Range("A2"), Worksheets("EndMarkets").Range("A2").End(xlDown)).Name = "EndMarketID"
    Worksheets("EndMarkets").Range(Worksheets("EndMarkets").Range("B2"), Worksheets("EndMarkets").Range("B2").End(xlToRight).End(xlDown)).Name = "EndMarkets"
    Worksheets("States").Range(Worksheets("States").Range("A2"), Worksheets("States").Range("A2").End(xlToRight).End(xlDown)).Name = "States"
    Worksheets("Countries").Range(Worksheets("Countries").Range("A2"), Worksheets("Countries").Range("A2").End(xlDown)).Name = "Countries"

    'setting up dropdowns combo boxes
    Me.cmbCountry.RowSource = "Countries"
    Me.cmbStateProvince.RowSource = "States"
    Me.cmbEndMarket.RowSource = "EndMarkets"
    Me.cmbProductType.RowSource = "ProductTypes"
    With Me.cmbSearch
        .AddItem "Company ID"
        .AddItem "Company Name"
    End With

    'set up first company
    Set comp = Worksheets("Companies").Range("A2")
    Call Setup
    
End Sub

Public Sub Setup() 'general use sub for new company info
    
    'company count
    numberCompanies = Range("CompanyIDs").Rows.Count
    currentCompany = Worksheets("Companies").Range(Worksheets("Companies").Range("A2"), comp).Rows.Count
    Me.lblCompanyOrder.Caption = "Company " & currentCompany & " of " & numberCompanies
    
    'fills in values from company ws
    Me.lblCompanyID.Caption = comp.Value
    Me.txtCompanyName.Text = comp.Offset(0, 1).Value
    Me.lblDate.Caption = comp.Offset(0, 2).Value
    Me.chkActiveMember.Value = comp.Offset(0, 3).Value
    Me.txtStreetAddress.Text = comp.Offset(0, 4).Value
    Me.txtCity.Text = comp.Offset(0, 5).Value
    Me.cmbStateProvince.Text = comp.Offset(0, 6).Value
    Me.txtZipcode.Text = comp.Offset(0, 7).Value
    Me.cmbCountry.Text = comp.Offset(0, 8).Value
    Me.txtCompanyWebsite.Text = comp.Offset(0, 9).Value
    Me.txtAnnualSales.Text = comp.Offset(0, 10).Value
    Me.txtNumberOfEmployees = comp.Offset(0, 11).Value
    For Each cell In Range("EndMarketID")
        If cell.Value = comp.Offset(0, 12).Value Then
            Me.cmbEndMarket.Text = cell.Offset(0, 1).Value
        End If
    Next cell
    For Each cell In Range("ProductCapabilityID")
        If cell.Value = comp.Offset(0, 13).Value Then
            Me.cmbProductType.Text = cell.Offset(0, 1).Value
        End If
    Next cell
    Me.txtComments.Text = comp.Offset(0, 14).Value
    
    'state combo box/text box
    If Me.cmbCountry.Text = "USA" Then
        Me.cmbStateProvince.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    Else: Me.cmbStateProvince.ShowDropButtonWhen = fmShowDropButtonWhenNever
    End If
    
End Sub

Public Sub Save() 'saving all fields
    comp.Value = Me.lblCompanyID.Caption
    comp.Offset(0, 1).Value = Me.txtCompanyName.Text
    comp.Offset(0, 2).Value = Me.lblDate.Caption
    If Me.chkActiveMember.Value = True Then
        comp.Offset(0, 3).Value = 1
    End If
    If Me.chkActiveMember.Value = False Then
        comp.Offset(0, 3).Value = 0
    End If
    comp.Offset(0, 4).Value = Me.txtStreetAddress.Text
    comp.Offset(0, 5).Value = Me.txtCity.Text
    comp.Offset(0, 6).Value = Me.cmbStateProvince.Text
    comp.Offset(0, 7).Value = Me.txtZipcode.Text
    comp.Offset(0, 8).Value = Me.cmbCountry.Text
    comp.Offset(0, 9).Value = Me.txtCompanyWebsite.Text
    comp.Offset(0, 10).Value = Me.txtAnnualSales.Text
    comp.Offset(0, 11).Value = Me.txtNumberOfEmployees.Text
    For Each cell In Range("EndMarkets")
        If cell.Value = Me.cmbEndMarket.Text Then
            comp.Offset(0, 12).Value = cell.Offset(0, -1).Value
        End If
    Next cell
    For Each cell In Range("ProductTypes")
        If cell.Value = Me.cmbProductType.Text Then
            comp.Offset(0, 13).Value = cell.Offset(0, -1).Value
        End If
    Next cell
    comp.Offset(0, 14).Value = Me.txtComments.Text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    quit = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
    
    'want to quit
    If quit = vbYes Then
        Cancel = 0
    End If
    
    'mistake
    If quit = vbNo Then
        Cancel = 1
    End If

End Sub
