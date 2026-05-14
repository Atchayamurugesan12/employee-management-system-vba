VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Employee Database"
   ClientHeight    =   5030
   ClientLeft      =   30
   ClientTop       =   370
   ClientWidth     =   5140
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

cmbdept.AddItem "Finance"
cmbdept.AddItem "Sales"
cmbdept.AddItem "Marketing"
cmbdept.AddItem "Customer support"
cmbdept.AddItem "Administration"
cmbdept.AddItem "Procurement"
cmbdept.AddItem "Logistics"
cmbdept.AddItem "Quality Assurance"
cmbdept.AddItem "Training and development"

tglstatus.Caption = "Inactive"

End Sub
Private Sub tglstatus_Click()

    If tglstatus.Value = True Then
        tglstatus.Caption = "Active"
    Else
        tglstatus.Caption = "Inactive"
    End If
    
End Sub

Private Sub cmdnew_Click()

Dim maxid As Long
Dim cell As Range

On Error GoTo errorhandler

maxid = 0

For Each cell In Sheets("sheet1").Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
    If cell.Value <> "" Then
        If Val(Mid(cell.Value, 4)) > maxid Then
            maxid = Val(Mid(cell.Value, 4))
        End If
    End If
Next cell
    
txtid.Value = "EMP" & Format(maxid + 1, "000")
Exit Sub

errorhandler:
    MsgBox "Error in ID generation: " & Err.Description
    

End Sub
Private Sub cmdsubmit_Click()

Dim gender As String
Dim lastrow As Long

On Error GoTo errorhandler

'Gender validation
    If optmale.Value = True Then
        gender = "Male"
    ElseIf optfemale.Value = True Then
        gender = "Female"
    End If

'Name validation
    If txtname.Value = "" Then
        MsgBox "Enter your name"
        Exit Sub
    End If
    
'Salary validation
    If Not IsNumeric(txtsalary.Value) Then
        MsgBox "Enter valid salary"
        Exit Sub
    End If
    
lastrow = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row + 1

    Sheets("Sheet1").Cells(lastrow, 1).Value = txtid.Value
    Sheets("Sheet1").Cells(lastrow, 2).Value = txtname.Value
    Sheets("Sheet1").Cells(lastrow, 3).Value = cmbdept.Value
    Sheets("Sheet1").Cells(lastrow, 4).Value = txtjoindt.Value
    Sheets("Sheet1").Cells(lastrow, 5).Value = txtsalary.Value
    Sheets("Sheet1").Cells(lastrow, 6).Value = gender
    Sheets("Sheet1").Cells(lastrow, 7).Value = tglstatus.Caption
    
MsgBox "Data added"

txtid.Value = ""
Exit Sub

errorhandler:
    MsgBox "Error in submit: " & Err.Description
    
    
End Sub
Private Sub cmdsearch_Click()

Dim foundcell As Range

    If txtid.Value = "" Then
        MsgBox "Enter Employee ID"
    Exit Sub
    End If
    
On Error GoTo errorhandler

    
Set foundcell = Sheets("sheet1").Range("A:A").Find(txtid.Value, LookAt:=xlWhole)

    If Not foundcell Is Nothing Then
        txtname.Value = Sheets("sheet1").Cells(foundcell.Row, 2).Value
        cmbdept.Value = Sheets("sheet1").Cells(foundcell.Row, 3).Value
        txtjoindt.Value = Sheets("sheet1").Cells(foundcell.Row, 4).Value
        txtsalary.Value = Sheets("sheet1").Cells(foundcell.Row, 5).Value
        
        If Sheets("Sheet1").Cells(foundcell.Row, 6).Value = "Male" Then
            optmale.Value = True
            optfemale.Value = False
        Else
            optfemale.Value = True
            optmale.Value = False
        End If
        
        
        tglstatus.Caption = Sheets("sheet1").Cells(foundcell.Row, 7).Value
    Else
        MsgBox "Not Found"
    End If
Exit Sub

errorhandler:
    MsgBox "Error in search: " & Err.Description
    
    
End Sub
Private Sub cmddlt_Click()

Dim foundcell As Range

    If txtid.Value = "" Then
        MsgBox "Enter Employee ID"
    Exit Sub
    End If
    
On Error GoTo errorhandler
    
Set foundcell = Sheets("sheet1").Range("A:A").Find(txtid.Value, LookAt:=xlWhole)

    If Not foundcell Is Nothing Then
        Sheets("sheet1").Rows(foundcell.Row).Delete
            MsgBox "Record deleted"
    Else
            MsgBox "Record Not Found"
    End If
Exit Sub

errorhandler:
    MsgBox "Error in Delete: " & Err.Description
    
    
End Sub
Private Sub cmdclear_Click()

txtid.Value = ""
txtname.Value = ""
cmbdept.Value = ""
txtjoindt.Value = ""
txtsalary.Value = ""

optmale.Value = False
optfemale.Value = False

tglstatus.Caption = "Inactive"

End Sub
