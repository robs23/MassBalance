VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dbFunctions 
   Caption         =   "Funkcje bazy danych"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   OleObjectBlob   =   "dbFunctions.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dbFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FunctionService As clsFunctionService

Private Sub btnOutput2Input_Click()
Dim cOrder As clsOrder

If FunctionService.ReturnedOrders.Count > 0 Then
    Dim str As String
    For Each cOrder In FunctionService.ReturnedOrders
        str = str & cOrder.sapId & ","
    Next cOrder
    If Len(str) > 0 Then str = Left(str, Len(str) - 1)
    
    txtIn.Value = str
Else
    MsgBox "Brak zleceń do przeniesienia", vbExclamation + vbOKOnly, "Brak zleceń"
End If
End Sub

Private Sub btnSend_Click()
Dim chosenId As Integer
Dim parameters As String
Dim cFunction As clsFunction

chosenId = cmbFunctions.ListIndex
parameters = txtIn.Value

If chosenId >= 0 Then
    If Len(parameters) > 0 Then
        For Each cFunction In FunctionService.Functions
            If cFunction.Id = chosenId Then
                Set FunctionService.ChosenFunction = cFunction
                Set FunctionService.ReturnedOrders = FunctionService.ChosenFunction.Execute(parameters)
                txtOut.Value = FunctionService.ChosenFunction.Output
                Exit For
            End If
        Next cFunction
    Else
        MsgBox "Wpisz parametry zapytania", vbCritical + vbOKOnly, "Brak parametrów"
    End If
Else
    MsgBox "Nie wybrano funkcji", vbCritical + vbOKOnly, "Wybierz funkcje"
End If



End Sub

Private Sub cmbFunctions_Change()
Dim chosenId As Integer
Dim cFunction As clsFunction

chosenId = cmbFunctions.ListIndex

If chosenId >= 0 Then
    For Each cFunction In FunctionService.Functions
        If cFunction.Id = chosenId Then
            lblHint.Caption = cFunction.Hint
            Exit For
        End If
    Next cFunction
End If

End Sub

Private Sub UserForm_Initialize()
    Set FunctionService = New clsFunctionService
    Dim cFunction As clsFunction
    
    For Each cFunction In FunctionService.Functions
        cmbFunctions.AddItem cFunction.Name, cFunction.Id
    Next cFunction

End Sub

