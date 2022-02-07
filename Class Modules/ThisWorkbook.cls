VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)
updateProperty "userId", 0
updateProperty "isUserLogged", False
End Sub

Private Sub Workbook_Open()
initializeObjects
updateProperty "userId", 0
updateProperty "isUserLogged", False
Application.Caption = "Loaded week " & w & "|" & y
updateProperty "import path", "K:\Common\Bilans Mas\Tygodniowo\" & ThisWorkbook.CustomDocumentProperties("yearLoaded").Value & "\Tydzień " & ThisWorkbook.CustomDocumentProperties("weekLoaded").Value
End Sub

Private Sub Workbook_Deactivate()

On Error Resume Next

With Application
    .CommandBars("Cell").Controls("Remove connection").Delete
    .CommandBars("Cell").Controls("Cancel order").Delete
    .CommandBars("Cell").Controls("Check missing connections").Delete
    .CommandBars("Cell").Controls("Correct order").Delete
End With

On Error GoTo 0

End Sub

Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)

Dim cBut As CommandBarButton

On Error Resume Next

With Application

    .CommandBars("Cell").Controls("Remove connection").Delete

    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)

End With



With cBut

   .Caption = "Remove connection"

   .Style = msoButtonCaption

   .OnAction = "removeConnection"

End With

With Application

    .CommandBars("Cell").Controls("Cancel order").Delete

    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)

End With



With cBut

   .Caption = "Cancel order"

   .Style = msoButtonCaption

   .OnAction = "cancelOrder"

End With

With Application

    .CommandBars("Cell").Controls("Check missing connections").Delete

    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)

End With



With cBut

   .Caption = "Check missing connections"

   .Style = msoButtonCaption

   .OnAction = "checkMissingConnections"

End With


With Application

    .CommandBars("Cell").Controls("Correct order").Delete

    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)

End With



With cBut

   .Caption = "Correct order"

   .Style = msoButtonCaption

   .OnAction = "correctOrder"

End With

With Application

    .CommandBars("Cell").Controls("Missing connections By Session number").Delete

    Set cBut = .CommandBars("Cell").Controls.Add(Temporary:=True)

End With



With cBut

   .Caption = "Missing connections By Session number"

   .Style = msoButtonCaption

   .OnAction = "MissingConnectionsBySessionNumber"

End With

On Error GoTo 0


End Sub
