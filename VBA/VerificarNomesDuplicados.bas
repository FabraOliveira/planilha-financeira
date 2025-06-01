VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim intervalo As Range
    Dim nomeVerificar As String
    Dim contador As Long
    
    ' Defina o intervalo que ser� verificado para duplicados
    Set intervalo = Me.Range("J8:J29")
    
    ' Verifica se a c�lula alterada est� dentro do intervalo
    If Not Intersect(Target, intervalo) Is Nothing Then
        Application.EnableEvents = False
        If Target.Count = 1 Then ' Garante que apenas uma c�lula foi alterada
            nomeVerificar = Target.Value
            If nomeVerificar <> "" Then
                contador = Application.WorksheetFunction.CountIf(intervalo, nomeVerificar)
                If contador > 1 Then
                    MsgBox "O nome '" & nomeVerificar & "' j� est� na lista. Por favor, escolha outro nome.", vbExclamation, "Nome Duplicado"
                    Target.ClearContents ' Apaga apenas o valor rec�m digitado
                End If
            End If
        End If
        Application.EnableEvents = True
    End If
End Sub

