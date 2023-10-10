VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ExcluirCelulasAcima()
    ' Defina o termo que voc� est� procurando
    Dim Termo As String
    Termo = "See all details"
    
    ' Percorra todas as c�lulas na coluna A
    Dim rng As Range
    For Each rng In ThisWorkbook.Sheets("Dados Extraidos").Range("A:A")
        ' Se a c�lula cont�m o termo desejado
        If rng.Value = Termo Then
            ' Exclua todas as c�lulas acima desta c�lula na coluna A
            Range("A1:A" & rng.Row - 1).Delete Shift:=xlUp
            Exit Sub
        End If
    Next rng
End Sub

Sub ManterUltimas15Celulas()
    ' Defina a coluna que voc� deseja modificar
    Dim Coluna As Range
    Set Coluna = ThisWorkbook.Sheets("Dados Extraidos").Columns("A")
    
    ' Encontre a �ltima c�lula n�o vazia na coluna
    Dim UltimaLinha As Long
    UltimaLinha = Coluna.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    ' Se h� mais de 15 c�lulas na coluna
    If UltimaLinha > 15 Then
        ' Exclua todas as c�lulas acima das �ltimas 15
        Range("A1:A" & UltimaLinha - 15).Delete Shift:=xlUp
    End If
End Sub
Sub TransposeRange()
    ' Defina o intervalo que voc� deseja transpor
    Dim SourceRange As Range
    Set SourceRange = ThisWorkbook.Sheets("Dados Extraidos").Range("A1:A16")
    
    ' Defina onde voc� deseja que o intervalo transposto comece
    Dim TargetCell As Range
    Set TargetCell = ThisWorkbook.Sheets("Planilha2").Range("B1")
    
    ' Transponha o intervalo
    SourceRange.Copy
    TargetCell.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    
    ' Limpe a �rea de transfer�ncia
    Application.CutCopyMode = False
End Sub
Sub ExcluirPrimeiras3Colunas()
    ' Defina a planilha onde voc� deseja excluir as colunas
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Planilha2")
    
    ' Exclua as tr�s primeiras colunas
    ws.Columns("A:C").Delete Shift:=xlToLeft
End Sub
