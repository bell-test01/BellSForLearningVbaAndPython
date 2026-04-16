VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Public Sub BtnGetWindowInfo()
    Application.EnableEvents = False
    Me.Range("A2:E" & Me.Rows.Count).ClearContents
    ウィンドウ一覧取得_完全版
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ExitHandler
    
    Dim cell As Range
    
    Application.EnableEvents = False
    
    ' --- 透過処理（既存） ---
    If Not Intersect(Target, Me.Columns("E")) Is Nothing Then
        
        For Each cell In Target
            If cell.Row > 1 Then
                If IsNumeric(cell.Value) Then
                    If cell.Value >= 0 And cell.Value <= 255 Then
                        SetTransparency Me.Cells(cell.Row, 1).Value, CByte(cell.Value)
                    End If
                End If
            End If
        Next cell
        
    End If
    
    ' --- 最前面処理（追加） ---
    If Not Intersect(Target, Me.Columns("F")) Is Nothing Then
        
        For Each cell In Target
            If cell.Row > 1 Then
                
                Dim hwnd As LongPtr
                hwnd = Me.Cells(cell.Row, 1).Value
                
                If cell.Value = 1 Then
                    SetTopMost hwnd
                Else
                    UnsetTopMost hwnd
                End If
                
            End If
        Next cell
        
    End If

ExitHandler:
    Application.EnableEvents = True

End Sub
