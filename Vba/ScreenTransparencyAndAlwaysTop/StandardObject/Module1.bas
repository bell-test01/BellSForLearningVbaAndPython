Attribute VB_Name = "Module1"
Option Explicit

'==============================
' API定義
'==============================
#If VBA7 Then
    Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
    
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
    
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetModuleBaseName Lib "psapi.dll" Alias "GetModuleBaseNameA" (ByVal hProcess As LongPtr, ByVal hModule As LongPtr, ByVal lpBaseName As String, ByVal nSize As Long) As Long

    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal uFlags As Long _
    ) As Long
#End If

Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_VM_READ As Long = &H10

Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = &H2


Private Const HWND_TOPMOST As LongPtr = -1
Private Const HWND_NOTOPMOST As LongPtr = -2

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40

'==============================
' 変数
'==============================
Public WindowList As Collection
Private ProcCache As Object

'==============================
' ウィンドウ列挙
'==============================
Sub GetWindowList()
    Set WindowList = New Collection
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Function EnumWindowsProc(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long
    
    Dim buff As String * 255
    Dim title As String
    Dim length As Long
    
    If IsWindowVisible(hwnd) <> 0 Then
        length = GetWindowText(hwnd, buff, 255)
        title = Left(buff, length)
        
        If title <> "" Then
            WindowList.Add hwnd & "|" & title
        End If
    End If
    
    EnumWindowsProc = 1
End Function

'==============================
' プロセス情報
'==============================
Function GetProcessId(hwnd As LongPtr) As Long
    Dim pid As Long
    GetWindowThreadProcessId hwnd, pid
    GetProcessId = pid
End Function

Function GetProcessName(hwnd As LongPtr) As String
    
    Dim pid As Long
    Dim hProcess As LongPtr
    Dim buff As String * 260
    Dim length As Long
    
    GetWindowThreadProcessId hwnd, pid
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pid)
    
    If hProcess <> 0 Then
        length = GetModuleBaseName(hProcess, 0, buff, 260)
        If length > 0 Then
            GetProcessName = Left(buff, length)
        End If
        CloseHandle hProcess
    End If

End Function

' 高速化キャッシュ
Function GetProcessNameCached(hwnd As LongPtr) As String

    If ProcCache Is Nothing Then
        Set ProcCache = CreateObject("Scripting.Dictionary")
    End If
    
    Dim pid As Long
    pid = GetProcessId(hwnd)
    
    If ProcCache.Exists(pid) Then
        GetProcessNameCached = ProcCache(pid)
    Else
        Dim name As String
        name = GetProcessName(hwnd)
        ProcCache(pid) = name
        GetProcessNameCached = name
    End If

End Function

'==============================
' 一覧出力
'==============================
Sub ウィンドウ一覧取得_完全版()

    Dim i As Long
    
    GetWindowList
    
    With ActiveSheet
        .Range("A2:E" & .Rows.Count).ClearContents
        
        For i = 1 To WindowList.Count
            
            Dim parts() As String
            parts = Split(WindowList(i), "|")
            
            Dim hwnd As LongPtr
            hwnd = CLngPtr(parts(0))
            
            .Cells(i + 1, 1).Value = hwnd
            .Cells(i + 1, 2).Value = parts(1)
            .Cells(i + 1, 3).Value = GetProcessId(hwnd)
            .Cells(i + 1, 4).Value = GetProcessNameCached(hwnd)
            
        Next i
        
        .Range("A1:E1").Value = Array("hwnd", "タイトル", "PID", "プロセス名", "透過度")
    End With

End Sub

'==============================
' 透過処理
'==============================
Sub SetTransparency(ByVal hwnd As LongPtr, ByVal alpha As Byte)

    Dim style As Long
    
    style = GetWindowLong(hwnd, GWL_EXSTYLE)
    SetWindowLong hwnd, GWL_EXSTYLE, style Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, alpha, LWA_ALPHA

End Sub

Sub SetTopMost(ByVal hwnd As LongPtr)

    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

End Sub

Sub UnsetTopMost(ByVal hwnd As LongPtr)

    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

End Sub
