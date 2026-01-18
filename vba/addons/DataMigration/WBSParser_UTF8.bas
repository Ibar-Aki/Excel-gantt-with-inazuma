Attribute VB_Name = "WBSParser"
Option Explicit

' ==========================================
'  WBS番号解析モジュール
' ==========================================
' WBS番号（例: "1", "1.1", "1.1.1", "1.1.1.1"）から
' 階層レベル（LV1〜LV4）を判定するコアエンジン

' ==========================================
'  WBS番号から階層レベルを取得
' ==========================================
' 引数:
'   wbsText - WBS番号（例: "1.1.1.1"）
' 戻り値:
'   1〜4 (LV1〜LV4)
'   0 (無効なWBS番号の場合)
' ==========================================
Public Function ParseWBSLevel(ByVal wbsText As String) As Long
    On Error GoTo ErrorHandler
    
    ' 空文字チェック
    If Trim$(wbsText) = "" Then
        ParseWBSLevel = 0
        Exit Function
    End If
    
    ' ドット区切りでレベルを判定
    Dim dotCount As Long
    dotCount = CountChar(wbsText, ".")
    
    ' レベル = ドット数 + 1
    Dim level As Long
    level = dotCount + 1
    
    ' LV1〜LV4の範囲チェック
    If level < 1 Or level > 4 Then
        ParseWBSLevel = 0
        Exit Function
    End If
    
    ' フォーマット妥当性チェック
    If Not ValidateWBS(wbsText) Then
        ParseWBSLevel = 0
        Exit Function
    End If
    
    ParseWBSLevel = level
    Exit Function
    
ErrorHandler:
    ParseWBSLevel = 0
End Function

' ==========================================
'  階層判定モード定義
' ==========================================
Public Enum HierarchyMode
    Mode_WBS = 0       ' WBS形式 (1.1.1)
    Mode_Level = 1     ' レベル数値 (1, 2, 3...)
End Enum

' ==========================================
'  汎用階層レベル判定
' ==========================================
' 引数:
'   text - 解析対象の文字列
'   mode - 判定モード
' 戻り値:
'   1〜4 (レベル)
'   0 (無効)
' ==========================================
Public Function ParseHierarchyLevel(ByVal text As String, ByVal mode As HierarchyMode) As Long
    On Error GoTo ErrorHandler
    
    If mode = Mode_WBS Then
        ParseHierarchyLevel = ParseWBSLevel(text)
    ElseIf mode = Mode_Level Then
        ParseHierarchyLevel = ParseDirectLevel(text)
    Else
        ParseHierarchyLevel = 0
    End If
    Exit Function
    
ErrorHandler:
    ParseHierarchyLevel = 0
End Function

' ==========================================
'  レベル数値の直接判定
' ==========================================
Private Function ParseDirectLevel(ByVal text As String) As Long
    On Error GoTo ErrorHandler
    
    If Trim$(text) = "" Then
        ParseDirectLevel = 0
        Exit Function
    End If
    
    If IsNumeric(text) Then
        Dim val As Long
        val = CLng(text)
        If val >= 1 And val <= 4 Then
            ParseDirectLevel = val
            Exit Function
        End If
    End If
    
    ParseDirectLevel = 0
    Exit Function
    
ErrorHandler:
    ParseDirectLevel = 0
End Function

' ==========================================
'  WBS番号の妥当性チェック
' ==========================================
' 引数:
'   wbsText - WBS番号（例: "1.1.1.1"）
' 戻り値:
'   True - 妥当なWBS番号
'   False - 無効なWBS番号
' ==========================================
Public Function ValidateWBS(ByVal wbsText As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 空文字チェック
    If Trim$(wbsText) = "" Then
        ValidateWBS = False
        Exit Function
    End If
    
    ' トリム
    wbsText = Trim$(wbsText)
    
    ' ドットで分割
    Dim parts() As String
    parts = Split(wbsText, ".")
    
    ' 各パートが数値かチェック
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If Not IsNumeric(parts(i)) Then
            ValidateWBS = False
            Exit Function
        End If
        
        ' 数値が1以上かチェック
        If CLng(parts(i)) < 1 Then
            ValidateWBS = False
            Exit Function
        End If
    Next i
    
    ValidateWBS = True
    Exit Function
    
ErrorHandler:
    ValidateWBS = False
End Function

' ==========================================
'  文字列内の特定文字の出現回数をカウント
' ==========================================
Private Function CountChar(ByVal text As String, ByVal char As String) As Long
    Dim count As Long
    Dim i As Long
    count = 0
    
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = char Then
            count = count + 1
        End If
    Next i
    
    CountChar = count
End Function

' ==========================================
'  WBS番号からタスク番号を抽出
' ==========================================
' 引数:
'   wbsText - WBS番号（例: "1.2.3.4"）
'   level - 取得したいレベル（1〜4）
' 戻り値:
'   指定レベルのタスク番号
'   例: ParseWBSNumber("1.2.3.4", 2) → 2
' ==========================================
Public Function ParseWBSNumber(ByVal wbsText As String, ByVal level As Long) As Long
    On Error GoTo ErrorHandler
    
    If Not ValidateWBS(wbsText) Then
        ParseWBSNumber = 0
        Exit Function
    End If
    
    Dim parts() As String
    parts = Split(wbsText, ".")
    
    ' レベルが範囲外
    If level < 1 Or level > UBound(parts) + 1 Then
        ParseWBSNumber = 0
        Exit Function
    End If
    
    ParseWBSNumber = CLng(parts(level - 1))
    Exit Function
    
ErrorHandler:
    ParseWBSNumber = 0
End Function
