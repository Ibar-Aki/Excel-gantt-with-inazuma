Attribute VB_Name = "ErrorHandler"
Option Explicit

' ==========================================
'  エラーハンドリング改善モジュール
' ==========================================
' 統一されたエラー処理とログ機能を提供
' ==========================================

' エラーログの保存場所
Private Const ERROR_LOG_FILE As String = "InazumaGantt_ErrorLog.txt"

' ログレベル
Public Enum LogLevel
    LOG_DEBUG = 1
    LOG_INFO = 2
    LOG_WARNING = 3
    LOG_ERROR = 4
    LOG_CRITICAL = 5
End Enum

' ==========================================
'  統一エラーハンドラー
' ==========================================
Public Sub HandleError(ByVal moduleName As String, ByVal procedureName As String, Optional ByVal userMessage As String = "")
    Dim errNumber As Long
    Dim errDescription As String
    Dim errSource As String
    Dim logMessage As String
    
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    
    ' ログメッセージ作成
    logMessage = "[ERROR] " & Now & vbCrLf & _
                 "Module: " & moduleName & vbCrLf & _
                 "Procedure: " & procedureName & vbCrLf & _
                 "Error #" & errNumber & ": " & errDescription & vbCrLf & _
                 "Source: " & errSource
    
    ' デバッグ出力
    Debug.Print logMessage
    
    ' ファイルにログ出力
    Call WriteLog(logMessage, LOG_ERROR)
    
    ' ユーザーへのメッセージ
    If userMessage = "" Then
        userMessage = "処理中に問題が発生しました。"
    End If
    
    Dim displayMessage As String
    displayMessage = userMessage & vbCrLf & vbCrLf & _
                     "エラー情報:" & vbCrLf & _
                     "コード: ERR" & errNumber & vbCrLf & _
                     "詳細はログファイルを確認してください。"
    
    MsgBox displayMessage, vbCritical, "エラー - " & moduleName
    
    ' エラーをクリア
    Err.Clear
End Sub

' ==========================================
'  ログ出力
' ==========================================
Public Sub WriteLog(ByVal message As String, Optional ByVal level As LogLevel = LOG_INFO)
    On Error Resume Next
    
    Dim logFilePath As String
    Dim fileNum As Integer
    Dim levelText As String
    
    ' ログレベルのテキスト
    Select Case level
        Case LOG_DEBUG: levelText = "DEBUG"
        Case LOG_INFO: levelText = "INFO"
        Case LOG_WARNING: levelText = "WARNING"
        Case LOG_ERROR: levelText = "ERROR"
        Case LOG_CRITICAL: levelText = "CRITICAL"
        Case Else: levelText = "UNKNOWN"
    End Select
    
    ' ログファイルパス（Excelファイルと同じフォルダ）
    logFilePath = ThisWorkbook.Path & "\" & ERROR_LOG_FILE
    
    ' ファイルに追記
    fileNum = FreeFile
    Open logFilePath For Append As #fileNum
    Print #fileNum, "[" & levelText & "] " & message
    Print #fileNum, String(80, "-")
    Close #fileNum
    
    ' デバッグ出力
    Debug.Print "[" & levelText & "] " & message
End Sub

' ==========================================
'  入力値検証: 必須チェック
' ==========================================
Public Function ValidateRequired(ByVal value As Variant, ByVal fieldName As String) As Boolean
    If IsEmpty(value) Or Trim$(CStr(value)) = "" Then
        MsgBox fieldName & " は必須項目です。", vbExclamation, "入力エラー"
        ValidateRequired = False
    Else
        ValidateRequired = True
    End If
End Function

' ==========================================
'  入力値検証: 数値チェック
' ==========================================
Public Function ValidateNumeric(ByVal value As Variant, ByVal fieldName As String, Optional ByVal minValue As Double = -999999, Optional ByVal maxValue As Double = 999999) As Boolean
    If Not IsNumeric(value) Then
        MsgBox fieldName & " は数値で入力してください。", vbExclamation, "入力エラー"
        ValidateNumeric = False
        Exit Function
    End If
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    If numValue < minValue Or numValue > maxValue Then
        MsgBox fieldName & " は " & minValue & " から " & maxValue & " の範囲で入力してください。", vbExclamation, "入力エラー"
        ValidateNumeric = False
        Exit Function
    End If
    
    ValidateNumeric = True
End Function

' ==========================================
'  入力値検証: 日付チェック
' ==========================================
Public Function ValidateDate(ByVal value As Variant, ByVal fieldName As String) As Boolean
    If Not IsDate(value) Then
        MsgBox fieldName & " は有効な日付を入力してください。", vbExclamation, "入力エラー"
        ValidateDate = False
    Else
        ValidateDate = True
    End If
End Function

' ==========================================
'  進捗バー表示（長時間処理用）
' ==========================================
Public Sub ShowProgress(ByVal currentStep As Long, ByVal totalSteps As Long, ByVal message As String)
    Dim percentage As Double
    percentage = (currentStep / totalSteps) * 100
    
    Application.StatusBar = message & " (" & Format(percentage, "0") & "%)"
    DoEvents
End Sub

' ==========================================
'  進捗バークリア
' ==========================================
Public Sub ClearProgress()
    Application.StatusBar = False
End Sub
