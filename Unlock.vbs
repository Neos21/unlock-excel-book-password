Option Explicit

' ================================================================================
' Unlock Excel Book Password
' ================================================================================

' 必ず CScript で実行させる
If Instr(LCase(WScript.FullName), "wscript") > 0 Then
  WScript.CreateObject("WScript.Shell").Run("CScript //NoLogo """ & WScript.ScriptFullName & """")
  WScript.Quit
End If

WScript.Echo "────────────────────"
WScript.Echo "Unlock Excel Book Password"
WScript.Echo ""

' 設定ファイル・解除パスワード出力ファイル
Const PROPERTY_FILE = "Property.txt"
Const UNLOCKED_FILE = "Unlocked.txt"

' 設定ファイルから取得して使うパラメータ
Dim fileName    : fileName    = ""
Dim min         : min         = ""
Dim max         : max         = ""
Dim strings     : strings     = "" 
Dim progressStr : progressStr = ""

' 設定ファイルを読み込んで変数に値を設定する
Call readPropertyFile()

' レジューム文字列まで到達したか
Dim isResumed : isResumed = False

' 探索に使用する文字列を配列にする
Dim strs : strs = Split(strings, ",")

' 探索に使用する文字列をカンマ抜きで連結し直す (レジューム判定に使用する)
Dim joinStrings : joinStrings = Join(strs, "")

' Excel の準備
Dim excel
On Error Resume Next
Set excel = GetObject(, "Excel.Application")
If Err.Number <> 0 Then
  Set excel = CreateObject("Excel.Application")
End If
On Error GoTo 0
excel.DisplayAlerts = False
excel.Visible = False

' ブルートフォース
Dim n
For n = min To max
  WScript.Echo "▼Next Length : " & n
  Call unlock("", n)
Next

' 見つからなかった…
WScript.Echo "────────────────────"
WScript.Echo "見つかりませんでした。"
WScript.Echo "────────────────────"
excel.Quit
Set excel = Nothing
WScript.Quit



' ================================================================================
' 設定ファイルを読み込む
' ================================================================================
Sub readPropertyFile()
  Dim fso : Set fso = WScript.CreateObject("Scripting.FileSystemObject")
  
  If fso.FileExists(PROPERTY_FILE) = False Then
    WScript.Echo PROPERTY_FILE & " が存在しないため処理できません。終了します。"
    WScript.Echo "────────────────────"
    Set fso = Nothing
    WScript.Quit
  End If
  
  Dim propertyFile : Set propertyFile = fso.OpenTextFile(PROPERTY_FILE)
  
  Do While propertyFile.AtEndOfStream <> True
    Dim line : line = propertyFile.ReadLine
    If fileName    = "" And regSearch(line, "fileName"   ) Then fileName    = getPropertyValue(line, "fileName")
    If min         = "" And regSearch(line, "min"        ) Then min         = getPropertyValue(line, "min")
    If max         = "" And regSearch(line, "max"        ) Then max         = getPropertyValue(line, "max")
    If strings     = "" And regSearch(line, "strings"    ) Then strings     = getPropertyValue(line, "strings")
    If progressStr = "" And regSearch(line, "progressStr") Then progressStr = getPropertyValue(line, "progressStr")
  Loop
  
  propertyFile.Close
  Set propertyFile = Nothing
  
  If fileName = "" Or min = "" Or max = "" Or strings = "" Then
    WScript.Echo "必須項目が定義されていないため処理できません。終了します。"
    WScript.Echo "────────────────────"
    Set fso = Nothing
    WScript.Quit
  ElseIf fso.FileExists(fileName) = False Then
    WScript.Echo fileName & " が存在しないため処理できません。終了します。"
    WScript.Echo "────────────────────"
    Set fso = Nothing
    WScript.Quit
  End If
  
  Set fso = Nothing
  
  WScript.Echo "以下の設定で処理を開始します。"
  WScript.Echo "  対象ブック名         : " & fileName
  WScript.Echo "  パスワード長最小桁   : " & min
  WScript.Echo "  パスワード長最大桁   : " & max
  WScript.Echo "  探索に使用する文字列 : " & strings
  WScript.Echo "  最後に探索した文字列 : " & progressStr
  
  ' 最後に探索した文字列があれば、その文字長を min に設定する
  If progressStr <> "" Then
    min = Len(progressStr)
    WScript.Echo "  → 最後に探索した文字列の長さ " & min & " 桁から再開します。"
  End If
  
  WScript.Echo "────────────────────"
  WScript.Echo ""
End Sub

' ================================================================================
' 設定項目を探索する
' ================================================================================
Function regSearch(testStr, propertyKey)
  Dim re : Set re = WScript.CreateObject("VBScript.RegExp")
  re.Pattern = "^" & propertyKey
  regSearch = re.test(testStr)
  Set re = Nothing
End Function

' ================================================================================
' 設定項目から値のみ取り出す
' ================================================================================
Function getPropertyValue(lineStr, propertyKey)
  ' "propertyName=Value" から "Value" を取り出すため "=" の次の文字以降を取得する
  getPropertyValue = Mid(lineStr, Len(propertyKey) + 2)
End Function

' ================================================================================
' ブルートフォース
' ================================================================================
Sub unlock(pass, length)
  On Error Resume Next
  
  ' 再帰の深さ
  Dim depth : depth = Len(pass) + 1
  
  ' インデント用スペース
  Dim t : t = String(depth * 2, " ")
  
  ' 状態不正
  If Len(pass) >= length Then
    WScript.Echo t & "◆Exit … pass:[" & pass & "], length:[" & length & "], depth:[" & depth & "]"
    Exit Sub
  End If
  
  ' 探索する文字列をループしていく
  Dim i
  For i = 0 To UBound(strs)
    ' Continue 処理用の Do : Loop Until 1
    Do
      ' 次の文字を1文字取り出す
      Dim p : p = strs(i)
      
      ' 既存の固定文字列と結合する
      Dim passStr : passStr = pass & p
      
      ' レジューム処理
      If progressStr <> "" And n <= min And isResumed = False Then
        ' 探索文字 : 最後に探索した文字列から深さに応じた1文字を取り出す
        Dim progressChar : progressChar = Mid(progressStr, depth, 1)
        
        ' 探索文字が strs の何番目にあるか探す (なければ -1 になる)
        Dim indexOf : indexOf = InStr(joinStrings, progressChar) - 1
        
        WScript.Echo t & "□Resume … passStr:[" & passStr & "], length:[" & length & "], depth:[" & depth & "], i:[" & i & "], progressChar:[" & progressChar & "], indexOf:[" & indexOf & "]"
        
        ' 探索に使用する文字列のループが探索文字の位置まで到達していなかったら、その文字は既に探索したものとして Continue する
        If i < indexOf Then
          WScript.Echo t & "←Continue"
          Exit Do
        End If
        
        ' 最後に探索した文字列と一致したらレジューム完了
        If passStr = progressStr Then
          isResumed = True
          WScript.Echo t & "☆Resumed … passStr:[" & passStr & "]"
          Exit Do
        End If
      Else
        ' 探索済文字なし・もしくはレジューム後
        WScript.Echo t & "■Progress … passStr:[" & passStr & "], length:[" & length & "], depth:[" & depth & "], i:[" & i & "]"
      End If
      
      If Len(passStr) <> length Then
        Call unlock(passStr, length)
      Else
        WScript.Echo t & "・" & passStr
        Err.Clear
        
        ' ブックを開いてみる (0:外部参照リンク更新なし・False:読取専用でなく編集モードで開く・5:ファイルの区切り文字なし)
        excel.Workbooks.Open fileName, 0, False, 5, passStr
        
        If Err.Number = 0 Then
          ' パスワード解読
          unlocked(passStr)
        ElseIf Err.Number <> 1004 Then
          WScript.Echo "────────────────────"
          WScript.Echo "◆予期しないエラーのため終了します : " & Err.Description
          WScript.Echo "────────────────────"
          excel.Quit
          Set excel = Nothing
          WScript.Quit
        End If
      End If
    Loop Until 1
  Next
End Sub

' ================================================================================
' パスワード解読時の処理
' ================================================================================
Sub unlocked(passStr)
  WScript.Echo "────────────────────"
  WScript.Echo "★パスワード解除 : [" & passStr & "]"
  WScript.Echo "────────────────────"
  
  ' Excel を表示する
  excel.Visible = True
  
  Dim fso : Set fso = WScript.CreateObject("Scripting.FileSystemObject")
  
  ' パスワード解除時のテキストファイル (2:書込専用・True:新規ファイル作成)
  Dim unlockedFile
  Set unlockedFile = fso.OpenTextFile(UNLOCKED_FILE, 2, True)
  
  ' ファイル名とパスワードを書き込む
  unlockedFile.WriteLine(fileName)
  unlockedFile.WriteLine(passStr)
  
  unlockedFile.Close
  Set unlockedFile = Nothing
  Set fso = Nothing
  WScript.Quit
End Sub