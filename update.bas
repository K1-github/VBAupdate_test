Attribute VB_Name = "Main"
Option Explicit
'Github rawなどUTF-8からShift-JISに変更が必要な場合
Sub UpdateVBA1()
    Dim http As Object
    Dim url As String
    Dim vbaFilePath As String
    Dim sjisFilePath As String
    Dim fso As Object
    Dim targetModule As Object
    Dim stream As Object
    Dim moduleName As String
    Dim confirmDelete As VbMsgBoxResult

    ' ? アップデートするモジュール名
    moduleName = "Main"

    ' ?? GitHub の `raw` ファイルのURL
    url = "https://github.com/K1-github/VBAupdate_test/raw/refs/heads/main/update.bas"
    

    ' ?? HTTPリクエストを送る（WinHttpを使用）
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.Send

    ' ?? ステータスチェック
    If http.Status <> 200 Then
        MsgBox "エラー: サーバーからデータを取得できませんでした！" & vbNewLine & "HTTP ステータス: " & http.Status, vbCritical
        Exit Sub
    End If

    ' ? 一時ファイルのパス（UTF-8でダウンロード）
    vbaFilePath = Environ("TEMP") & "\update.bas"
    sjisFilePath = Environ("TEMP") & "\update_sjis.bas"

    ' ?? `ResponseBody` を `.bas` に保存
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' adTypeBinary（バイナリモード）
    stream.Open
    stream.Write http.ResponseBody
    stream.SaveToFile vbaFilePath, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing

    ' ?? UTF-8 → Shift-JIS に変換
    ConvertUTF8ToSJI vbaFilePath, sjisFilePath

    ' ?? 既存のモジュールがあるか確認
    On Error Resume Next
    Set targetModule = ThisWorkbook.VBProject.VBComponents(moduleName)
    On Error GoTo 0

    ' ?? 既存のモジュールがある場合、削除前に確認
    If Not targetModule Is Nothing Then
        confirmDelete = MsgBox("既存のモジュール [" & moduleName & "] を削除して、新しいバージョンに更新しますか？", vbYesNo + vbQuestion, "VBAアップデート")

        ' ? ユーザーが「No」を選んだ場合、処理を中止
        If confirmDelete = vbNo Then
            MsgBox "アップデートをキャンセルしました。", vbInformation
            Exit Sub
        End If

        ' ?? 既存のモジュールを削除
        ThisWorkbook.VBProject.VBComponents.Remove targetModule
    End If

    ' ?? 新しいVBAコードをインポート（Shift-JIS 変換済み）　エラーが出る場合、オプション設定で許可する必要あり
    ThisWorkbook.VBProject.VBComponents.Import sjisFilePath

    ' ?? 一時ファイルを削除
    On Error Resume Next
    Kill vbaFilePath
    Kill sjisFilePath
    On Error GoTo 0

    ' ? 完了メッセージ
    MsgBox "VBAの更新が完了しました！", vbInformation

End Sub

' ?? UTF-8 から Shift-JIS に変換する関数
Sub ConvertUTF8ToSJI(ByVal utf8File As String, ByVal sjisFile As String)
    Dim stream As Object
    Dim textData As String
    
    ' UTF-8 のファイルを読み込む
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile utf8File
    textData = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' **改行コードを CRLF に修正**
    textData = FixCRLF(textData)

    ' **Shift-JIS で保存**
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "Shift_JIS"
    stream.Open
    stream.WriteText textData
    stream.SaveToFile sjisFile, 2 ' adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
End Sub

Function FixCRLF(ByVal textData As String) As String
    ' **すべての `CRLF` を `LF` に統一**
    textData = Replace(textData, vbCrLf, vbLf)
    ' **すべての `CR` を削除（単独の `CR` がある場合の対策）**
    textData = Replace(textData, vbCr, "")
    ' **`LF` を `CRLF` に変換（これで余分な `CR` を防ぐ）**
    FixCRLF = Replace(textData, vbLf, vbCrLf)
End Function

