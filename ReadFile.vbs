'*********************************************************************
' Purpose: filePathで指定されたファイルをUTF-8、改行コードLFで読み込み、
'          取得した行単位で動的配列のdatatListに設定する
' Inputs: filePath:対象となるファイルパス
'         datatList:１行のデータを格納する動的リスト（参照渡し）
' Returns: なし
'*********************************************************************
Sub ReadFile(ByVal filePath, ByRef datatList)
  Dim input
  Set input = CreateObject("ADODB.Stream")
  input.Open                   ' Stream オブジェクトを開く
  input.Type = 2               ' ★ポイント１
  input.Charset = "UTF-8"      ' ★ポイント２
  input.LineSeparator = 10     ' ★ポイント３
  input.LoadFromFile filePath  ' ★ポイント４

  ' 対象ファイルから1行ずつ読み込む
  Dim line
  Dim aryStrings
  Do Until input.EOS
    line = input.ReadText(-2)    ' ★ポイント５
    If InStr(1, line, "#") = 1 Then
      ' WScript.echo("COMMENT.")
    Else
      dataList.add line
    End If
  Loop

  ' Stream を閉じる
  input.Close
End Sub
