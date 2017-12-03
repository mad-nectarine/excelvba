Sub DoTranslation()
    
    Dim fromLang As String
    Dim toLang As String
    Dim apiKey As String
    
    'パラメーター取得
    fromLang = Application.ActiveWorkbook.Sheets(1).Cells(1, 2).text
    toLang = Application.ActiveWorkbook.Sheets(1).Cells(2, 2).text
    apiKey = Application.ActiveWorkbook.Sheets(1).Cells(3, 2).text

    '翻訳実行
    TranslateWithCS fromLang, toLang, apiKey
    
End Sub

Sub TranslateWithCS(fromLang As String, toLang As String, apiKey As String)

    '実行元Bookを退避
    Dim hostBook As Workbook
    Set hostBook = Application.ActiveWorkbook
    
    '対象ファイルを選択させて取得
    Dim filePath As String
    filePath = Application.GetOpenFilename("Excel File,*.xlsx", Title:="対象ファイルの選択")
    
    'キャンセルされたら終わり
    If filePath = "False" Then
        Exit Sub
    End If
    
    'Workbookとして読み込む
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=filePath)
    
    'セル値を入れておく配列
    Dim v() As Variant
    ReDim v(2, 0)
    
    '翻訳用結合文字列
    Dim txt() As Variant
    Dim totalTxtLength As Integer
    ReDim txt(0)
    totalTxtLength = 0
    
    Dim s As Worksheet
    For Each s In wb.Worksheets
        
        '読み込むセル範囲を判定
        Dim firstCell As Range
        Dim lastCell As Range
        Dim sheetAllRange As Range
        
        Set firstCell = s.Cells(1, 1)
        Set lastCell = s.Cells.SpecialCells(xlLastCell)
        Set sheetAllRange = Range(firstCell, lastCell)
        
        Dim c As Range
        For Each c In sheetAllRange.Cells
        
            '翻訳対象に相応しいか検証する
            'ここはcontinueが無くて悲しくなった・・・
            Dim valid As Boolean
            Dim var As Variant
            Dim ct As String
            
            valid = True
            var = c.Value
            ct = ""
            
            '値が無ければ無視
            If valid = True And IsEmpty(var) Then
                valid = False
            End If
            
            '数字と日付は無視
            If valid = True And IsNumeric(c.Value) Then
                valid = False
            End If
            If valid = True And IsDate(c.Value) Then
                valid = False
            End If
            
            '改行コードと空白を排除
            If valid = True Then
                ct = Replace(Trim(c.text), vbLf, " ")
            End If
            
            '空なら無視
            If valid = True And ct = "" Then
                valid = False
            End If
            
            '文字列長が1なら無視
            If valid = True And Len(ct) = 1 Then
                valid = False
            End If
            
            '翻訳する価値がある場合
            If valid = True Then
            
                '翻訳用結合文字列の配列を良い感じに初期化
                If IsEmpty(txt(UBound(txt))) Then
                    txt(UBound(txt)) = ""
                End If
                
                If txt(UBound(txt)) <> "" And LenB(txt(UBound(txt)) + ct) > 2000 Then
                    ReDim Preserve txt(UBound(txt) + 1)
                    txt(UBound(txt)) = ""
                End If
            
                '配列に値をセット
                '2次元配列は最終次元しか拡張できないので、列→1次元、行→2次元で書き込んで拡張する
                v(0, UBound(v, 2)) = c.Worksheet.Name + "!" + c.Address
                v(1, UBound(v, 2)) = ct
                txt(UBound(txt)) = txt(UBound(txt)) + ct + Chr(13) + Chr(10)
                
                totalTxtLength = totalTxtLength + LenB(ct) + 2
                ReDim Preserve v(2, UBound(v, 2) + 1)
            
            End If
        Next c
        
    Next s
    
    '対象が無ければ開いたBookを閉じて終わり
    If UBound(v, 2) = 0 And IsEmpty(v(0, 0)) Then
        MsgBox "変換対象がありませんでした"
        wb.Close (False)
        Exit Sub
    End If
    
    '不要な配列を削っておく
    ReDim Preserve v(2, UBound(v, 2) - 1)

    '実行前の確認
    Dim comfirmMsg As String
    comfirmMsg = "翻訳を実行しますがよろしいですか？（" + CStr(totalTxtLength) + "バイト / " + CStr(UBound(txt) + 1) + "回のAPI呼び出し）"
    If Not MsgBox(comfirmMsg, vbYesNo + vbQuestion, "実行確認") = vbYes Then
        wb.Close (False)
        Exit Sub
    End If

    'ここからCognitive Serviceで翻訳するよ！
    Dim httpClient  As Object
    Dim url As String
    Dim sourceText  As String
    Dim resultText  As String
    Dim translatedText  As String
    translatedText = ""

    '翻訳対象の文字列は結合して配列にいれてあるのでグルグルする
    For i = 0 To UBound(txt)
        
        '翻訳結果は最後にXMLの改行コードでSplitするので"&#xD;"でつなげておくよ
        If Not translatedText = "" Then
            translatedText = translatedText + "&#xD;"
        End If
        
        sourceText = txt(i)
        
        'API呼び出し!!
        Set httpClient = CreateObject("MSXML2.XMLHTTP")
        url = "https://api.microsofttranslator.com/V2/Http.svc/Translate?" _
                + "to=" + toLang _
                + "&from=" + fromLang _
                + "&text=" + Application.WorksheetFunction.EncodeURL(sourceText)
                
        With httpClient
            .Open "GET", url, False
            .SetRequestHeader "Ocp-Apim-Subscription-Key", apiKey
            .Send
            resultText = .ResponseText
        End With
        
        '翻訳結果からタグと改行を削除
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.ignoreCase = True
        re.Pattern = "<[^>]+?>"
        resultText = re.Replace(resultText, "")
        resultText = Replace(resultText, vbLf, "")
        
        translatedText = translatedText + resultText
        
    Next i
    
    
    '翻訳結果の文字列をXMLの改行コードでスプリット
    Dim splitted As Variant
    splitted = Split(translatedText, "&#xD;")
    
    '配列に翻訳結果を書き込む
    For i = 0 To UBound(v, 2)
        If i <= UBound(splitted) Then
            v(2, i) = splitted(i)
        Else
            v(2, i) = ""
        End If
    Next i
    
    '実行元Book、書き出しシートをアクティブにする→セルクリア
    hostBook.Activate
    hostBook.Sheets(2).Select
    Cells.Clear
    
    '自動更新/計算をOFFにしておくよ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    '2次元配列の縦横を入れ替えて書き込み！
    Range(Cells(1, 1), Cells(UBound(v, 2) + 1, UBound(v, 1) + 1)) = WorksheetFunction.Transpose(v)
    
    '自動更新/計算をONにもどすよ
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    '開いたBookを閉じる
    wb.Close (False)
    
    '結果シートを表示。お疲れ様でした！
    hostBook.Sheets(2).Cells(1, 1).Select

End Sub
