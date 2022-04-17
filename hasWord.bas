Option Explicit

Sub regexptest()
    Dim str As String
    str = "イケメン"
    
    If hasWord(str) = True Then
        Debug.Print str & "→Youはかっこいいね"
    Else
        Debug.Print str & "→そうじゃないみたいだね"
    End If

End Sub


Function hasWord(ByVal s As String) As Boolean
' =======================================================================
' 関数名  :hasWord
' 関数概要:引数として渡した文字列がPatternに合致するかどうか
' 引数    :s Patternに合致するか調べたい文字列
' 返り値  :合致する→True がっちしない→False
' =======================================================================

    'RegExpを使えるようにオブジェクト宣言
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    Dim strPat As String: strPat = "かっこいい|素敵|イケメン|"
    
    With reg
        .Pattern = strPat
        .Global = True '文字列全体を見る
        
        If .Test(s) Then 'Patternに合致したら
            hasWord = True
        Else
            hasWord = False
        End If
    End With

End Function
