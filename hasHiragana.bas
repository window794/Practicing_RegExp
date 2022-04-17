Option Explicit
Sub RegExpTest()

	Dim chkStr As String
	chkStr = "加茂の流れ"
	
	If hasHiragana(chkStr) = True Then
		MsgBOx "ひらがなを含みます"
	Else
		MsgBox "ひらがなはありませんでした"
	End If

End Sub

Function hasHiragana(ByVal s As String) As Boolean
' =======================================================================
' 関数名  :hasHiragana
' 関数概要:引数として渡した文字列にひらがなが含まれるかを判定
' 引数    :s ひらがなを含んでいるか調べたい文字列
' 返り値  :ひらがなが含まれる→True ひらがなが含まれない→False
' =======================================================================

    'RegExpを使えるようにオブジェクト宣言
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = "[\u3040-\u309F]" '正規表現でひらがな判定する
        .Global = True '文字列全体を見る
        
        If .Test(s) Then 'Patternに合致したら（ひらがなだったら）
            hasHiragana = True
        Else
            hasHiragana = False
        End If
    End With

End Function
