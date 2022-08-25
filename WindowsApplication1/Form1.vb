Imports System.Text.RegularExpressions

Public Class Form_calc


    'デバッグ用
    Function debugLogLine(any)
        System.Diagnostics.Debug.WriteLine(any)
    End Function
    Function debugLog(any)
        System.Diagnostics.Debug.Write(any)
    End Function
    'デバッグ用




    'リセット
    Function reset()
        currentVal = Nothing
        enterNum = Nothing
        answer = Nothing
        TextBox_form.Text = Nothing
        Label_answer.Text = 0
    End Function

    Function formDisplay(displayStr As String)
        currentVal = displayStr
        enterNum += currentVal
        TextBox_form.Text += currentVal
    End Function

    '入力した数値をTextBoxに出力
    Function setTextNum(numSender As Button)
        If enterNum Is Nothing Then
            formDisplay(numSender.Text)
        ElseIf Not bhdSym.IsMatch(enterNum) Then
            If enterNum = "0" Then
                currentVal = numSender.Text
                enterNum = currentVal
                TextBox_form.Text = TextBox_form.Text.Remove(TextBox_form.Text.Length - 1, 1)
                TextBox_form.Text += currentVal
            Else
                formDisplay(numSender.Text)
            End If
        End If
    End Function

    'キーボード入力と記号入力処理
    Function setTextNumKey(numKey)
        If enterNum Is Nothing Then
            formDisplay(numKey)
        ElseIf numKey = ")" Then
            formDisplay(numKey)
        ElseIf Not bhdSym.IsMatch(enterNum) Then
            If enterNum = "0" Then
                currentVal = numKey
                enterNum = currentVal
                TextBox_form.Text = TextBox_form.Text.Remove(TextBox_form.Text.Length - 1, 1)
                TextBox_form.Text += currentVal
            Else
                formDisplay(numKey)
            End If
        ElseIf numKey = "²" And currentVal = ")" Then
            formDisplay(numKey)
        End If
    End Function

    '入力した演算子をTextBoxに出力
    Function setTextOp(opSender As Button)
        If currentVal IsNot Nothing Then
            If Not fourArithOp.IsMatch(currentVal) Then
                currentVal = opSender.Text
                TextBox_form.Text += currentVal
                enterNum = Nothing
            End If
        End If
    End Function

    'キーボードから入力した演算子をTextBoxに出力
    Function setTextOpKey(opKey)
        If currentVal IsNot Nothing Then
            If Not fourArithOp.IsMatch(currentVal) Then
                currentVal = opKey
                TextBox_form.Text += currentVal
                enterNum = Nothing
            End If
        End If
    End Function

    '数式記号の入力処理
    Function setTextSym(sym As Button, Optional str As String = Nothing, Optional head As Boolean = False)
        If sym.Text = "√" Then
            If currentVal IsNot "√" Then
                setTextNumKey(sym.Text)
            End If
        ElseIf sym.Text = "(" Then
            setTextNumKey(sym.Text)
            If currentVal = "(" Then
                bktCnt += 1
            End If
        ElseIf enterNum IsNot Nothing Then
            If str Is Nothing Then
                If sym.Text = ")" Then
                    If currentVal IsNot "(" And bktCnt > 0 Then
                        setTextNumKey(sym.Text)
                        If currentVal = ")" Then
                            bktCnt -= 1
                        End If
                    End If
                ElseIf Not enterNum.Contains(sym.Text) Then
                    setTextNumKey(sym.Text)
                End If
            ElseIf Not enterNum.Contains(str) Then
                setTextNumKey(str)
            End If
        End If
    End Function


    'textBoxに入力された中置記法の計算式を後置記法に変換する関数
    Function convertRPN(resultFormula As String) As List(Of Object)
        Dim numList = New List(Of Object)()
        Dim opStack = New Stack(Of Object)()
        Dim bktList = New List(Of String)()
        Dim bkeStack = New Stack(Of Object)()

        calcList = (Regex.Split(resultFormula, "([\+\-×÷()²])"))

        Dim rx0 = New Regex("[\+\-×÷]", RegexOptions.Compiled)
        Dim rx1 = New Regex("[²]", RegexOptions.Compiled)
        Dim rx2 = New Regex("[×÷]", RegexOptions.Compiled)

        Dim bkt As Integer = 0

        For Each obj In calcList
            '数値の処理
            If obj = "(" Then
                bkt += 1
            End If
            If bkt > 0 Then
                bktList.Add(obj)
                If obj = ")" Then
                    bkt -= 1
                End If
                If bkt = 0 Then
                    bktList.RemoveRange(0, bktList.IndexOf("(") + 1)
                    bktList.RemoveAt(bktList.LastIndexOf(")"))
                    Dim joinBktString = String.Join("", bktList.ToArray())
                    Dim bktRpnList As List(Of Object) = convertRPN(joinBktString)
                    For Each bktItem In bktRpnList
                        numList.Add(bktItem)
                    Next
                End If
            ElseIf IsNumeric(obj) Or obj.contains("%") Or obj = "²" Then
                numList.Add(obj)
                '演算記号の処理開始
            ElseIf rx0.IsMatch(obj) Then
                'stackが空ならそのままpush
                If opStack.Count = 0 Then
                    'ElseIf rx1.IsMatch(opStack.Peek())
                    'stackの先頭が×か÷のときの演算子処理
                ElseIf rx2.IsMatch(opStack.Peek()) And rx2.IsMatch(obj) Then
                    'stackの先頭が×か÷のときのみその演算子をadd(+と-は残る)
                    numList.Add(opStack.Pop())
                ElseIf rx2.IsMatch(obj) Then
                Else
                    Dim cnt As Integer = opStack.Count
                    For i As Integer = cnt To 1 Step -1
                        '演算子stack内にある演算子を全てにadd
                        numList.Add(opStack.Pop())
                    Next
                End If
                opStack.Push(obj)
            End If
        Next
        Dim remCnt As Integer = opStack.Count
        For i As Integer = remCnt To 1 Step -1
            'stack内の残りの演算子をadd
            numList.Add(opStack.Pop())
        Next

        For Each i In numList
            debugLog(i)
        Next
        debugLogLine("")

        Return numList
    End Function


    '引数に渡された後置記法の計算式を計算し、結果を返す関数
    Function calcRPN(formulaRpn) As Double
        Dim calcVal As Double
        Dim calcStack = New Stack(Of Object)

        For Each f In formulaRpn
            debugLogLine("c")
            'stackから取り出した演算記号にしたがって計算する
            If allOp.IsMatch(f) Then
                debugLogLine("a")
                Select Case f
                    Case "+"
                        calcVal = calcStack.Pop() + calcStack.Pop()
                        calcStack.Push(calcVal)
                    Case "-"
                        calcVal = calcStack.Pop() - calcStack.Pop()
                        calcStack.Push(calcVal)
                    Case "×"
                        calcVal = calcStack.Pop() * calcStack.Pop()
                        calcStack.Push(calcVal)
                    Case "÷"
                        Dim divNum As Double = calcStack.Pop()
                        calcVal = calcStack.Pop() / divNum
                        calcStack.Push(calcVal)
                    Case "²"
                        calcVal = calcStack.Pop() ^ 2
                        calcStack.Push(calcVal)
                End Select

            Else
                debugLogLine("b")
                '百分率表記の場合の修正
                If f.IndexOf("%") > 0 Then
                    f = Double.Parse(f.replace("%", "")) * 0.01
                End If
                f = Double.Parse(f)
                calcStack.Push(f)
            End If
        Next

        For Each i In calcStack
            debugLogLine(i)
        Next

        Return calcStack.Peek()
    End Function




    'メンバ関数定義
    Dim currentVal As String
    Dim enterNum As String
    Dim answer As Double = 0
    Dim bktCnt As Integer = 0
    Dim calcList = New List(Of Object)()
    Dim fourArithOp = New Regex("[\+\-×÷]", RegexOptions.Compiled)
    Dim allOp = New Regex("[\+\-×÷²]", RegexOptions.Compiled)
    Dim bhdSym = New Regex("[²%)]\z", RegexOptions.Compiled)
    Dim headSym = New Regex("[√]", RegexOptions.Compiled)






    Private Sub frmFocus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActiveControl = Button_equal
    End Sub

    Private Sub Button_0_Click(sender As System.Object, e As System.EventArgs) Handles Button_0.Click
        '整数で0が重複しない
        If Not enterNum = "0" Then
            setTextNum(sender)
        End If

    End Sub

    Private Sub Button_1_Click(sender As System.Object, e As System.EventArgs) Handles Button_1.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_2_Click(sender As System.Object, e As System.EventArgs) Handles Button_2.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_3_Click(sender As System.Object, e As System.EventArgs) Handles Button_3.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_4_Click(sender As System.Object, e As System.EventArgs) Handles Button_4.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_5_Click(sender As System.Object, e As System.EventArgs) Handles Button_5.Click
        setTextNum(sender)
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_7_Click(sender As System.Object, e As System.EventArgs) Handles Button_7.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_8_Click(sender As System.Object, e As System.EventArgs) Handles Button_8.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_9_Click(sender As System.Object, e As System.EventArgs) Handles Button_9.Click
        setTextNum(sender)
    End Sub

    Private Sub Button_sum_Click(sender As System.Object, e As System.EventArgs) Handles Button_sum.Click
        setTextOp(sender)
    End Sub

    Private Sub Button_sub_Click(sender As System.Object, e As System.EventArgs) Handles Button_sub.Click
        setTextOp(sender)
    End Sub

    Private Sub Button_mul_Click(sender As System.Object, e As System.EventArgs) Handles Button_mul.Click
        setTextOp(sender)
    End Sub

    Private Sub Button_div_Click(sender As System.Object, e As System.EventArgs) Handles Button_div.Click
        setTextOp(sender)
    End Sub

    Private Sub Button_leftBkt_Click(sender As System.Object, e As System.EventArgs) Handles Button_leftBkt.Click
        setTextSym(sender)
    End Sub

    Private Sub Button_rightBkt_Click(sender As System.Object, e As System.EventArgs) Handles Button_rightBkt.Click
        setTextSym(sender)
    End Sub

    Private Sub Button_dot_Click(sender As System.Object, e As System.EventArgs) Handles Button_dot.Click
        setTextSym(sender)
    End Sub

    Private Sub Button_pct_Click(sender As System.Object, e As System.EventArgs) Handles Button_pct.Click
        setTextSym(sender)
    End Sub

    Private Sub Button_sq_Click(sender As System.Object, e As System.EventArgs) Handles Button_sq.Click
        setTextSym(sender, "²")
    End Sub

    Private Sub Button_sqrt_Click(sender As System.Object, e As System.EventArgs) Handles Button_sqrt.Click
        setTextSym(sender)
    End Sub

    Private Sub Button_equal_Click(sender As System.Object, e As System.EventArgs) Handles Button_equal.Click
        'Try
        If currentVal IsNot Nothing Then
            If Not fourArithOp.IsMatch(currentVal) Then
                Dim rpm = New List(Of Object)()
                rpm = convertRPN(TextBox_form.Text)
                answer = calcRPN(rpm)
                Label_answer.Text = answer
            End If
        End If
        'Catch ex As Exception
        'Label_answer.Text = "error"
        'End Try

    End Sub

    Private Sub Button_ans_Click(sender As System.Object, e As System.EventArgs) Handles Button_ans.Click
        Dim bfAns As String = answer
        reset()
        currentVal = bfAns
        enterNum += bfAns
        TextBox_form.Text = currentVal
    End Sub

    Private Sub Button_clear_Click(sender As System.Object, e As System.EventArgs) Handles Button_clear.Click
        reset()
    End Sub

    Private Sub Button_bs_Click(sender As System.Object, e As System.EventArgs) Handles Button_bs.Click
        If TextBox_form.Text.Length > 0 Then
            TextBox_form.Text = TextBox_form.Text.Remove(TextBox_form.Text.Length - 1, 1)
            currentVal = Strings.Right(TextBox_form.Text, 1)
            If Not allOp.IsMatch(currentVal) Then
                If allOp.IsMatch(TextBox_form.Text) Then
                    enterNum = TextBox_form.Text.Substring(
                        TextBox_form.Text.LastIndexOfAny(New Char() {"+", "-", "×", "÷"}) + 1,
                        TextBox_form.Text.Length - TextBox_form.Text.LastIndexOfAny(New Char() {"+", "-", "×", "÷"}) - 1
                    )
                Else
                    enterNum = TextBox_form.Text
                End If
            End If
        End If

    End Sub

    Private Sub Form_calc_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Select Case e.KeyChar
            Case "0" To "9"
                setTextNumKey(e.KeyChar)
            Case "+"
                setTextOpKey("+")
            Case "-"
                setTextOpKey("-")
            Case "*"
                setTextOpKey("×")
            Case "/"
                setTextOpKey("÷")
            Case "."
                If enterNum IsNot Nothing Then
                    If Not enterNum.Contains(".") Then
                        setTextNumKey(e.KeyChar)
                    End If
                End If
        End Select

    End Sub

End Class
