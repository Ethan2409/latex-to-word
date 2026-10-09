' LaTeX 公式转 Word 格式（剪贴板版）
' 用法：先复制 LaTeX 文本到剪贴板，再运行本宏。
' 宏会以纯文本粘贴剪贴板内容，并把其中的 $...$ 内联公式转成 Word 论文格式
'（变量斜体、数学函数正体、上下标、希腊字母 Unicode 化）。

Sub PasteAsTextAndConvertLatex()
    Dim startPos As Long
    Dim endPos As Long
    Dim pasteRng As Range
    
    ' 1. 记录当前光标起始位置
    startPos = Selection.Start
    
    ' 2. 以纯文本形式粘贴剪贴板内容
    On Error Resume Next
    Selection.PasteSpecial DataType:=wdPasteText
    If Err.Number <> 0 Then
        MsgBox "剪贴板为空或不包含可粘贴的文本数据。", vbExclamation, "错误"
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. 记录粘贴后的光标结束位置
    endPos = Selection.End
    
    ' 4. 框定并选中刚刚粘贴的文本范围
    Set pasteRng = ActiveDocument.Range(startPos, endPos)
    pasteRng.Select
    
    ' 拦截异常：如果粘贴后没有实际内容，则退出
    If Selection.Type = wdSelectionIP Then Exit Sub
    
    ' ------------------ 以下为 Latex 转换核心逻辑 ------------------
    Dim rng As Range
    Dim txt As String
    Dim i As Long
    Dim regEx As Object
    
    ' 为当前选区创建一个动态隐形书签，Word 会自动跟踪其长度变化
    Dim bkm As Bookmark
    Set bkm = ActiveDocument.Bookmarks.Add(Name:="TempLatexBounds", Range:=Selection.Range)
    
    ' 初始化正则表达式引擎
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    Application.ScreenUpdating = False
    
    ' 将搜索范围严格绑定在书签内
    Set rng = bkm.Range
    
    With rng.Find
        .ClearFormatting
        .Text = "\$[!\$]{1,}\$"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            ' 绝对防御：如果检索到的结果已经跑出了我们的书签范围，立刻强制退出循环
            If Not rng.InRange(bkm.Range) Then Exit Do
            
            ' 判定是否属于双美金符号 $$...$$ 独立公式块
            Dim isDoubleDollar As Boolean
            isDoubleDollar = False
            
            If rng.Start > ActiveDocument.Content.Start Then
                If ActiveDocument.Range(rng.Start - 1, rng.Start).Text = "$" Then
                    isDoubleDollar = True
                End If
            End If
            
            If isDoubleDollar Then
                ' 如果是双美金公式，向后多步进一个字符以跨过结尾的第二个 $
                If rng.End < ActiveDocument.Content.End Then
                    If ActiveDocument.Range(rng.End, rng.End + 1).Text = "$" Then
                        rng.End = rng.End + 1
                    End If
                End If
                ' 直接折叠光标至末尾，跳过本次循环的所有格式转化
                rng.Collapse wdCollapseEnd
            Else
                ' ------------------ 原有单美金内联公式转化逻辑 ------------------
                ' 剥离首尾的 $ 符号
                txt = Mid(rng.Text, 2, Len(rng.Text) - 2)
                
                ' 清洗 \text{} 和 \mathrm{}
                regEx.Pattern = "\\text\{([^}]+)\}"
                txt = regEx.Replace(txt, "$1")
                regEx.Pattern = "\\mathrm\{([^}]+)\}"
                txt = regEx.Replace(txt, "$1")
                
                ' 符号转换与清洗
                txt = Replace(txt, "-", ChrW(&H2212)) ' 转换为标准减号
                txt = Replace(txt, "\times", ChrW(&H00D7)) ' 转换为标准乘号
                txt = Replace(txt, "\cdot", ChrW(&H22C5))   ' 转换为点乘号
                txt = Replace(txt, "\Phi", ChrW(&H3A6))
                txt = Replace(txt, "\phi", ChrW(&H3C6))
                txt = Replace(txt, "\alpha", ChrW(&H3B1))
                txt = Replace(txt, "\beta", ChrW(&H3B2))
                txt = Replace(txt, "\gamma", ChrW(&H3B3))
                txt = Replace(txt, "\theta", ChrW(&H3B8))
                txt = Replace(txt, "\Delta", ChrW(&H394))
                txt = Replace(txt, "\delta", ChrW(&H3B4))
                txt = Replace(txt, "\mu", ChrW(&H3BC))
                txt = Replace(txt, "\rho", ChrW(&H3C1))
                txt = Replace(txt, "\sigma", ChrW(&H3C3))
                txt = Replace(txt, "\eta", ChrW(&H3B7))
                txt = Replace(txt, "\varepsilon", ChrW(&H3B5))
                txt = Replace(txt, "\", "") 
                
                ' 回写文本并强行设置新罗马字体
                rng.Text = txt
                rng.Font.Name = "Times New Roman"
                
                ' 第一遍扫描：设置斜体与正体
                Dim chRng As Range
                Dim charCode As Long
                For i = 1 To rng.Characters.Count
                    Set chRng = rng.Characters(i)
                    charCode = AscW(chRng.Text)
                    If chRng.Text Like "[A-Za-z]" Or (charCode >= &H370 And charCode <= &H3FF) Then
                        chRng.Font.Italic = True
                    Else
                        chRng.Font.Italic = False
                    End If
                Next i
                
                ' 修正标准数学函数
                Dim mathFuncs As Variant
                Dim funcStr As Variant
                Dim matchPos As Long
                
                mathFuncs = Array("log", "ln", "sin", "cos", "tan", "exp", "lim", "max", "min")
                For Each funcStr In mathFuncs
                    matchPos = InStr(1, rng.Text, funcStr, vbTextCompare)
                    Do While matchPos > 0
                        Dim funcRng As Range
                        Set funcRng = ActiveDocument.Range(Start:=rng.Characters(matchPos).Start, End:=rng.Characters(matchPos + Len(funcStr) - 1).End)
                        funcRng.Font.Italic = False
                        matchPos = InStr(matchPos + Len(funcStr), rng.Text, funcStr, vbTextCompare)
                    Loop
                Next funcStr
                
                ' 第二遍扫描（倒序）：处理下标
                For i = rng.Characters.Count To 1 Step -1
                    If i <= rng.Characters.Count Then
                        If rng.Characters(i).Text = "_" Then
                            If i < rng.Characters.Count Then
                                If rng.Characters(i + 1).Text = "{" Then
                                    Dim j As Long, closeBrace As Long
                                    closeBrace = 0
                                    For j = i + 2 To rng.Characters.Count
                                        If rng.Characters(j).Text = "}" Then
                                            closeBrace = j
                                            Exit For
                                        End If
                                    Next j
                                    
                                    If closeBrace > 0 Then
                                        Dim subRng As Range
                                        Set subRng = ActiveDocument.Range(rng.Characters(i + 2).Start, rng.Characters(closeBrace - 1).End)
                                        subRng.Font.Subscript = True
                                        subRng.Font.Italic = False
                                        
                                        rng.Characters(closeBrace).Delete
                                        rng.Characters(i + 1).Delete
                                        rng.Characters(i).Delete
                                    End If
                                Else
                                    rng.Characters(i + 1).Font.Subscript = True
                                    rng.Characters(i + 1).Font.Italic = False
                                    rng.Characters(i).Delete
                                End If
                            End If
                        End If
                    End If
                Next i
                
                ' 第三遍扫描（倒序）：处理上标
                For i = rng.Characters.Count To 1 Step -1
                    If i <= rng.Characters.Count Then
                        If rng.Characters(i).Text = "^" Then
                            If i < rng.Characters.Count Then
                                If rng.Characters(i + 1).Text = "{" Then
                                    Dim k As Long, closeBraceSup As Long
                                    closeBraceSup = 0
                                    For k = i + 2 To rng.Characters.Count
                                        If rng.Characters(k).Text = "}" Then
                                            closeBraceSup = k
                                            Exit For
                                        End If
                                    Next k
                                    
                                    If closeBraceSup > 0 Then
                                        Dim supRng As Range
                                        Set supRng = ActiveDocument.Range(rng.Characters(i + 2).Start, rng.Characters(closeBraceSup - 1).End)
                                        supRng.Font.Superscript = True
                                        supRng.Font.Italic = False
                                        
                                        rng.Characters(closeBraceSup).Delete
                                        rng.Characters(i + 1).Delete
                                        rng.Characters(i).Delete
                                    End If
                                Else
                                    rng.Characters(i + 1).Font.Superscript = True
                                    rng.Characters(i + 1).Font.Italic = False
                                    rng.Characters(i).Delete
                                End If
                            End If
                        End If
                    End If
                Next i
                
                ' 折叠光标，继续向后检索
                rng.Collapse wdCollapseEnd
            End If
        Loop
    End With
    
    ' 销毁临时书签
    bkm.Delete
    Application.ScreenUpdating = True
End Sub
