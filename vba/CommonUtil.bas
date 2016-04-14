Attribute VB_Name = "CommonUtil"

'**************************************************
'**
'** 返回 EG, SH, SP, SF, B, M, A之一
'**
'**************************************************
Public Function getLayer(trans As String) As String

    'e.g. azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    Dim tmp As String
    ' Dim ssys As String
    tmp = SWITCH(Left(trans, 1) = "I", VBA.Mid(trans, 14, 3), Left(trans, 1) = "T", VBA.Mid(trans, 9, 3))
    
    'Debug.Print trans & " " & tmp
    
    getLayer = SWITCH(Left(tmp, 1) = "S" Or Left(tmp, 1) = "E", Left(tmp, 1) & Right(tmp, 1), Left(tmp, 1) <> "S" And Left(tmp, 1) <> "E", Left(tmp, 1))

End Function


' 未启用
Public Function getLayer_new(trans As String) As String

    'e.g. azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    Dim tmp As String
    Dim o As String
    
    Dim res As String
    
    ' Dim ssys As String
    tmp = SWITCH(Left(trans, 1) = "I", Mid(trans, 14, 4), Left(trans, 1) = "T", Mid(trans, 9, 4))
    
    o = Left(tmp, 1)
    
    Select Case o
    
        Case "S"
        
            res = Left(tmp, 1) & Mid(tmp, 3, 1) 'SH SP SF
        
        Case "B", "M", "A", "E"
        
            res = o
    
        Case "E"
        
            res = Left(tmp, 1) & Right(tmp, 2) 'EGA EGC
        
        Else
            res = "-"
        
    End Select
    
    'Debug.Print trans & " " & tmp
    getLayer_new = res

End Function

'**************************************************
'*
'*
'*
'*
'**************************************************
Public Function makeCMD(REP As String, _
                           usr As String, _
                           pass As String, _
                           trans As String, _
                           cmd As String _
) As String

    'e.g. azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    
    Dim DIR As String
    
    Dim layer As String
    
    layer = Left(getLayer(trans), 1)
    
    DIR = SWITCH(layer = "S", "/hcdw/ods/StdLayer/ktr/daily", _
                 layer = "B", "/hcdw/edw/BasLayer/ktr/daily", _
                 layer = "M", "/hcdw/edw/MidLayer/ktr/daily", _
                 layer = "A", "/hcdw/edw/AppLayer/ktr/daily")
                 
    makeCMD = cmd _
              & " -rep=" & REP _
              & " -trans=" & trans _
              & " -dir=" & DIR _
              & " -user=" & usr _
              & " -pass=" & pass
              
End Function

Public Function isTransName(name As String) As Boolean

    isTransName = SWITCH(Left(name, 4) = "INIT" Or Left(name, 4) = "TRAN", True, Left(name, 4) <> "INIT" And Left(name, 4) <> "TRAN", False)

End Function

Public Function jobType(name As String) As String

    Dim left4 As String
    
    left4 = Left(name, 4)
    
    
    Select Case left4
        Case "INIT", "TRAN"
            jobType = "TRAN"
        Case "FLOW"
            jobType = "FLOW"
            
        Case Else
            jobType = "ECHO"
            
    End Select
    
End Function

'**************************************************
'*
'* azkpan2.sh
'*
'*
'**************************************************
Public Function makeCMD2(REP As String, _
                           usr As String, _
                           pass As String, _
                           trans As String, _
                           cmd As String _
) As String

    'e.g. azkpan2.sh TRANS_4_S_PL_CRM_INTOPIECES_DK_H
    
    Dim DIR As String
    
    Dim layer As String
    
    'Layer = LEFT(getLayer(trans), 1)
    
    'DIR = Switch(Layer = "S", "/hcdw/ods/StdLayer/ktr/daily", _
    '             Layer = "B", "/hcdw/edw/BasLayer/ktr/daily", _
    '             Layer = "M", "/hcdw/edw/MidLayer/ktr/daily", _
    '             Layer = "A", "/hcdw/edw/AppLayer/ktr/daily")
    
    Dim transFlg As Boolean
    Dim jtype As String
    
    '***************************************************************
    '*
    '* 把需要用looppan.sh调度的放在这里 生产环境时需要注释掉
    '*
    '***************************************************************
    Dim LOOPPAN_SH_LIST As Object
    
    Set LOOPPAN_SH_LIST = CreateObject("Scripting.Dictionary")
    
    LOOPPAN_SH_LIST.Add "TRANS_5_A_HYR_LEND_MATCH_STAT", BLANK
    '
    '////////////////////////////////////////////////////////////////////

    transFlg = isTransName(trans)
    jtype = jobType(trans)
    
    If LOOPPAN_SH_LIST.Exists(trans) Then
        cmd = LOOP_PAN
    End If
                
    ' makeCMD2 = SWITCH(transFlg, cmd & " " & trans, Not transFlg, " echo ""JOB " & trans & " IS DONE. """)
    
    ' 2015/12/14
    Select Case jtype
    
        Case "TRAN"
        
            makeCMD2 = cmd & " " & trans
            
        Case "ECHO"
        
            makeCMD2 = " echo ""JOB " & trans & " IS DONE. """
            
        Case "FLOW"
            
            makeCMD2 = Mid(trans, 6)
            
        Case Else
            
            makeCMD2 = " echo ""JOB " & trans & " IS DONE. """
            
    End Select
              
End Function


'**************************************************
'**
'** 生成JOB文件对应的内容（废弃2015/10/16）
'**
'**************************************************
Public Function makeJOB_old(cmt As String, _
                        typ As String, _
                        cmd As String, _
                        Optional dep As String = "" _
                          ) As String

    ' e.g.
    ' # TRANS_4_S_PL_CRM_INTOPIECES_DK_H
    ' type=command
    ' command=azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    
        makeJOB = "# " _
                  & cmt _
                  & Chr(13) _
                  & "type=" _
                  & typ _
                  & Chr(13) _
                  & "command=" _
                  & cmd _
                  & SWITCH(dep <> "", Chr(13) & "dependencies=") _
                  & SWITCH(dep <> "", dep)
              
End Function

'**************************************************
'**
'**
'** 生成JOB文件对应的内容
'**
'** 2015/10/16 修改 ： 新增参数c：被依赖的任务列表
'**
'**************************************************
Public Function makeJOB(trans As String, _
                        typ As String, _
                        cmd As String, _
                        Optional dep As String = "", _
                        Optional ByRef c As Collection = Nothing, _
                        Optional cmdname As String = "command" _
                          ) As String

    ' e.g.
    ' # TRANS_4_S_PL_CRM_INTOPIECES_DK_H
    ' type=command
    ' command=azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    
        Dim o
        Dim deplist As String
        Dim cmt As String
        
        deplist = ""
        cmt = trans

        If Not c Is Nothing Then
            
            For Each o In c
                
                ' 过滤掉""
                If o <> "" Then
                
                        If Left(o, 1) <> "#" Then
                            
                            deplist = deplist & "," & o ' e.g. [ ,job1,job2,job3 ]
                        
                        Else
                        
                            ' 取最后一次#开头的内容作为注释
                            cmt = o
                        
                        End If
                End If
            Next
        End If
        
        'Debug.Print dep & deplist
                
        ' 去掉多余的逗号
        'deplist = Switch(LEFT(dep & deplist, 1) = ",", Mid(dep & deplist, 2), LEFT(dep & deplist, 1) <> ",", dep & deplist)
        deplist = trimComma(dep & deplist)
        
        ' Debug.Print deplist
        
        ' embbed flow job 2015/12/14
        If jobType(trans) = "FLOW" Then
            typ = "flow"
            cmdname = "flow.name"
        End If
        
        makeJOB = "# " _
                  & cmt _
                  & Chr(13) _
                  & "type=" _
                  & typ _
                  & Chr(13) _
                  & cmdname & "=" _
                  & cmd _
                  & SWITCH(deplist <> "", Chr(13) & "dependencies=") _
                  & SWITCH(deplist <> "", deplist)
              
End Function

Public Sub CreateIfNotExists(strDirName As String)

    If DIR(strDirName, vbDirectory) = Empty Then ' 16
            ''''''' to do
        
            MKDIR strDirName
            
        End If

End Sub

'******************
'*
'* package
'*
'******************
Public Sub package(ByVal sdir As String, rardir As String, suffix As String)

    
    Dim cmd, str, i
    
    If Right(sdir, 1) = "\" Then
        sdir = Left(sdir, Len(sdir) - 1)
    End If
    
    str = Split(sdir, "\")
    
    i = UBound(str)
    
    ' cmd /c "C:\Program Files (x86)\WinRAR\RAR.exe" a D:\tmp\20150909_azk2proc\20150909_azk2proc.zip D:\tmp\20150909_azk2proc\*.job
    
    ' 临时文件bat
    'Shell "cmd /c echo @cd /d " & sdir & " >  " & sdir & "\rartest.bat"
    
    'D:\tmp\20151218_azk2procv2>"C:\Program Files (x86)\WinRAR\WinRAR.exe" a -afzip zipfile *.job
    
    'Shell "cmd /c echo @cd /d """ & sdir & """ >  " & sdir & "\rartest.bat"
    'Shell "cmd /c echo """ & rardir & "\winrar.exe"" a  -afzip " & str(i) & " *.job  >> " & sdir & "\rartest.bat"
    
    'Debug.Print "cmd /c echo @cd /d """ & sdir & """ >  " & sdir & "\rartest.bat"
        
    'Debug.Print "cmd /c echo """ & rardir & "\winrar.exe"" a  -afzip " & str(i) & " *.job  >> " & sdir & "\rartest.bat"
    
    Dim objTXTFO As New FileOperation
    
    objTXTFO.OpenFile sdir & "\rartest.bat", "W"
    
    objTXTFO.WriteLine ("title PACKAGE......")
    
    ' cd /d "**"
    objTXTFO.WriteLine ("cd /d """ & sdir & """")
    
    '"C:\Program Files (x86)\WinRAR\WinRAR.exe" a -afzip zipfile *.job
    objTXTFO.WriteLine ("""" & rardir & "\WinRAR.exe"" a -afzip " & str(i) & " *.job")
    
    'objTXTFO.WriteLine ("PAUSE")
    objTXTFO.CloseFile
    
    Set objTXTFO = Nothing
    
    ' 运行bat文件
    cmd = sdir & "\rartest.bat"
    
    ' cmd = "cmd /c """ & rardir & "\RAR.exe"" a " & sdir & "\" & str(i) & "." & suffix & " " & sdir & "\*.job"
    
    'Debug.Print cmd
    
    ' Shell cmd
    
    Dim WSH As Object, wExec As Object, result
   
    Set WSH = CreateObject("WScript.Shell")
       
    Set wExec = WSH.Exec(cmd)
    
    result = wExec.StdOut.ReadAll
    
    Logger.LogInfo CStr(result)
    
    Set WSH = Nothing
    
End Sub


Public Function getFullRepPath(trans As String) As String
    
    Dim layer As String
    
    layer = Left(getLayer(trans), 1)
    
    getFullRepPath SWITCH(layer = "S", "/hcdw/ods/StdLayer/ktr/daily/", _
                          layer = "B", "/hcdw/edw/BasLayer/ktr/daily/", _
                          layer = "M", "/hcdw/edw/MidLayer/ktr/daily/", _
                          layer = "A", "/hcdw/edw/AppLayer/ktr/daily/") _
                          & trans
End Function

'**************************************************
'*
'* 自定义函数：传入unix时间戳，返回时间(string型)
'*
'**************************************************
Public Function from_unixtime(ran As Range) As String

    Dim dtVal As Date
    
    dtVal = (CLng(ran.Value) + 8 * 3600) / 86400 + 70 * 365 + 19
    
    Debug.Print Format(dtVal, "yyyy-mm-dd")
    
    from_unixtime = Format(dtVal, "yyyy-mm-dd")
    
    ' 也可以返回date类型，需要前段设置下现实格式
    
End Function

'**************************************************
'*
'* trim掉逗号，默认是左comma
'*
'**************************************************
Public Function trimComma(s As String, Optional posflag As Byte = tcLEFT) As String

    Dim t As String

    
    Select Case posflag
    
        Case tcLEFT
        
            trimComma = LTrimComma(s)
        
        Case tcRIGHT
        
            trimComma = StrReverse(LTrimComma(StrReverse(s)))
        
        Case tcBOTH
        
            ' 先trim掉左边
            t = LTrimComma(s)
            
            ' 在反转，trim掉左边（此时是源s的右边）
             trimComma = StrReverse(LTrimComma(StrReverse(t)))
        
        Case Else
        
            trimComma = s
            
    End Select
    
End Function

''' 将左边的逗号给trim掉
Public Function LTrimComma(s As String) As String
    
    Dim i As Integer, slen As Integer
    
    Dim Char
    
    Dim skip As Boolean
    
    slen = Len(s)
    
    skip = False
    
    For i = 1 To slen
    
         Char = Mid(s, i, 1)
         
         ' 第一次遇到非comma，设定skip为true，即：跳过判断，从此不再修改skip的值
         If Char <> "," Then
            
            skip = True
            
            Exit For
            
         End If
         
         ' 从skip为true开始，累加字符
         'If skip Then
         '   val = val & char
         'End If
        
    Next
    
    LTrimComma = Right(s, Len(s) - i + 1)
    
End Function


Public Function quota(s As Variant, Optional mark As String = """") As String

    quota = SWITCH(CStr(s) = Empty, "", CStr(s) <> Empty, mark & CStr(s) & mark)

End Function

' 找上游
Private Function findDep(trans As String, ByRef ws As Worksheet) As Collection

    Dim cur_sheet As Worksheet
    
    Set cur_sheet = ws
    
    ' 计数变量
    Dim i As Integer
    
    Dim obj As Variant, dep As Variant, desc As Variant ' 临时变量，记录每行读取的trans和其依赖trans
    
    Dim c1 As New Collection
    
    i = 2
    
    
    Do While cur_sheet.Cells(i, 1) <> ""
        
            obj = cur_sheet.Cells(i, 1).Value  ' trans名字
            dep = cur_sheet.Cells(i, 2).Value  ' 依赖的trans名
        
            
            If obj = trans Then
                c1.Add (dep)
            End If
             
            ' 指向下一张表
            i = i + 1
                
    Loop
        
    ' MsgBox m0.Count
    Set findDep = c1
        
    
End Function

' 找下游
' 找不到时返回长度为0的集合
Public Function findDeped(trans As String, ByRef ws As Worksheet) As Collection

    Dim cur_sheet As Worksheet
    
    Set cur_sheet = ws
    
    ' 计数变量
    Dim i As Integer
    
    Dim obj As Variant, dep As Variant, desc As Variant ' 临时变量，记录每行读取的trans和其依赖trans
    
    Dim c1 As New Collection
    
    i = 2
    
    
    Do While cur_sheet.Cells(i, 1) <> ""
        
            obj = cur_sheet.Cells(i, 1).Value  ' trans名字
            dep = cur_sheet.Cells(i, 2).Value  ' 依赖的trans名
        
            ' 依赖关系和参数一致时
            If dep = trans Then
                c1.Add (obj)
            End If
             
            ' 指向下一张表
            i = i + 1
                
    Loop
        
    ' MsgBox m0.Count
    Set findDeped = c1
        
End Function

Private Function findNext(trans As String, ws As Worksheet, Optional ByRef D As Object = Nothing) As Boolean

    Dim dep As Collection
    
    Dim o As Variant
    Dim k
    Dim r As Boolean
    
    r = False
     
     If D Is Nothing Then
        Set D = CreateObject("Scripting.Dictionary")
     End If
    
    ' 将当期的trans放入检索集合
    D.Add trans, BLANK
    Debug.Print WorksheetFunction.Rept(">", 2) & trans
                    
    Set dep = findDep(trans, ws)
    
    If dep.Count > 0 Then
    
    
        For Each o In dep
            
            ' 如果已经依赖过一次，便可判定重复依赖
            If D.Exists(o) Then
            
                findNext = True
                
                Exit Function
                
            End If
            
            'For Each k In d.keys
            '    Debug.Print k
            'Next
            
             ' 找到依赖关系的话，再次往下找（递归）
             r = findNext(CStr(o), ws, D)
             
             ' 找到后直接退出循环，不再继续
             'If r Then
             '   Exit For
             'End If
             
        Next
        
    End If
        
    findNext = r

End Function

Private Function findPrevious(trans As String, ws As Worksheet, Optional ByRef D As Object = Nothing) As Boolean

    Dim dep As Collection
    
    Dim o As Variant
    Dim k
    Dim r As Boolean
    
    r = False
     
     If D Is Nothing Then
        Set D = CreateObject("Scripting.Dictionary")
     End If
    
    ' 将当期的trans放入检索集合
    ' d.Add trans, BLANK
    Debug.Print WorksheetFunction.Rept("> ", 2) & trans
                    
    Set dep = findDep(trans, ws)
    
    If dep.Count > 0 Then
    
    
        For Each o In dep
            
            '
            If CStr(o) = "START" Then
            
                findPrevious = True
        
            End If
            
            'For Each k In d.keys
            '    Debug.Print k
            'Next
                
             ' 找到依赖关系的话，再次往下找（递归）
             r = findPrevious(CStr(o), ws, D)
             
             
        Next
        
    End If
        
    findPrevious = r

End Function

' 是否循环依赖
Public Function isCircleDep(trans As String, ws As Worksheet) As Boolean

    isCircleDep = findNext(trans, ws)
    
End Function

' 是否循环依赖
Public Function checkHead(trans As String, ws As Worksheet) As Boolean

    checkHead = findPrevious(trans, ws)
    
End Function

' 建立完成的依赖关系时，把那些从未被依赖的任务找出来
'
' 遍历所有左侧的任务，查看其"在"右侧（有下游、被依赖）是否存在，把所有没有被依赖（没有下游）的任务返回
'
Public Function getNeverDependencyTaskList(ByRef ws As Worksheet) As Object

  Dim cur_sheet As Worksheet
    
    Set cur_sheet = ws
    
    ' 计数变量
    Dim i As Integer, j As Integer, k As Integer
    
    Dim obj As Variant
    
    Set hcdw = CreateObject("Scripting.Dictionary")
    
    i = 6 ' 从第六个开始找
    
    j = i - 2
    
    Do While cur_sheet.Cells(i, 1) <> "" And UCase(cur_sheet.Cells(i, 1)) <> "END"
        
            obj = cur_sheet.Cells(i, 1).Value  ' trans名字
         
            If Not isDependTask(CStr(obj), ws) Then
            
                If Not hcdw.Exists(obj) Then
                    
                        hcdw.Add obj, BLANK
                        
                End If
            
            End If
             
            ' 指向下一张表
            i = i + 1
            j = j + 1
                
        Loop
        
        Set getNeverDependencyTaskList = hcdw

End Function

' 查找当前元素在右侧是否存在 / 是否被依赖 / 是否在上游 / 是否在右侧 / 是否还有下游
Public Function isDependTask(task As String, ByRef ws As Worksheet) As Boolean

    Dim cur_sheet As Worksheet
    
    Set cur_sheet = ws
    
    ' 计数变量
    Dim i As Integer
    
    i = 6
    
    Do While cur_sheet.Cells(i, 1) <> "" And UCase(cur_sheet.Cells(i, 1)) <> "END"
                    
            If cur_sheet.Cells(i, 2) = task Then
                isDependTask = True
                Exit Function
            End If

             
            ' 指向下一张表
            i = i + 1
                
    Loop
    
    isDependTask = False
    
End Function

'*********************************************************************************************************
'
' 用找上游的方式读取指定的worksheet，利用表格数据描述的依赖关系，建立起邻接表，本函数自带检查是否有自循环功能
' 假设A依赖B，则表示为 [ A -> B ](!!!而不是[ B -> A ]!!!)
'_____________________________________________________________________________________________
'
' 【AdjacencyList Graph = 邻接表 存储依赖关系】
'
' Dict - key ----- value (Dict)                  <==> key -> value 表示key依赖（于）value
'        key1 --   Dict1(<d1,blank>,<d2,blank>)
'        key2 --   Dict2(<d3,blank>,<d4,blank>)
'        key3 --   Dict3(<d5,blank>           )
'_____________________________________________________________________________________________
'
' trans --- 最终依赖点，如果对整体依赖关系建立邻接表（假设依赖关系正确无误），值为『END』
' ws    --- 要读取的依赖关系所在表格
'
'*********************************************************************************************************
Public Function buildAdjaGraph(trans As String, ws As Worksheet, Optional ByRef D As Object = Nothing, Optional ByRef counter As Object = Nothing) As Object

    Dim dep As Collection, c As Object
    
    Dim o As Variant
    Dim k As Integer
    Dim circleFlg As Boolean
    
    r = False
     
     If D Is Nothing Then
        Set D = CreateObject("Scripting.Dictionary")
     End If
     
    ' counter 只用来判断是否产生自循环的情况
    If counter Is Nothing Then
        Set counter = CreateObject("Scripting.Dictionary")
        counter.Add "counter", 0
        counter.Add "b1", False
        counter.Add "b2", False
        counter.Add "circle", True ' no use
        counter.Add "conter2", 0 ' 连续不用『新增的』递归次数
        counter.Add "continueFlg", False ' 连续不新增flag
     End If
     
    counter.Item("b1") = False
    counter.Item("b2") = False
    
    ' 将当期的trans放入检索集合
    If Not D.Exists(trans) Then
      D.Add trans, CreateObject("Scripting.Dictionary")
      ' Debug.Print (" >> " & trans)
    Else
        counter.Item("b1") = True
    End If
    
    Set dep = findDep(trans, ws)
    
    If dep.Count > 0 Then
    
    
        For Each o In dep
        
            If CStr(o) = "" Then
            
                'Set buildAdjaGraph = d
                Exit For
        
            End If
            
            '
            Set c = D.Item(trans)
            
            'For Each k In d.keys
            '    Debug.Print k
            'Next
            If Not c.Exists(o) Then
                c.Add o, BLANK
            Else
               counter.Item("b2") = True
            End If
            ' Debug.Print (" >>> " & o)
            
            '连续『不新增』时，1++，否则初始化成0或者1
            If counter.Item("b1") And counter.Item("b2") And counter.Item("continueFlg") Then
                counter.Item("counter2") = counter.Item("counter2") + 1
            Else
                counter.Item("counter2") = (counter.Item("b1") And counter.Item("b2")) * -1
            End If
            
            ' 本次『不新增』发生时，将continueflg置true
            If counter.Item("b1") And counter.Item("b2") Then
            
                counter.Item("continueFlg") = True
            
            Else
            
                counter.Item("continueFlg") = False
                
            End If
            
            ' 产生Circle的条件：此时要停止Circle，退出递归
            'If counter.Item("counter") > 99 Then
            
            '    If counter.Item("circle") Then
            '        Debug.Print " ** WARNING ** MAY HAVE A CIRCLE, PLEASE CHECK !!! "
            '    End If
                
            '    counter.Item("circle") = False
                
            '    Set buildAdjaGraph = Nothing
            '    Exit Function
            ' End If
            ''''Debug.Print ">>>>> " & counter.Item("counter2")
            If counter.Item("counter2") > 20 Then
            
                If counter.Item("circle") Then
                    Debug.Print " ** WARNING ** MAY HAVE A CIRCLE, PLEASE CHECK !!! "
                End If
                
                counter.Item("circle") = False
                
                'Set buildAdjaGraph = d
                
                Exit For
                
            End If
            
            'counter.Item("counter") = counter.Item("counter") + (counter.Item("b1") And counter.Item("b2")) * -1
             
            ' 找到依赖关系的话，再次往下找（递归）
            buildAdjaGraph CStr(o), ws, D, counter
        Next
        
    End If
        
    Set buildAdjaGraph = D

End Function

' breadth first search - AdjacencyListGraph
Public Function BFS(adjacencyList As Object, aHead As String) As Collection

    ' bfs用队列，为展示方便使用的父节点队列
    Dim queue As New Collection, pqueue As New Collection
    
    Dim tmpRs As New Collection, resultset As New Collection
    

    ' 临时变量
    Dim head As Variant, phead As Variant, maxHeadLen As Byte, maxPHeadLen As Byte, o
    
    '初始化
    maxHeadLen = 0
    
    maxPHeadLen = 0
    
    '-- bfs(Breadth First Search)
    queue.Add (aHead) ' 队列，用以遍历邻接图
    pqueue.Add ("⊙") ' 存放父节点
    
    Do While queue.Count > 0
    
        ' 取得并删除首节点
        head = queue.Item(1)
        phead = pqueue.Item(1)
        queue.Remove (1)
        pqueue.Remove (1)
        
        ' 长度最大值保存起来，用以格式化打印
        If Len(phead) > maxPHeadLen Then
            maxPHeadLen = Len(phead)
        End If
        
        If Len(head) > maxHeadLen Then
            maxHeadLen = Len(head)
        End If
        
        'Debug.Print "[ " & phead & " -> " & head & " ]"
        'resultset.Add "[ " & phead & " -> " & head & " ]"
        
        ' 放入集合
        tmpRs.Add phead & "," & head
        '
        
        'a.Item(head).keys
        If adjacencyList.Exists(head) Then
        
            For Each o In adjacencyList.Item(head).Keys
            
                If adjacencyList.Item(head).Item(o) = BLANK Then
                
                    queue.Add (o)     ' 存当前节点
                    pqueue.Add (head) ' 存对应的父节点
                    adjacencyList.Item(head).Item(o) = "visited" ' 标示为已遍历
                    
                End If
               
                'If Not okList.exists(o) Then
                '    queue.Add (o)     ' 存当前节点
                '    pqueue.Add (head) ' 存对应的父节
                '    okList.Add o, "visited"  ' 标示为已遍历
                'End If
                
            Next
        End If

    Loop
    
    ' 注释信息（head）
    resultset.Add WorksheetFunction.Rept("=", maxPHeadLen) & " adjacency - list - bfs " & WorksheetFunction.Rept("=", maxHeadLen)
    resultset.Add BLANK
    For Each o In tmpRs
    
        phead = Mid(o, 1, InStr(o, ",") - 1)
        
         head = Mid(o, InStr(o, ",") + 1)
         
        resultset.Add "[ " _
                 & phead _
                 & WorksheetFunction.Rept(SPACE, maxPHeadLen - Len(phead)) _
                 & " -> " _
                 & head _
                 & WorksheetFunction.Rept(SPACE, maxHeadLen - Len(head)) _
                 & " ]"
    
    Next
    
    ' 注释信息（rear）
    resultset.Add BLANK
    resultset.Add WorksheetFunction.Rept("=", maxPHeadLen) & " adjacency - list - bfs " & WorksheetFunction.Rept("=", maxHeadLen)
    
    Set BFS = resultset
    
End Function

' 上传到azkaban web server
' 项目名称，web ip地址，用户名，密码，要上传的zip文件名(绝对路径)，（curl地址，工作目录  ==> 暂时不用）
Public Sub upload2web(meta As UploadMeta)
                                            
   Dim WSH As Object, wExec As Object, result
   
   Dim json As Dictionary, sessionid As String, errmsg As Variant
   
   Dim gensessionCURL As String, uploadCURL As String
   Dim params As New Dictionary
   
   params.Add "username", meta.username
   params.Add "password", meta.password
   params.Add "ip", meta.ip
   params.Add "port", meta.port
   params.Add "project", meta.project
   params.Add "zipfile", meta.zipfile
   
   ' "cmd /c curl -k -X POST --data ""action=login&username=dashuju&password=dashuju"" https://118.26.169.161:8443/manager"
   gensessionCURL = "curl -k -X POST --data ""action=login&username=${username}&password=${password}"" https://${ip}:${port}/manager"
   
   uploadCURL = "curl -k -i -H ""Content-Type: multipart/mixed"" -X POST --form ""session.id=${sessionid}"" --form ""ajax=upload"" --form ""file=@${zipfile};type=application/zip"" --form ""project=${project};type=plain"" https://${ip}:${port}/manager"""
   

   
    ' 使用wscript对象进行外度调用
    Set WSH = CreateObject("WScript.Shell")
    
    Debug.Print antiParametrization(gensessionCURL, params)
    
    ' 开始产生sessionid处理
    Set wExec = WSH.Exec("cmd /c " & antiParametrization(gensessionCURL, params))
    
    result = wExec.StdOut.ReadAll
    
    ' 解析结果
    Set json = JsonConverter.ParseJson(CStr(result))
    
    errmsg = json("error")
    
    If (errmsg <> Empty) Then
        Err.Raise 1, , "CURL_UPLOAD_01, NO SESSION ID FOUND."

    End If
    
    params.Add "sessionid", json("session.id")
    
    ' 开始上传
    Set wExec = WSH.Exec("cmd /c " & antiParametrization(uploadCURL, params))

    ' 上传结束
    result = wExec.StdOut.ReadAll
    
    
    ' Deallocate Resource
    Set WSH = Nothing
    Set wExec = Nothing
    
    Logger.LogInfo json("session.id")
    Logger.LogInfo CStr(result)

End Sub

' 反参数化
Public Function antiParametrization(ByRef pstr As String, ByRef D As Dictionary) As String

    Dim s As String
    
    s = pstr
    
    For Each o In D.Keys
        s = Replace(s, "${" & o & "}", D.Item(o))
    Next
    
    antiParametrization = s
    
End Function

Public Sub printfAdjaColl()

    
End Sub

Public Function int2ip(intip As Integer) As String
    
    Dim binstr As String
    
    binstr = WorksheetFunction.Dec2Bin(intip, 32)
    
    int2ip = WorksheetFunction.Bin2Dec(Mid(binstr, 1, 8)) & "." & _
    WorksheetFunction.Bin2Dec(Mid(binstr, 9, 8)) & "." & _
    WorksheetFunction.Bin2Dec(Mid(binstr, 17, 8)) & "." & _
    WorksheetFunction.Bin2Dec(Mid(binstr, 25, 8))
    
End Function

' demo
Public Function ip2int(strip As String) As Integer
    
    Dim binstr As String ' 169.168.1.4
    
    

    
End Function

Public Function rpad(t As String, pad As String, Optional lens As Byte = 50) As String
    rpad = t & WorksheetFunction.Rept(pad, lens - 1 - VBA.Len(t)) & "|"
End Function

' 取得"/"分隔的字符串中，最大的值
Public Function maxLevel(vtar As String, Optional delim As String = "/") As Integer

    Dim tarArr() As String, vmax As Integer
    
    tarArr = VBA.Split(vtar, delim)
    
    For Each o In tarArr
        
        If CInt(o) > vmax Then
            vmax = CInt(o)
        End If
    Next
    
    maxLevel = vmax
    
End Function

Public Function getFileList(s As String) As Collection

    Dim a As New Collection
    
    ' 确保以\结果
    mypath = SWITCH(Right(s, 1) <> "\", s & "\", Right(s, 1) = "\", s)
    
    myname = DIR(mypath, VBA.vbDirectory)
    
    
    Do While myname <> ""
    
        If myname <> "." And myname <> ".." Then
        
            a.Add (mypath & myname)
            
        End If
        
        myname = DIR
        
    Loop
    
    Set getFileList = a
    
End Function

' 除去comment，这个为了适应hive（hive不支持/* */类型的comment）
Public Function removeAssignedComment(ByRef v As String, Optional CommentStyle As String = "/*")

    Dim t As String
    Dim c_posl As Long, c_posr As Long
    Dim s_left As String, s_right As String
    
    t = v
    
    ' 初始化位置
    c_posl = InStr(1, t, "/*")
    c_posr = InStr(1, t, "*/")
    
    Do While c_posl > 0 And c_posr > 0
        
        s_left = VBA.Left(t, c_posl - 1)
        
        s_right = VBA.Right(t, VBA.Len(t) - c_posr - 1)
        
        t = s_left & rpad(" ", " ", c_posr + 2 - c_posl) + s_right
        
        '"a/* abc */b" => l = 2, r = 9 r - l =
        c_posl = VBA.InStr(1, t, "/*")
        c_posr = VBA.InStr(1, t, "*/")
        
    Loop
    
    removeAssignedComment = t
    
End Function

Public Function addToDict(d As Dictionary, k As String, v As String) As Boolean

    If Not d.Exists(k) Then
        d.Add k, v
        addToDict = True
    Else
        addToDict = False
    End If
    
End Function
