VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdjaListProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************************************
'*
'*'''''' 按照找下游（START -> START_PL -> TRANS_1_PL_S_XXX -> END）的方式，
'*
'*'''''' 使用指定的Sheet，建立起一张邻接表
'*
'* 使用前，需要 < 人 为 修 改 > 的变量：
'*
'*                             1) 当前读取的sheet索引（CURR_SHEET_IDX，【默认：9】  ）
'*
'*                             2) 始节点所在位置，头节点即START（POSITION_ROW、POSITION_COL，【默认：2,1】）
'*
'*                             3) 连续发生「要查找的元素已存在」次数，用以判断是否产生循环依赖（CIRCLE_JUDGE_MAX，【默认：100 * 50】） 【不要，删除】
'*
'*                             4) 使用{level判断}功能开关（SWITCH，【默认：开启】）
'*
'* Mod By Lujx 2015/12/28
'*
'********************************************************************************************
'
'
'
' 存储邻接图使用的Dictionary
'
'-------------------------------------------------------------------
'|                ADJACENCY - LIST DICTIONARY                      |
'------------------------------------------------------------------|
'|      KEY   |      VALUE                                         |
'|------------|----------------------------------------------------|
'|     Node1  |      Sub_Dict1<                                    |
'|            |                down_stream_job1 , VISITED_FLAG     |
'|            |                down_stream_job2 , VISITED_FLAG     |
'|            |               >                                    |
'|------------|----------------------------------------------------|
'|     Node2  |      Sub_Dict2<                                    |
'|            |                down_stream_job3 , VISITED_FLAG     |
'|            |                down_stream_job4 , VISITED_FLAG     |
'|            |               >                                    |
'|------------|----------------------------------------------------|
'|     Node3  |      Sub_Dict3<                                    |
'|            |                down_stream_job5 , VISITED_FLAG     |
'|            |                down_stream_job6 , VISITED_FLAG     |
'|            |               >                                    |
'|------------|----------------------------------------------------|
'
Public dict As Object


'
' 存储节点所在层级（深度），一个节点可能隶属于多个层次
'
'-------------------------------------------------------------------
'|                    JOBS - LEVEL DICTIONARY                      |
'------------------------------------------------------------------|
'|      KEY   |      VALUE                                         |
'|------------|----------------------------------------------------|
'|     Node1  |        2/3/4/5                                     |
'|------------|----------------------------------------------------|
'|     Node2  |        5                                           |
'|------------|----------------------------------------------------|
'|     Node3  |        6/7                                         |
'|------------|----------------------------------------------------|
'
Public levelDict As Object

    ' 标记为已出列（同时记录节点访问次数）
Public outQ As Object

    ' 已在队列中或已处理（已放入队列）
Public inQ As Object

Public dhead As String

'''
Private Const CURR_SHEET_IDX = 3 '18

Private Const POSITION_ROW = 181 '2

Private Const POSITION_COL = 1

' 使用的队列
Private queue As Collection


' 初始化
Private Sub Class_Initialize()

    If outQ Is Nothing Then
    
        Set outQ = CreateObject("Scripting.Dictionary")
        
    End If
    
    If outQ Is Nothing Then
    
        Set outQ = CreateObject("Scripting.Dictionary")
    
    End If
    
    Set queue = New Collection
    
End Sub

' *********************************************
' *
' * 使用「控制入队」处理方法
' *
' * 也可使用「控制出队」的处理方法实现，待添加
' *
' *********************************************
Public Sub process(target As Worksheet, r As Integer, c As Integer)
    
    Dim tmpColl As Collection
    
    Dim cur_sheet As Worksheet
    
    Dim last As String
        
    Dim start As String
    
    ' 用以存储从last分离出的节点（adjacency list node）及对应层级（level）的临时变量
    Dim node As String, level As Byte
    
    Dim o As Variant, stdo As String
    
    Set cur_sheet = target
    
    ' 获取开头节点：一般是第二行第一个节点，即START
    start = cur_sheet.Cells(r, c)
    'start = "TRANS_4_S_PL_CRM_OVERDUE_DK_H"
    
    dhead = start
    
    ' 先把start加入到队列
    queue.Add stdNode(start, 1, COMMA)  ' FORMAT => "Trans,1"
    
    ' 加入「已入队」
    putInQ stdNode(start, 1, COMMA)
    
    ' 遍历队列
    Do While queue.Count > 0
    
                '' just for debug
                printseq = ""
                
                For Each e In queue
                    printseq = printseq & " [ " & e & " ] "
                Next e
                '' debug ok
                
                Logger.LogDebug CStr(printseq)
                
                ' 取出并删除最后一个元素
                last = queue.Item(1)
                queue.Remove (1)
                
                node = Split(last, COMMA)(0)
                level = CByte(Split(last, COMMA)(1))
      
                ' 找出last节点对应的下游节点
                Set tmpColl = findDeped(node, cur_sheet)
                
                ' 遍历下游节点，放入队列进行处理
                For Each o In tmpColl
                
                    stdo = stdNode(o, level + 1, COMMA)
                    
                    ' 不在inQ时才加入队列，不重复加入队列，确保效率
                    If Not inQ.Exists(stdo) Then
                    
                        ' 放入队列，同时放入「已入队」集合，标记为已入队
                        queue.Add stdo
                        putInQ stdo
                        
                    End If
                    
                    ' 加入邻接表
                    addToDict node, CStr(o)
            
                Next
                
                ' 找不到下游任务时，通用处理：插入节点，但其下游为空对象（参考addtodict）
                If tmpColl.Count = 0 Then
                        addToDict node, Empty
                End If
                    
                
                ' 更新层级
                addLevelToDict node, level
                
                ' 将节点标记为已访问，并更新节点访问次数 + 1
                putOutQ stdNode(node, level, COMMA)

    Loop
        
End Sub

' 将元素添加到邻接表 = 既存返回true, 添加成功返回false（）
Private Function addToDict(k As String, v As String) As Boolean

        Dim subdict As Object
        
        Dim exist As Boolean
        
        exist = False
        
        If dict Is Nothing Then
        
            Set dict = CreateObject("Scripting.Dictionary")
        
        End If
        
        If Not dict.Exists(k) Then
        
            Set subdict = CreateObject("Scripting.Dictionary")
            
            If v <> Empty Then
                subdict.Add v, BLANK
            End If
            
            dict.Add k, subdict
            
        Else
        
             Set subdict = dict.Item(k)
             
             If Not subdict.Exists(v) Then
                
                If v <> Empty Then
                    subdict.Add v, BLANK
                End If
                
             Else
             
                exist = True
                
             End If

        End If
        
        addToDict = exist
        
End Function

' 加入到层级，返回插入时节点是否已存在
Private Function addLevelToDict(k As String, l As Byte) As Boolean
        
        Dim exist As Boolean
        
        exist = False
        
        ' 存储层级
        If levelDict Is Nothing Then
        
            Set levelDict = CreateObject("Scripting.Dictionary")
        
        End If
        
        If Not levelDict.Exists(k) Then

            ' 层级加入到levelDict
            levelDict.Add k, l
            
        Else
        
            exist = True
        
             '''''' add level start ''''''
             Dim values() As String
             Dim concat As Boolean
             
             values = Split(levelDict.Item(k), "/")
             concat = True
            
             For Each o In values
                If o = l Then
                    concat = False
                    Exit For
                End If
             Next
            
             If concat Then
             
                levelDict.Item(k) = levelDict.Item(k) & "/" & l
             
              End If
              '''''' add level end ''''''

        End If
        
        addLevelToDict = exist
        
End Function

' 更新节点访问次数 + 1
' 返回当前访问次数
Private Function putOutQ(ByVal k As String) As Integer
    
    If Not outQ.Exists(k) Then
        outQ.Add k, 1
        
    Else
    
        ' 访问次数 + 1
        outQ.Item(k) = outQ.Item(k) + 1
    
    End If
    
    putOutQ = CInt(outQ.Item(k))
    
End Function


' 加入「已入队」集合
Private Function putInQ(ByVal k As String) As Integer
    
    If inQ Is Nothing Then
    
        Set inQ = CreateObject("Scripting.Dictionary")
    
    End If
    
    If Not inQ.Exists(k) Then
        inQ.Add k, 1
        
    Else
    
        ' 访问次数 + 1
        inQ.Item(k) = inQ.Item(k) + 1
    
    End If
    
    putInQ = CInt(inQ.Item(k))
    
End Function



Public Function head() As String

 head = dhead ' Worksheets(CURR_SHEET_IDX).Cells(POSITION_ROW, POSITION_COL)
 
End Function

Public Function rear() As Collection

    Dim endnodes As New Collection
    Dim o1, o2
    
    For Each o1 In dict.Keys
    
        For Each o2 In dict.Item(o1).Keys
            
            If Not dict.Exists(o2) Then
                 endnodes.Add o2
            End If
        Next
        
    Next
    
    Set rear = endnodes

End Function

' 返回节点数量
Public Function Count() As Integer

    Count = dict.Count
    
End Function

' 遍历元素
Public Function elements() As Collection

    Dim res As New Collection
    
    Dim pos As Integer, addflg As Boolean
    

    ' 冒泡法排序（ 大 -> 小 ）
    For Each o In dict.Keys
    
        pos = 1
        addflg = True
        
        
        For Each io In res
        
            If o > io Then
            
                           
             res.Add o, , pos
            
             addflg = False
             
            Exit For

            Else
                pos = pos + 1
            End If

        Next
        
        If addflg Then
            res.Add o
        End If
        

        Set elements = res
        
    Next

End Function

' breath first search traversal
Public Function BFS() As Collection

    Set BFS = CommonUtil.BFS(dict, head)
    
End Function

' print bfs detail
Public Sub printBFS()

        Dim resultset As Collection
    
        Set resultset = BFS
    
        For Each o In resultset
        
            'Debug.Print o
            Logger.LogTrace " - " & " - " & o
        Next
End Sub

' print dict detail
Public Sub printDict()

        For Each ko In dict.Keys()
        
            'Debug.Print " [  -k-  ] " & ko
            Logger.LogTrace " - " & " - [  -k-  ] " & ko
            For Each vo In dict.Item(ko).Keys()
                
                'Debug.Print " [ --v-- ] " & vo
                Logger.LogTrace " - " & " - [ --v-- ] " & vo
            Next

    Next
End Sub


Public Sub printLEVEL()

    For Each o In levelDict.Keys
    
        Logger.LogInfo " [ " & o & " >>> " & levelDict.Item(o) & " ] "
        
    Next
End Sub

Public Sub printAccessTimes()

    For Each o In outQ.Keys
    
        Logger.LogDebug " [ " & o & " >>> " & outQ.Item(o) & " ] "
        
    Next
End Sub


' 使用矩阵打印图形
Public Sub printf_DIAGRAM_MATRIX()


End Sub

' 打印按照不同level（depth深度）分组的结果
Public Sub print_LEVEL_DIMENSION()

    Dim levelDIM As Object
    Dim levelNDS As Collection
    
    ' 矩阵长宽 = 列数，行数 = 最大层级数， 最多元素数
    Dim mLength As Byte, mWidth As Byte
    
    Dim lens As Byte ' 取得最长节点对应的长度
    
    Set levelDIM = CreateObject("Scripting.Dictionary")
    
    For Each o In levelDict.Keys
        
        Dim values() As String, maxv As Integer
        
        values = Split(levelDict.Item(o), "/")
        
        ' 某些情况下，最后一个值并不是最大值（取好几个头结点时），此时要判断取出最大值才行
        ' 2016/01/11 lujx
        'maxv = values(UBound(values))
        maxv = CommonUtil.maxLevel(CStr(levelDict.Item(o)))
        
        If Not levelDIM.Exists(maxv) Then
                    
            Set levelNDS = New Collection
            levelNDS.Add o
            levelDIM.Add maxv, levelNDS
            
        Else
        
            levelDIM.Item(maxv).Add o
            
        End If
        
        ' 找出最大长度作为矩阵的长（最大level）
        If CByte(maxv) > mLength Then
            mLength = CByte(maxv)
        End If
        
        ' 找出最大宽度作为矩阵的宽（job最多的level对应的job数量）
        If CByte(levelDIM(maxv).Count) > mWidth Then
            mWidth = CByte(levelDIM(maxv).Count)
        End If
        
        ' 取得单节点最长的长度
        If Len(o) > lens Then
            lens = Len(o)
        End If

    Next
    
    ' ---------------------------------------------------------------------
    '
    ' 挪到上面，废弃code
    ' For Each o In levelDIM.Keys
    
        ' 统计某level中job数量，计数
    '     mCounter = 0
    '     For Each v In levelDIM.Item(o)
    '         Logger.LogInfo CStr(v), CStr(o)
    '         mCounter = mCounter + 1
    '     Next v
        
        ' 找出最大宽度作为矩阵的宽（job最多的level对应的job数量）
    '     If CByte(mCounter) > mWidth Then
    '         mWidth = CByte(mCounter)
    '     End If

        ' 找出最大长度作为矩阵的长（最大level）
    '     If CByte(o) > mLength Then
    '         mLength = CByte(o)
    '     End If
    ' Next o
    '
    ' ------------------------------------------------------------------------
    
    ' 打印矩阵（行 * 列）
    Dim r As Byte, c As Byte, rowstr
    
    ' lens = getNodeMaxLen() + 2
    lens = lens + 2 ' 确保长度更大
    
    For r = 1 To mWidth Step 1 ' 行数
        rowstr = Empty
        For c = 1 To mLength Step 1
            If levelDIM(c).Count >= r Then
                 rowstr = rowstr & rpad(CStr(levelDIM(c)(r)), SPACE, lens)
            Else
                 rowstr = rowstr & rpad(SPACE, SPACE, lens)
            End If
        Next c
        Logger.LogInfo CStr(rowstr)
        'Logger.LogInfo CStr(WorksheetFunction.Rept("-", lens * mLength))
    Next r
    
End Sub


' 元素是否已在队尾
Private Function isRear(e As String, l As Byte) As Boolean

    ' 空队列
    If queue.Count = 0 Then
    
        isRear = False
    
    Else
    
       ' isRear = (e = Trim(Split(queue.Item(queue.Count), ",")(0)) And l = CByte(Trim(Split(queue.Item(queue.Count), ",")(1))))
        
         isRear = (e = Trim(Split(queue.Item(queue.Count), ",")(0)))
        
    End If
    

End Function

' 将节点n和层级l（深度）使用delim组合 ==> "START,1"
Private Function stdNode(n As Variant, l As Byte, delim As String) As String

    stdNode = n & delim & CStr(l)
    
End Function

' 返回节点最大长度，便于打印结果
Private Function getNodeMaxLen() As Byte

    Dim r_max As Byte
    
    For Each e In dict
        If r_max < Len(e) Then
            r_max = Len(e)
        End If
    Next e
    
    getNodeMaxLen = r_max
    
End Function

' destroy
Private Sub Class_Terminate()

End Sub



