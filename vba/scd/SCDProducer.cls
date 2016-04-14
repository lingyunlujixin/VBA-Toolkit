VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SCDProducer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************************************
'*
'* 名称：
'*
'* 功能：
'*
'* 时间：
'*
'* 修改记录：
'*
'*******************************************************************************************************

'-------------------------------------------------
'『变量替换列表』
'
' 存储预定义的【变量】及【对应值】,用以处理合成的
'
' sql中的变量(${VAR})
'-------------------------------------------------
'
' Dict - <"${MY_DT}"        , "'${MY_DT}'"    >
'        <"${MAX_DT}"       , "'${MAX_DT}'"   >
'        <"${V_CAT}"        , "test.vt_xx"    >
'        ........
'
'-------------------------------------------------
  Private varDict As Dictionary

'-------------------------------------------------
' 存储表描述信息，分两部分：
'
'    --  表信息域（以#开头）
'
'    --  字段信息域
'-------------------------------------------------
'
' Dict - <"#SRC_TABLE_NAME"   , Table_Name  >
'        <"#SRC_TABLE_SCHEMA" , SCHEMA_VAR  >
'        <"#TAR_TABLE_NAME"   , Table_Name  >
'        <"#TAR_TABLE_SCHEMA" , SCHEMA_VAR  >
'        <"#SRC_TABLE"        , SCHEMA.TABLE>
'        <"#TAR_TABLE"        , SCHEMA.TABLE>
'        <FieldName           , FIELD_INFO  >
'        ........
'
'
'-------------------------------------------------
Private tableMetaDataDict As Dictionary

' 记录各分类字段（主键，用来增量的字段，保持不变（更新）的字段）
Private tableKeyFileds As Collection
Private tableInsertFileds As Collection
Private tableUpdateFileds As Collection

' 初始化
Private Sub Class_Initialize()
    
    Set tableMetaDataDict = New Dictionary
    
    Set varDict = New Dictionary
    
End Sub

' 初始化『变量替换列表』
Private Sub preVarDict()
    
    ' 源表
    varDict.Add "${SRC_TABLE}", tableMetaDataDict(SRC_TABLE)
    varDict.Add "${SRC_TABLE_SCHEMA}", tableMetaDataDict(SRC_TABLE_SCHEMA)
    varDict.Add "${SRC_TABLE_NAME}", tableMetaDataDict(SRC_TABLE_NAME)
    
    ' 目标表
    varDict.Add "${TAR_TABLE}", tableMetaDataDict(TAR_TABLE)
    varDict.Add "${TAR_TABLE_SCHEMA}", tableMetaDataDict(TAR_TABLE_SCHEMA)
    varDict.Add "${TAR_TABLE_NAME}", tableMetaDataDict(TAR_TABLE_NAME)
    varDict.Add "${TAR_TABLE_BAK}", tableMetaDataDict(TAR_TABLE) & "_bak"
    
    ' 临时表
    varDict.Add "${V_CAT}", stdVCatTbl(tableMetaDataDict(SRC_TABLE_NAME), tableMetaDataDict(TAR_TABLE_SCHEMA))
    varDict.Add "${V_TAR}", stdVTarTbl(tableMetaDataDict(TAR_TABLE_NAME), tableMetaDataDict(TAR_TABLE_SCHEMA))
    
    ' 日期类
    varDict.Add "${MAX_DT}", tableMetaDataDict(MAX_DT)
    varDict.Add "${MY_DT}", tableMetaDataDict(MY_DT)
    varDict.Add "${LITERAL_MY_DT}", "${MY_DT}"
     
    ' 其他
    varDict.Add "${MYSQL_COMMENT}", "${MYSQL_COMMENT}"
    
    ' 为适应hive
    varDict.Add "/* 'closed' */", rpad(" ", " ", VBA.Len("/* 'closed' */"))
    varDict.Add "/* 'unchanged' */", rpad(" ", " ", VBA.Len("/* 'unchanged' */"))
    varDict.Add "/* 'changed' */", rpad(" ", " ", VBA.Len("/* 'changed' */"))
    varDict.Add "/* 'insert' */", rpad(" ", " ", VBA.Len("/* 'insert' */"))
    
End Sub

' 从指定的excel表格初始化表描述信息
Public Sub setExcel(target As Worksheet, r As Integer, c As Integer)

    Dim curWkst As Worksheet
    
    Dim irow As Integer
    
    Dim FieldName As String, FieldType As String, FieldOperType As String, FieldComment As String
    
    Dim td As Dictionary
    
    Set curWkst = target
    
    ' 添加表名
    'tableMetaDataDict.Add SRC_TABLE_NAME, VBA.Split(curWkst.Name, ",")(0)

    ' 添加表名(SRC)
    tableMetaDataDict.Add SRC_TABLE, VBA.Trim(curWkst.Cells(1, 7).Value)
    tableMetaDataDict.Add SRC_TABLE_SCHEMA, VBA.Split(curWkst.Cells(1, 7).Value, ".")(0)
    tableMetaDataDict.Add SRC_TABLE_NAME, VBA.Split(curWkst.Cells(1, 7).Value, ".")(1)

    ' 添加表名(TAR)
    tableMetaDataDict.Add TAR_TABLE, VBA.Trim(curWkst.Cells(2, 7).Value)
    tableMetaDataDict.Add TAR_TABLE_SCHEMA, VBA.Split(curWkst.Cells(2, 7).Value, ".")(0)
    tableMetaDataDict.Add TAR_TABLE_NAME, VBA.Split(curWkst.Cells(2, 7).Value, ".")(1)
    
    ' 添加日期参数
    tableMetaDataDict.Add MAX_DT, "cast('" & curWkst.Cells(3, 7).Value & "' as date)"
    tableMetaDataDict.Add MY_DT, "cast('" & curWkst.Cells(4, 7).Value & "' as date)"
    
    ' 初始化变量字典
    preVarDict
    
    ' 遍历开始行数
    irow = r
    
    ' 遍历
    Do While curWkst.Cells(irow, c) <> Empty
    
        ' 字段名称
        FieldName = VBA.Replace(curWkst.Cells(irow, c), "`", "")
        
        ' 字段类型
        FieldType = curWkst.Cells(irow, c + 1)
        
        ' 操作类型
        FieldOperType = curWkst.Cells(irow, c + 2)
        
        ' 注释
        FieldComment = curWkst.Cells(irow, c + 3)
        
        Set td = New Dictionary
        
        td.Add "FIELD_NAME", FieldName
        td.Add "FIELD_TYPE", VBA.UCase(FieldType)
        td.Add "FIELD_COMMENT", FieldComment
        td.Add "FIELD_OPER_TYPE", VBA.UCase(FieldOperType)
        
        'tableMetaDataDict.Add FieldName, getStates(FieldType, FieldOperType)
        tableMetaDataDict.Add FieldName, td
        
        irow = irow + 1
    Loop
    
    ' debug
    For Each o In tableMetaDataDict.keys
        'Logger.LogDebug CStr(o), CStr(tableMetaDataDict(o))
    Next

End Sub

'  返回所有的主键
Public Function keySet() As Collection

    If Not tableKeyFileds Is Nothing Then
        Set keySet = tableKeyFileds
        Exit Function
    End If
    
    Dim keys As New Collection
    
    For Each o In tableMetaDataDict.keys
    
        'If (tableMetaDataDict(o) And InterfaceStates.OPER_TYPE_KEY) > 0 Then
        '    keys.Add o
        'End If
        
        
        If VBA.Left(o, 1) <> "#" Then
            If tableMetaDataDict(o)("FIELD_OPER_TYPE") = "KEY" Then
                keys.Add o
            End If
        End If
    Next o
    
    Set tableKeyFileds = keys
    
    Set keySet = tableKeyFileds
    
End Function


'  返回所有的Insert字段（用来拉链的字段）
Public Function insertSet() As Collection

    If Not tableInsertFileds Is Nothing Then
        Set insertSet = tableInsertFileds
        Exit Function
    End If
    
    Dim keys As New Collection
    
    For Each o In tableMetaDataDict.keys
    
        'If (tableMetaDataDict(o) And InterfaceStates.OPER_TYPE_INSERT) > 0 Then
        '    keys.Add o
        'End If
        If VBA.Left(o, 1) <> "#" Then
            If tableMetaDataDict(o)("FIELD_OPER_TYPE") = "INSERT" Then
                keys.Add o
            End If
        End If
        
    Next o
    
    Set tableInsertFileds = keys
    Set insertSet = keys
    
End Function

'  返回所有的Update字段（不用来拉链的字段）
Public Function updateSet() As Collection

    If Not tableUpdateFileds Is Nothing Then
        Set updateSet = tableUpdateFileds
        Exit Function
    End If
    
    Dim keys As New Collection
    
    For Each o In tableMetaDataDict.keys
    
        'If (tableMetaDataDict(o) And InterfaceStates.OPER_TYPE_UPDATE) > 0 Then
        '    keys.Add o
        'End If
        If VBA.Left(o, 1) <> "#" Then
                If tableMetaDataDict(o)("FIELD_OPER_TYPE") = "UPDATE" Then
                    keys.Add o
                End If
        End If
    Next o
    
    Set tableUpdateFileds = keys
    Set updateSet = keys
    
End Function

'  返回所有的insert & Update字段
Public Function upsertSet() As Collection
    
    Dim keys As New Collection
    
    For Each o In tableMetaDataDict.keys
    
        If VBA.Left(o, 1) <> "#" Then
                If tableMetaDataDict(o)("FIELD_OPER_TYPE") <> "KEY" Then
                    keys.Add o
                End If
        End If
    Next o
    
    Set upsertSet = keys
    
End Function


' 产生拉链程序使用的SQL
Public Sub start(FileName As String)

    Dim literal As String
    Dim textCache As New Collection
    
    
    ' 1.01 生成创建Catlog表结构的SQL
    literal = genCat_Table()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 1.02 初始化Truncate
    literal = genCat_Clear()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 1.03 找出标记为insert的记录
    literal = genCat_Insert()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 1.04 找出标记为closed，changed，unchanged的记录
    literal = genCat_Closed_Changed_Unchanged()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 2.01 将Insert的记录插入目标表（临时目标表）
    literal = genUnion_Insert()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 2.02 将历史数据，闭链，封链，未变化（可能是更新，也可能是一点无变化）的数据插入目标表（临时目标表）
    literal = genUnion_His_Closed_Changed_Unchanged()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 2.03 将changed的记录封链后，再次插入目标表（临时目标表）
    literal = genUnion_NewChanged()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' 2.04 刷新目标表~
    literal = getUnion_Target()
    Logger.LogDebug literal
    textCache.Add replaceVar(literal)
    
    ' persist to file
    persistence textCache, FileName
    
    'OK了
    ' 将以上生成的SQL文存储到文件即可
    
End Sub

'**********************************************************
'*
'* 生成使用的临时表
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genCat_Table() As String

    Dim keyDef As String
    For Each o In keySet()
    
        keyDef = Switch(keyDef = Empty, Empty, keyDef <> Empty, keyDef & vbCrLf) _
                 & formatColumn(CStr(o), CStr(tableMetaDataDict(o)("FIELD_TYPE")) _
                            , False, CStr(tableMetaDataDict(o)("FIELD_COMMENT")), "") _
                 & ","
        
    Next o
    
    genCat_Table = vbCrLf _
                    & WorksheetFunction.Rept("#", 50) _
                    & vbCrLf _
                    & vbCrLf _
                    & "-- SPECIFIC TABLE - CREATION " _
                    & vbCrLf _
                    & "CREATE TABLE    IF NOT EXISTS ${V_CAT} (" _
                    & vbCrLf _
                    & keyDef _
                    & vbCrLf _
                    & formatColumn("typ", "TINYINT", False, "1 - INSERT/2 - CLOSED/3 - UNCHANGED/4 - CHANGED", "-1") & "," _
                    & vbCrLf _
                    & formatColumn("ver", "INT", False, "VERSION NUMBER", "1") & "," & vbCrLf _
                    & formatPrimaryKey(keySet()) & "," _
                    & vbCrLf _
                    & formatKey(keySet()) _
                    & vbCrLf _
                    & " ) ENGINE=INNODB DEFAULT CHARSET=utf8 " _
                    & vbCrLf _
                    & " PARTITION BY HASH(typ) PARTITIONS 4;" _
                    & vbCrLf _
                    & vbCrLf _
                    & "CREATE TABLE IF NOT EXISTS ${V_TAR} LIKE ${TAR_TABLE};" _
                    & vbCrLf & vbCrLf _
                    & WorksheetFunction.Rept("#", 50)

End Function

'**********************************************************
'*
'* 如果临时表已存在，使用前确保临时表清空
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genCat_Clear() As String

    genCat_Clear = vbCrLf _
                    & "-- Clear temporary table" _
                    & vbCrLf _
                    & "TRUNCATE TABLE ${V_CAT};" _
                    & vbCrLf _
                    & "TRUNCATE TABLE ${V_TAR};"
End Function

'**********************************************************
'*
'* 分类阶段：找出标记为insert（新增）的记录
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genCat_Insert() As String

    Dim selectKeys As String, joinKeys As String, whereKeys As String
    
    For Each o In keySet()
        selectKeys = Switch(selectKeys = Empty, vbTab & "  a." & o, selectKeys <> Empty, selectKeys & vbCrLf & vbTab & ", a." & o)
        joinKeys = Switch(joinKeys = Empty, "  ON a." & o & " = b." & o, joinKeys <> Empty, joinKeys & vbCrLf & " AND a." & o & " = b." & o)
        whereKeys = "b." & o & " IS NULL"
    Next o
    
    genCat_Insert = vbCrLf & "-- step 01. Pick out records marked 'insert' ( type = 1 )" _
                    & vbCrLf _
                    & "INSERT INTO ${V_CAT} " _
                    & vbCrLf _
                    & "SELECT" _
                    & vbCrLf _
                    & selectKeys _
                    & vbCrLf _
                    & vbTab & ", 1 -- 'insert'" _
                    & vbCrLf _
                    & vbTab & ", 1 -- 'version 1 for insert-record'" _
                    & vbCrLf _
                    & "FROM      ${SRC_TABLE} a" _
                    & vbCrLf _
                    & "LEFT JOIN ${TAR_TABLE} b" _
                    & vbCrLf _
                    & joinKeys _
                    & vbCrLf _
                    & "AND b.start_dt < ${MY_DT}" _
                    & vbCrLf _
                    & "AND b.end_dt  >= ${MY_DT}" _
                    & vbCrLf _
                    & "WHERE " & whereKeys & ";"

End Function

'**********************************************************
'*
'* 分类阶段：找出标记为closed（物理删除），changed（进行拉链），unchanged（不进行拉链，只更新）的记录
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genCat_Closed_Changed_Unchanged() As String

    Dim selectKeys As String, joinKeys As String, whereKeys As String, casewhen_sub As String, casewhen As String
    
    For Each o In keySet()
        selectKeys = Switch(selectKeys = Empty, vbTab & "  b." & o, selectKeys <> Empty, selectKeys & vbCrLf & vbTab & ", b." & o)
        joinKeys = Switch(joinKeys = Empty, "  ON a." & o & " = b." & o, joinKeys <> Empty, joinKeys & vbCrLf & " AND a." & o & " = b." & o)
        whereKeys = "a." & o & " IS NULL"
    Next o
    
    For Each o In insertSet()

        opad = rpad(CStr(o), " ", 30)
        
        If casewhen_sub = Empty Then
            casewhen_sub = vbTab _
                           & "                ((" & rpad("a." & o, " ", 30) & " = " & rpad("b." & o, " ", 30) & " ) OR (a." & opad & " IS NULL AND b." & opad & " IS NULL ))"
        Else
            casewhen_sub = casewhen_sub _
                           & vbCrLf _
                           & vbTab _
                           & "            AND ((" & rpad(CStr(specialTransform("a." & o, True)), " ", 30) & " = " & rpad(coalesce("b." & o), " ", 30) & " ) OR (a." & opad & " IS NULL AND b." & opad & " IS NULL ))"
        End If
    Next o
    
    casewhen = vbTab & ", CASE WHEN " & whereKeys & "                                " _
             & vbCrLf _
             & vbTab & "       THEN 2 -- 'closed'                                    " _
             & vbCrLf _
             & vbTab & "       WHEN                                                  " _
             & vbCrLf _
             & casewhen_sub & "                                              " _
             & vbCrLf _
             & vbTab & "       THEN     3 -- 'unchanged'                             " _
             & vbCrLf _
             & vbTab & "       ELSE     4 -- 'changed'                               " _
             & vbCrLf _
             & vbTab & "   END                                                       "
             
genCat_Closed_Changed_Unchanged = vbCrLf & "-- step 02. Pick out records marked 'closed'   ( type = 2 )" _
                                & vbCrLf _
                                & "--                                  'unchanged'( type = 3 )" _
                                & vbCrLf _
                                & "--                                  'changed'  ( type = 4 )" _
                                & vbCrLf _
                                & "INSERT INTO ${V_CAT} " _
                                & vbCrLf _
                                & "SELECT" _
                                & vbCrLf _
                                & selectKeys _
                                & vbCrLf _
                                & casewhen _
                                & vbCrLf _
                                & vbTab & ", b.version" _
                                & vbCrLf _
                                & "FROM      ${TAR_TABLE} b" _
                                & vbCrLf _
                                & "LEFT JOIN ${SRC_TABLE} a" _
                                & vbCrLf _
                                & joinKeys _
                                & vbCrLf _
                                & "WHERE b.start_dt   <   ${MY_DT} -- fetch his-data !! support reload operation" _
                                & vbCrLf _
                                & "  AND b.end_dt     >=  ${MY_DT} ; -- fetch his-data !! support reload operation"
End Function

'**********************************************************
'*
'* 合并阶段：把标记为insert（新增）的记录插入目标表
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genUnion_Insert() As String

    Dim res As String
    Dim selectFields As String, joinKeys As String, insertFields As String
    
    For Each o In keySet()
        joinKeys = Switch(joinKeys = Empty, "  ON a." & o & " = b." & o, joinKeys <> Empty, joinKeys & vbCrLf & " AND a." & o & " = b." & o)
    Next
    
    For Each o In tableMetaDataDict.keys
        If VBA.Left(o, 1) <> "#" Then

            If selectFields = Empty Then
                insertFields = vbTab & "  " & "`" & o & "`"
                selectFields = vbTab & "  " & rpad(CStr("a." & o), " ", 30) & " as `" & o & "`"
            Else
                insertFields = insertFields & vbCrLf & vbTab & ", " & "`" & o & "`"
                selectFields = selectFields & vbCrLf & vbTab & ", " & rpad(CStr(specialTransform("a." & o, True)), " ", 30) & " as `" & o & "`"
            End If
            
        End If
    Next
    
    res = vbCrLf & "-- step 03. reduce result(new target table)       " _
        & vbCrLf _
        & "INSERT INTO ${V_TAR} " _
        & vbCrLf _
        & "(" _
        & vbCrLf _
        & insertFields _
        & vbCrLf _
        & " , version       " _
        & vbCrLf _
        & " , start_dt      " _
        & vbCrLf _
        & " , end_dt        " _
        & vbCrLf _
        & " , load_stamp    " _
        & vbCrLf _
        & ")"
                    
    res = res & vbCrLf _
            & "--  insert (from src) " _
            & vbCrLf _
            & "SELECT" _
            & vbCrLf _
            & selectFields _
            & vbCrLf _
            & "  , 1                           AS version      " _
            & vbCrLf _
            & "  , ${MY_DT}                    AS start_dt     " _
            & vbCrLf _
            & "  , ${MAX_DT}                   AS end_dt       " _
            & vbCrLf _
            & "  , CURRENT_TIMESTAMP                           " _
            & vbCrLf _
            & "FROM       ${SRC_TABLE}                        a" _
            & vbCrLf _
            & "INNER JOIN ${V_CAT}                            b" _
            & vbCrLf _
            & joinKeys _
            & vbCrLf _
            & "  AND b.`typ` = 1 ; -- 'insert'"
            
    genUnion_Insert = res
        
End Function

'**********************************************************
'*
'* 合并阶段：把标记为
'*                    ① closed（物理删除）
'*                    ② changed（进行拉链）
'*                    ③ unchanged（不进行拉链，只更新）
'* 的数据以及、
'*                    ④ 历史数据
'* 插入目标表
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genUnion_His_Closed_Changed_Unchanged() As String

    Dim res As String
    Dim selectKeys As String, aJoinBKeys As String, bJoinCKeys As String, whereKeys As String, casewhen_sub As String, casewhen As String
    Dim insertFields As String
    
    For Each o In keySet()
        aJoinBKeys = Switch(aJoinBKeys = Empty, "  ON a." & o & " = b." & o, aJoinBKeys <> Empty, aJoinBKeys & vbCrLf & " AND a." & o & " = b." & o)
        bJoinCKeys = Switch(bJoinCKeys = Empty, "  ON b." & o & " = c." & o, bJoinCKeys <> Empty, bJoinCKeys & vbCrLf & " AND b." & o & " = c." & o)
        selectKeys = Switch(selectKeys = Empty, vbTab & "  a." & rpad(CStr(o), " ", 30) & " AS `" & o & "`", selectKeys <> Empty, selectKeys & vbCrLf & vbTab & ", a." & rpad(CStr(o), " ", 30) & " AS `" & o & "`")
        whereKeys = "b." & o & " IS NULL"
    Next
    
    For Each o In tableMetaDataDict.keys
    
        If VBA.Left(o, 1) <> "#" Then
            
            If insertFields = Empty Then
                insertFields = vbTab & "  `" & o & "`"
            Else
                insertFields = insertFields & vbCrLf & vbTab & ", `" & o & "`"
            End If
            
        End If
    Next
    
    For Each o In upsertSet()
        
        casewhen = casewhen _
                  & vbCrLf & vbTab & ", CASE WHEN " & rpad(whereKeys, " ") & VBA.Trim("                         ") & " THEN a." & o _
                  & vbCrLf & vbTab & "       WHEN " & rpad("b.typ  = 3  /* 'unchanged' */                  ", " ") & " THEN " & specialTransform("c." & o, True) _
                  & vbCrLf & vbTab & "       WHEN " & rpad("b.typ IN ( 2 /* 'closed' */, 4 /* 'changed' */)", " ") & " THEN a." & o _
                  & vbCrLf & vbTab & "  END  AS `" & o & "`"
    Next o
        
   res = vbCrLf & "-- # closed (tar) & changed (tar) & unchanged (src)" _
       & vbCrLf _
       & "INSERT INTO ${V_TAR}" _
       & vbCrLf _
       & "(" _
       & vbCrLf _
       & insertFields _
       & vbCrLf _
       & "   , version    " _
       & vbCrLf _
       & "   , start_dt   " _
       & vbCrLf _
       & "   , end_dt     " _
       & vbCrLf _
       & "   , load_stamp " _
       & vbCrLf _
       & ")" _
       & vbCrLf
       
    res = res _
       & "SELECT" _
       & vbCrLf _
       & selectKeys _
       & casewhen _
       & vbCrLf _
       & vbTab & ", a.version   AS version  " _
       & vbCrLf _
       & vbTab & ", a.start_dt  AS start_dt " _
       & vbCrLf _
       & vbTab & ", CASE WHEN " & rpad(whereKeys, " ") & VBA.Trim("                          ") & " THEN a.end_dt    -- his             " _
       & vbCrLf _
       & vbTab & "       WHEN " & rpad("b.typ = 3 /* 'unchanged' */                     ", " ") & " THEN ${MAX_DT}                    " _
       & vbCrLf _
       & vbTab & "       WHEN " & rpad("b.typ IN ( 2 /* 'closed' */ , 4 /* 'changed' */)", " ") & " THEN ${MY_DT}                     " _
       & vbCrLf _
       & vbTab & "       ELSE cast('0001-01-01' as date) -- never used                                                                " _
       & vbCrLf _
       & vbTab & "   END    As end_dt                                                                                    " _
       & vbCrLf
       
    res = res _
       & "  , a.load_stamp                                                                                        " _
       & vbCrLf _
       & " FROM       ${TAR_TABLE}   a          " _
       & vbCrLf _
       & " LEFT JOIN  ${V_CAT}       b          " _
       & vbCrLf _
       & aJoinBKeys _
       & vbCrLf _
       & "   -- AND b.`typ` IN ('closed', 'changed', 'unchanged')                                                 " _
       & vbCrLf _
       & "  AND a.start_dt < ${MY_DT} -- support reload operation                                               " _
       & vbCrLf _
       & "  AND a.end_dt  >= ${MY_DT} -- support reload operation                                               " _
       & vbCrLf _
       & " LEFT JOIN  ${SRC_TABLE}                                           c         " _
       & vbCrLf _
       & bJoinCKeys _
       & vbCrLf _
       & " WHERE a.start_dt < ${MY_DT}  -- contains his-data && support reload operation " _
       & vbCrLf _
       & ";"
       
       genUnion_His_Closed_Changed_Unchanged = res
       
End Function

'**********************************************************
'*
'* 将changed的记录封链后，再次插入到目标表（临时目标表）
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function genUnion_NewChanged() As String

    genUnion_NewChanged = vbCrLf & " -- changed(from src)" _
                        & vbCrLf _
                        & formatInsertIntoTar() _
                        & vbCrLf _
                        & formatSelectChanged()

End Function


'**********************************************************
'*
'* 将临时表的数据刷新到目标表，同时将目标表备份一下（测试阶段）
'*
'* 修改：无 （例：2016/01/01 by lujx XXX）
'*
'**********************************************************
Private Function getUnion_Target() As String

    s = vbCrLf & "-- Just for test - bakup"
    s = s & vbCrLf & "DROP   TABLE IF EXISTS       ${TAR_TABLE_BAK};"
    s = s & vbCrLf & "RENAME TABLE ${TAR_TABLE} TO ${TAR_TABLE_BAK};"  ' "TRUNCATE TABLE ${TAR_TABLE};"
    s = s & vbCrLf & "RENAME TABLE ${V_TAR}     TO ${TAR_TABLE};"
    
    getUnion_Target = s
    
End Function

' 临时分类表名称标准化
Private Function stdVCatTbl(ByVal SrcTbl As String, Optional ByVal Schema As String = "") As String

    Dim v As String
    
    v = "vt_" & SrcTbl & "_catlog"
    
    stdVCatTbl = Switch(Schema = Empty, v, Schema <> Empty, Schema & "." & v)
    
End Function

' 临时目标表名称标准化
Private Function stdVTarTbl(ByVal TarTbl As String, Optional ByVal Schema As String = "") As String

    Dim v As String
    
    v = "vt_" & TarTbl
    
    stdVTarTbl = Switch(Schema = Empty, v, Schema <> Empty, Schema & "." & v)
    
End Function

' 得到格式化的列 => 创建临时表时使用
Private Function formatColumn(FieldName As String, FieldType As String, Optional NullFlag As Boolean = True, Optional FieldComment As String = "", Optional DefualtValue As String = "")

    'Dim stdLen As Integer
    
    'stdLen = (WorksheetFunction.Floor(VBA.Len(FieldName) / INT_20, 1) + 1) * INT_20
    
    formatColumn = rpad(FieldName, " ", 30) _
                   & rpad(FieldType, " ", 30) _
                   & Switch(NullFlag, rpad(" ", " ", 30), Not NullFlag, rpad("NOT NULL ", " ", 30)) _
                   & Switch(DefualtValue = Empty, rpad(" ", " ", 30), DefualtValue <> Empty, rpad("DEFAULT " & DefualtValue, " ", 30)) _
                   & "COMMENT " _
                   & "'" & FieldComment & "'"

End Function

Public Function formatPrimaryKey(keys As Collection, Optional haveTyp As Boolean = True) As String
    
    Dim v As String
    
    v = "PRIMARY KEY ("
    For Each o In keys
       v = v & o & ","
    Next o
    
    formatPrimaryKey = Switch(haveTyp, v & "typ)", Not haveTyp, VBA.Left(v, VBA.Len(v) - 1) & ")")
    
End Function

Public Function formatKey(keys As Collection) As String
    
    Dim v As String
    
    v = "KEY idx("
    For Each o In keys
       v = v & o & ","
    Next o
    
    formatKey = VBA.Left(v, VBA.Len(v) - 1) & ")"
    
End Function

' 格式化连接条件 =>   on a.col1 = b.col1
'                    and a.col2 = b.col2
Private Function formatJoin(LeftAlias As String, RightAlias As String, cons As Collection)

        Dim joins As String
        
        For Each o In cons
            joins = Switch(joins = Empty, "  ON " & LeftAlias & "." & o & " = " & RightAlias & "." & o, joins <> Empty, joins & vbCrLf & " AND " & LeftAlias & "." & o & " = " & RightAlias & "." & o)
        Next
        
        formatJoin = joins
    
End Function

' 格式化插入目标表sql => INSERT INTO ${KTL_UDF_SCHEMA_DW}.vt_s_pl_crm_overdue_dk_h
'                        (
'                              ID
'                            , .....
'                            , start_dt
'                            , end_dt
'                            , load_stamp
'                        )
Private Function formatInsertIntoTar() As String

    Dim o, insertFields As String, res As String
        
    For Each o In tableMetaDataDict.keys
    
        If VBA.Left(o, 1) <> "#" Then
            
            If insertFields = Empty Then
                insertFields = vbTab & "  `" & o & "`"
            Else
                insertFields = insertFields & vbCrLf & vbTab & ", `" & o & "`"
            End If
            
        End If
    Next
        
   res = "INSERT INTO ${V_TAR} " _
       & vbCrLf _
       & "(" _
       & vbCrLf _
       & insertFields _
       & vbCrLf _
       & vbTab & ", version    " _
       & vbCrLf _
       & vbTab & ", start_dt   " _
       & vbCrLf _
       & vbTab & ", end_dt     " _
       & vbCrLf _
       & vbTab & ", load_stamp " _
       & vbCrLf _
       & ")" _

       formatInsertIntoTar = res
    
End Function

' 格式化插入目标表sql =>
'                        SELECT
'                              a.`id`            AS id
'                            , '${MY_DT}'        AS start_dt
'                            , '${MAX_DT}'       AS end_dt
'                            , CURRENT_TIMESTAMP
'                        FROM       ${KTL_UDF_SCHEMA_DW}.crm_overdue_dk                         a
'                        INNER JOIN ${KTL_UDF_SCHEMA_DW}.vt_s_pl_crm_overdue_dk_catlog          b
'                           ON a.id    = b.`k`
'                          AND b.`typ` = 4 /* 'changed' */
'                        ;
Private Function formatSelectChanged() As String

    Dim o, selectFields As String, res As String
        
    For Each o In tableMetaDataDict.keys
        If VBA.Left(o, 1) <> "#" Then
            
            If selectFields = Empty Then
                selectFields = vbTab & "  " & rpad(CStr("a." & o), " ", 30) & " as `" & o & "`"
            Else
                selectFields = selectFields & vbCrLf & vbTab & ", " & rpad(CStr(specialTransform("a." & o, True)), " ", 30) & " as `" & o & "`"
            End If
            
        End If
    Next
            
    res = "SELECT " _
          & vbCrLf _
          & selectFields _
          & vbCrLf _
          & vbTab & ", b.ver + 1         AS version   " _
          & vbCrLf _
          & vbTab & ", ${MY_DT}          AS start_dt  " _
          & vbCrLf _
          & vbTab & ", ${MAX_DT}         AS end_dt    " _
          & vbCrLf _
          & vbTab & ", CURRENT_TIMESTAMP              " _
          & vbCrLf _
          & "FROM       ${SRC_TABLE}                 a" _
          & vbCrLf _
          & "INNER JOIN ${V_CAT}                     b" _
          & vbCrLf _
          & formatJoin("a", "b", keySet()) _
          & vbCrLf _
          & "  AND b.`typ` = 4 /* 'changed' */       " _
          & vbCrLf _
          & " ;"

       formatSelectChanged = res
    
End Function

' 带coalesce功能的特殊处理： 将时间戳类型的列转换为int
Private Function specialTransform(ByVal vField As Variant, Optional ByVal coalesce As Boolean = False) As String

            ' mysql timestamp -> unix version
            Dim t
            t = VBA.Split(vField, ".")
            If CStr(tableMetaDataDict(t(UBound(t)))("FIELD_TYPE")) = "TIMESTAMP" Then
            
                    specialTransform = "UNIX_TIMESTAMP(" & vField & ")"
                    
            ' coalesce process
            ElseIf coalesce _
                    And (VBA.InStr(CStr(tableMetaDataDict(t(UBound(t)))("FIELD_TYPE")), "CHAR") > 0 _
                          Or VBA.InStr(CStr(tableMetaDataDict(t(UBound(t)))("FIELD_TYPE")), "TEXT") > 0 _
                        ) Then  ' CHAR TYPE
                        
                    specialTransform = "COALESCE( " & vField & " , '')"
            Else
                    specialTransform = CStr(vField)
            End If
            
End Function

' 对字符串类型加入coalesce操作
Private Function coalesce(ByVal vField As String) As String

    Dim t
    t = VBA.Split(vField, ".")
    
    If _
                             VBA.InStr(CStr(tableMetaDataDict(t(UBound(t)))("FIELD_TYPE")), "CHAR") > 0 _
                          Or VBA.InStr(CStr(tableMetaDataDict(t(UBound(t)))("FIELD_TYPE")), "TEXT") > 0 _
    Then  ' CHAR TYPE
                        
        coalesce = "COALESCE( " & vField & " , '')"
    
    Else
        coalesce = vField

    End If
    
End Function

'  持久化到文件sql
Private Sub persistence(textCache As Collection, FileName As String)

        Dim objWriter As New FileOperation
             
        ' default
        If FileName = Empty Then
            FileName = "d:\tmp\" & tableMetaDataDict(TAR_TABLE_NAME) & ".sql"
        End If
        
         ' write to file
         objWriter.OpenFile FileName, "W"
        
        For Each o_line In textCache
            objWriter.WriteLine (o_line)
        Next
    
        objWriter.CloseFile
        
End Sub

Private Function replaceVar(ByRef s As String) As String

    Dim result As String
    
    result = s
    
    For Each o In varDict.keys
        result = VBA.Replace(result, o, varDict(o))
    Next o
    
    replaceVar = result

End Function

