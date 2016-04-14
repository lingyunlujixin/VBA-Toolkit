VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HCDWLayerOperationV3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**********************************************************************************************************
'*
'* 暂时未使用，未来打算替换当前的LayerOperation
'*
'* 将指定的sheet中所有trans名加载
'*
'* 到容器中，并根据命名规则进行分
'*
'* 类
'*
'* 对【层间依赖】的情况，进行了系统化处理：
'*
'*   1） 第一列为job名，第二列为依赖的job
'*
'*   2） 当1个job依赖多个job时，写多行表示
'*
'*   3） 从excel读取数据后，结果如下(json格式)：
'*
'*       # format1
'*       dictMlayer -- {
'*                          "jobs":[
'*
'*                                   { "job"  : "job1", "dep" : [ {"dep":"depjob1" } , { "dep" : "depjob2" }                          ] } ,
'*                                   { "job"  : "job2", "dep" : [ {"dep":"depjob1" } , { "dep" : "depjob2" } , { "dep" : "depjob3" }  ] } ,
'*                                   { "job"  : "job3", "dep: : [ {"dep":"depjob1" }                                                  ] }
'*
'*                                 ]
'*
'*                     }  # 外层是一个Dictionary 【key = "jobs"，value = Collection [{key = jobName, value = dep_job_list}] 】
'*
'*       # format2
'*       dictMlayer -- {                                                                                         # ==>> {
'*                           "job1" : [ {"dep":"depjob1" } , { "dep" : "depjob2" }                          ]  , # ==>>    "job1" : [ "dep1",               ] ,
'*                           "job2" : [ {"dep":"depjob1" } , { "dep" : "depjob2" } , { "dep" : "depjob3" }  ]  , # ==>>    "job2" : [ "dep1", "dep2"        ] ,
'*                           "job3" : [ {"dep":"depjob1" }                                                  ]    # ==>>    "job3" : [ "dep1", "dep2", "dep3"]
'*                                                                                                               # ==>> }
'*                     }  # 外层是一个Dictionary 【key = job名，value = 包含被依赖job名的Collection】
'*
'*       # format3
'*       dictMlayer -- { "jobs":
'*                        [
'*                           {"job" : "job1" , "desc" : "job1 description" , "dep" : ["dep1", "# dep1-desc"]                                                 },
'*                           {"job" : "job2" , "desc" : "job2 description" , "dep" : ["dep1" ,"# dep1-desc", "dep2",  "# dep2-desc"]                         },
'*                           {"job" : "job3" , "desc" : "job3 description" , "dep" : ["dep1" ,"# dep1-desc" ,"dep2" , "# dep2-desc", "dep3", "# dep3-desc"]  },
'*                           ……
'*                        ]
'*
'*                     }  #
'*
'**********************************************************************************************************

Public hcdw As Object


' ***** 初始化 *****
Public Sub init(ByRef ws As Worksheet)

    Dim cur_sheet As Worksheet
    
    Set cur_sheet = ws
    
    ' 计数变量
    Dim i As Integer, j As Integer, k As Integer
    
    Dim obj As Variant, dep As Variant, desc As Variant ' 临时变量，记录每行读取的trans和其依赖trans
    
    Dim c1 As Collection               ' 作为词典的items部分使用的临时变量
    
    Set hcdw = CreateObject("Scripting.Dictionary")
    
    i = 2
    
    j = i - 2
    
    Do While cur_sheet.Cells(i, 1) <> ""
        
            obj = cur_sheet.Cells(i, 1).Value  ' trans名字
            dep = cur_sheet.Cells(i, 2).Value  ' 依赖的trans名
            desc = cur_sheet.Cells(i, 3).Value ' 中文名称（注释）
        
        
            If hcdw.Exists(obj) Then
                
                    Set c1 = hcdw.Item(obj)
                    
                    c1.Add (dep)
                    c1.Add ("# " & desc)

                Else
                    Set c1 = New Collection
                    
                    c1.Add (dep)
                    c1.Add ("# " & desc)
                    
                    hcdw.Add obj, c1
                    
            End If

             
            ' 指向下一张表
            i = i + 1
            j = j + 1
                
        Loop
        
        ' MsgBox m0.Count
        
    
End Sub

' 打印各层级Job名称及对应依赖关系, json显示《简易》
Public Function sprintf()

    Dim o, oo, sval, fval
    
    For Each o In hcdw.Keys

        sval = "["
        For Each oo In hcdw.Item(o)
            sval = sval & quota(oo) & ","
        Next
        
        sval = trimComma(CStr(sval), tcRIGHT) & "]"
    
        fval = fval & Chr(9) & quota(o) & " : " & sval & Chr(13)
    Next
    
    sprintf = "{ " & Chr(13) & fval & " }"

End Function

