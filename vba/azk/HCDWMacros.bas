Attribute VB_Name = "HCDWMacros"
'############################################################################################################
'#
'#
'# 2015/10/22 LUJX
'#
'# 带v2的函数增加了Job层间依赖关系的管理
'#
'#
'#
'#
'#
'############################################################################################################

'***************************************
'*
'* 运行macro，产生azk2UATv3使用的job文件
'*
'* （1）. 本版本消去层的概念，建立整体依赖
'*
'* 2015/11/17 lujx
'*
'* 2015/12/14 lujx 增加了对embedded-flow job的管理
'*
'**************************************
Sub azkpan2uatv3()

    Dim objTXTFO As New FileOperation
    
    Dim Layer  As New HCDWLayerOperationV3
    
    Layer.init Worksheets(21)
    
    'Debug.Print (layer.sprintfA)
    
    Dim azklay As Object, str As String
    
    Dim WDIR As String
    
    ' 用当前日期作为路径：d:\tmp\20150901\
    Let WDIR = BASE_DIR & Format(Date, "yyyymmdd") & "_azk2uatv3_\"
    
    CreateIfNotExists (WDIR)
    
    For Each obj In Layer.hcdw.Keys
        
        objTXTFO.OpenFile WDIR & obj & ".job", "W"
    
        str = makeJOB(CStr(obj), "command", makeCMD2(REPOSITORY, REP_USER, REP_PASSWORD, CStr(obj), KTL_PAN2), "", Layer.hcdw.Item(obj))
              
        objTXTFO.WriteLine (str)
    
        objTXTFO.CloseFile
    
    Next
    
    package WDIR, RAR, ZIP
        
    ''' *** FILE PRODUCE OK !!! ***
    
    MsgBox "All job files are saved in [ " & WDIR & " ] ! "
    
End Sub

Sub testRar()
    package "D:\tmp\20150909_azk2proc\", "C:\Program Files (x86)\WinRAR", "zip"
End Sub

Sub testOpenxlsx()

    Dim curwkbk As Workbook
    Dim curwkst As Worksheet

    ' 显示 C:\ 目录下的名称。
    mypath = "d:\temp"    ' 指定路径。
    myname = DIR(mypath & "\*.xls")       ' 找寻第一项。
    
    ' Debug.Print MyName
    
    Do While myname <> ""    ' 开始循环。
        ' 跳过当前的目录及上层目录。
        If myname <> "." And myname <> ".." Then
        
            ' 使用位比较来确定 MyName 代表一目录。
           ' If (GetAttr(MyPath & MyName) And vbNormal) = vbDirectory Then
                
                ' Debug.Print myname    ' 如果它是一个目录，将其名称显示出来。
           
                Set curwkbk = Workbooks.Open(mypath & "\" & myname)
                
                Set curwkst = curwkbk.Sheets(1)
                
                Debug.Print (curwkst.Cells(1, 1).Value)
                
                curwkbk.Close savechanges:=False
                
           ' End If
        End If
        
        myname = DIR    ' 查找下一个目录。
    Loop

End Sub
