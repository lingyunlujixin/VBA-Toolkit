Attribute VB_Name = "Macros"
'***************************************************************************
'*
'*  功能：检查svn上src文件下所有当天{Date}发生变更的ktr文件（trans），
'*
'*        并将检查结果存储在log文件中。
'*
'*        * 注 ：使用前请先更新svn的src目录到最新状态
'*
'*  作者：卢吉欣
'*
'*  日期：2015/09/02
'*
'***************************************************************************
Sub MACRO_CHECK_MOD()
Attribute MACRO_CHECK_MOD.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' 宏1 宏
    '
    
    '
    Dim dt As String, dt2 As String, path As String
    
    Dim cmd1, cmd2, res1, res2
    
    dt = Format(Date, "yyyy/mm/dd")
    
    dt2 = Format(Date, "yyyymmdd")
    
    path = "D:\个人文件夹\卢\SVN_150803\0600 ETL\src\hcdw\"
    
    res1 = path & "..\tmp\ModResult" & dt2 & ".log" ' 例：D:\个人文件夹\卢\SVN_150803\0600 ETL\src\tmp\result20150909.log
    res2 = path & "..\tmp\DirResult" & dt2 & ".log" ' 例：D:\个人文件夹\卢\SVN_150803\0600 ETL\src\tmp\result20150909.log
    
    ' 例： cmd /c cd /d *path* && dir /s *.ktr | find "2019/09/09" > *path*
    cmd1 = "cmd /c cd /d """ _
          & path _
          & """ && dir /s *.ktr | find """ _
          & dt _
          & """ > """ _
          & res1 _
          & """"
          
       'cmd2 = "cmd /c cd /d """ _
       '    left(path, len(path) - 5) _
       '    """ &&" _
       '    mGetCyCmd()
          
    cmd2 = "cmd /c cd /d """ _
          & path _
          & "..\tmp"" && check_dir.bat" _
          & " > """ _
          & res2 _
          & """"
          
          Debug.Print cmd2
          
          Shell cmd1
          ' Shell cmd2
    
    Debug.Print "Check result is saved in " & res1
    
    ' cmd /c @for /f "tokens=4 delims= " %i in (res) do @echo %i

End Sub

Public Sub testGetFullRepPath()
    ' Debug.Print getFullRepPath("TRANS_4_S_HYR_USER_H")
    
    Debug.Print mGetCyCmd("TRANS_4_S_HYR_USER_H")
End Sub

Private Function mGetCyCmd(trans As String) As String
    Dim s, l
    
    l = getFullRepPath(trans, "\")
    
    ' COPY /Y .\hcdw\ods\StdLayer\ktr\daily\TRANS_4_S_HYR_USER_H.ktr tmp
    s = "COPY /Y "
    s = s & "." & l & ".ktr"
    s = s & " tmp"
    
    mGetCyCmd = s
End Function


Public Sub tMergeExcel()
    
    Dim curwkbk As Workbook
    Dim curwkst As Worksheet
    Dim merge   As New Collection
    Dim arrval(11) As String
    Dim i As Integer
    Dim o
    
    Logger.LogEnabled = True
    
    Logger.LogCallback = Array("logFile", "ImmediateLog")
    
    Logger.FileName = "D:\个人文件夹\卢\debug_" & Format(Now, "yyyyMMdd") & ".log"
    
    ''''' 把mypath修改为要展示的
    'mypath = "D:\个人文件夹\卢\SVN_150803\0600 ETL\投产计划\20150906-20150911"
    '
    'mypath = "D:\个人文件夹\卢\SVN_150803\0600 ETL\投产计划\20150914-20150920"
    
    'mypath = "D:\个人文件夹\卢\SVN_150803\0600 ETL\投产计划\20150906-20150911"
    
    'mypath = "D:\个人文件夹\卢\SVN\0600 ETL\投产计划\20151026-20151101"
    
    mypath = "D:\个人文件夹\马迎新\-100-HC-大数据分析平台\0600 ETL\投产计划\20160121"
    
    myname = DIR(mypath & "\*.xls*")
    
    Do While myname <> ""
    
        i = 2
        
        If myname <> "." And myname <> ".." And InStr(myname, "模板") <= 0 And InStr(myname, "补数") <> 1 Then
        
            Set curwkbk = Workbooks.Open(mypath & "\" & myname)
            
            Set curwkst = curwkbk.Sheets("投产清单")
            
                Do While curwkst.Cells(i, 1) <> ""
                
                    arrval(0) = curwkbk.Name
                    arrval(2) = curwkst.Cells(i, 1)
                    arrval(4) = curwkst.Cells(i, 2)
                    arrval(6) = curwkst.Cells(i, 3)
                    arrval(8) = curwkst.Cells(i, 4)
                    arrval(10) = curwkst.Cells(i, 5)
                    
                    arrval(1) = 30
                    arrval(3) = 5
                    arrval(5) = 50
                    arrval(7) = 10
                    arrval(9) = 10
                    arrval(11) = 10
                    
                    merge.Add (arrval)
                    
                    i = i + 1
                    
                Loop
            
            curwkbk.Close savechanges:=False
            
        End If
        
        myname = DIR
    
    Loop
    
    For Each o In merge
    
        ' Debug.Print o(2) & " , " & o(3)
        
        ' printf (o)
        
        printf o
    Next o
    
    Logger.flush
    
End Sub

'****************************************
'*
'*
'* 遍历文件夹下所有的文件（递归）
'*
'*
'*
'****************************************
Sub dirRecursion()

    'tRecursion "D:\个人文件夹\卢\SVN_150803\0600 ETL\投产计划\"
    'tRecursion "D:\个人文件夹\卢\SVN_150803\0600 ETL\src\hcdw\edw\BasLayer"
    'tRecursion "D:\PID_PROC\已完成"
    tRecursion "D:\soft\pdi-ce-5.3.0.0-213\data-integration\samples\import-lists-bak"
End Sub


'*************************************
'*
'*
'* 不要
'*
'*
'*************************************
Private Sub tRecursion2(sdir As String)

    mypath = sdir
    myname = DIR(mypath, vbDirectory)
    
    ' Do While myname <> ""
        
        ' If myname <> "." And myname <> ".." Then
        
            ' Set curwkbk = Workbooks.Open(mypath & "\" & myname)
            
            If (GetAttr(mypath & "" & myname) And vbDirectory) Then
            
                tRecursion (mypath & myname & "\")
                
             Else
                 Debug.Print (mypath & "" & myname)
            
            End If
            
        ' End If
        

                       ' myname = DIR
                        
    ' Loop

End Sub


'********************************************
'*
'*
'* 传递一个路径，打印所有路径下的文件
'*
'* 如果有子文件夹，则递归所有文件夹
'*
'*******************************************
Private Sub tRecursion(s As String)

        Dim a As Collection

        Set a = getFileList(s)
        
        For Each o In a
        
            If (GetAttr(o) And vbDirectory) = vbDirectory Then
            
                tRecursion (o)

            Else
            
                Debug.Print o
                
            End If

        Next o
        
End Sub

Public Sub testPrintf()

    Dim arr As Variant
    
    arr = Array("1st col", 10, "sec col", 10, "3rd col", 7)
    
    'a(0) = "first word"
    'a(1) = 18
    'a(2) = "sec word"
    'a(3) = 19
    
   printf (arr)
   
End Sub
