Attribute VB_Name = "Macros"
'***************************************************************************
'*
'*  ���ܣ����svn��src�ļ������е���{Date}���������ktr�ļ���trans����
'*
'*        ����������洢��log�ļ��С�
'*
'*        * ע ��ʹ��ǰ���ȸ���svn��srcĿ¼������״̬
'*
'*  ���ߣ�¬����
'*
'*  ���ڣ�2015/09/02
'*
'***************************************************************************
Sub MACRO_CHECK_MOD()
Attribute MACRO_CHECK_MOD.VB_ProcData.VB_Invoke_Func = " \n14"
    '
    ' ��1 ��
    '
    
    '
    Dim dt As String, dt2 As String, path As String
    
    Dim cmd1, cmd2, res1, res2
    
    dt = Format(Date, "yyyy/mm/dd")
    
    dt2 = Format(Date, "yyyymmdd")
    
    path = "D:\�����ļ���\¬\SVN_150803\0600 ETL\src\hcdw\"
    
    res1 = path & "..\tmp\ModResult" & dt2 & ".log" ' ����D:\�����ļ���\¬\SVN_150803\0600 ETL\src\tmp\result20150909.log
    res2 = path & "..\tmp\DirResult" & dt2 & ".log" ' ����D:\�����ļ���\¬\SVN_150803\0600 ETL\src\tmp\result20150909.log
    
    ' ���� cmd /c cd /d *path* && dir /s *.ktr | find "2019/09/09" > *path*
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
    
    Logger.FileName = "D:\�����ļ���\¬\debug_" & Format(Now, "yyyyMMdd") & ".log"
    
    ''''' ��mypath�޸�ΪҪչʾ��
    'mypath = "D:\�����ļ���\¬\SVN_150803\0600 ETL\Ͷ���ƻ�\20150906-20150911"
    '
    'mypath = "D:\�����ļ���\¬\SVN_150803\0600 ETL\Ͷ���ƻ�\20150914-20150920"
    
    'mypath = "D:\�����ļ���\¬\SVN_150803\0600 ETL\Ͷ���ƻ�\20150906-20150911"
    
    'mypath = "D:\�����ļ���\¬\SVN\0600 ETL\Ͷ���ƻ�\20151026-20151101"
    
    mypath = "D:\�����ļ���\��ӭ��\-100-HC-�����ݷ���ƽ̨\0600 ETL\Ͷ���ƻ�\20160121"
    
    myname = DIR(mypath & "\*.xls*")
    
    Do While myname <> ""
    
        i = 2
        
        If myname <> "." And myname <> ".." And InStr(myname, "ģ��") <= 0 And InStr(myname, "����") <> 1 Then
        
            Set curwkbk = Workbooks.Open(mypath & "\" & myname)
            
            Set curwkst = curwkbk.Sheets("Ͷ���嵥")
            
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
'* �����ļ��������е��ļ����ݹ飩
'*
'*
'*
'****************************************
Sub dirRecursion()

    'tRecursion "D:\�����ļ���\¬\SVN_150803\0600 ETL\Ͷ���ƻ�\"
    'tRecursion "D:\�����ļ���\¬\SVN_150803\0600 ETL\src\hcdw\edw\BasLayer"
    'tRecursion "D:\PID_PROC\�����"
    tRecursion "D:\soft\pdi-ce-5.3.0.0-213\data-integration\samples\import-lists-bak"
End Sub


'*************************************
'*
'*
'* ��Ҫ
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
'* ����һ��·������ӡ����·���µ��ļ�
'*
'* ��������ļ��У���ݹ������ļ���
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
