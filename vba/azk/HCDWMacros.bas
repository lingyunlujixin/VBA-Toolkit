Attribute VB_Name = "HCDWMacros"
'############################################################################################################
'#
'#
'# 2015/10/22 LUJX
'#
'# ��v2�ĺ���������Job���������ϵ�Ĺ���
'#
'#
'#
'#
'#
'############################################################################################################

'***************************************
'*
'* ����macro������azk2UATv3ʹ�õ�job�ļ�
'*
'* ��1��. ���汾��ȥ��ĸ��������������
'*
'* 2015/11/17 lujx
'*
'* 2015/12/14 lujx �����˶�embedded-flow job�Ĺ���
'*
'**************************************
Sub azkpan2uatv3()

    Dim objTXTFO As New FileOperation
    
    Dim Layer  As New HCDWLayerOperationV3
    
    Layer.init Worksheets(21)
    
    'Debug.Print (layer.sprintfA)
    
    Dim azklay As Object, str As String
    
    Dim WDIR As String
    
    ' �õ�ǰ������Ϊ·����d:\tmp\20150901\
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

    ' ��ʾ C:\ Ŀ¼�µ����ơ�
    mypath = "d:\temp"    ' ָ��·����
    myname = DIR(mypath & "\*.xls")       ' ��Ѱ��һ�
    
    ' Debug.Print MyName
    
    Do While myname <> ""    ' ��ʼѭ����
        ' ������ǰ��Ŀ¼���ϲ�Ŀ¼��
        If myname <> "." And myname <> ".." Then
        
            ' ʹ��λ�Ƚ���ȷ�� MyName ����һĿ¼��
           ' If (GetAttr(MyPath & MyName) And vbNormal) = vbDirectory Then
                
                ' Debug.Print myname    ' �������һ��Ŀ¼������������ʾ������
           
                Set curwkbk = Workbooks.Open(mypath & "\" & myname)
                
                Set curwkst = curwkbk.Sheets(1)
                
                Debug.Print (curwkst.Cells(1, 1).Value)
                
                curwkbk.Close savechanges:=False
                
           ' End If
        End If
        
        myname = DIR    ' ������һ��Ŀ¼��
    Loop

End Sub
