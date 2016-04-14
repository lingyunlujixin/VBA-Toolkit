'****************************************************************************
'*
'* Create By lujx 2015/08
'*
'* Common Utilities
'*
'* 更新记录：
'*
'****************************************************************************

Attribute VB_Name = "CommonUtil"
'**************************************************
'**
'** 返回 SH, SP, SF, B, M, A之一
'**
'**************************************************
Public Function getLayer(trans As String) As String

    'e.g. azkpan.sh -rep="MyKettleRepository" -trans=TRANS_4_S_PL_CRM_INTOPIECES_DK_H -dir=/hcdw/ods/StdLayer/ktr/daily -user=lujx -pass=lujx
    Dim tmp As String
    ' Dim ssys As String
    tmp = Switch(Left(trans, 1) = "I", Mid(trans, 15, 3), Left(trans, 1) = "T", Mid(trans, 9, 3))
    
    getLayer = Switch(Left(tmp, 1) = "S", Left(tmp, 1) & Right(tmp, 1), Left(tmp, 1) <> "S", Left(tmp, 1))

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
    
    DIR = Switch(layer = "S", "/hcdw/ods/StdLayer/ktr/daily", _
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

'**************************************************
'**
'** 生成JOB文件对应的内容
'**
'**************************************************
Public Function makeJOB(cmt As String, _
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
                  & Chr(13) _
                  & Switch(dep <> "", "dependencies=") _
                  & Switch(dep <> "", dep)
              
End Function

Public Sub CreateIfNotExists(strDirName As String)

    If DIR(strDirName, vbDirectory) = Empty Then ' 16
            ''''''' to do
        
            MKDIR strDirName
            
        End If

End Sub


Public Function getFullRepPath(trans As String) As String
    
    Dim layer As String
    
    layer = Left(getLayer(trans), 1)
    
    getFullRepPath Switch(layer = "S", "/hcdw/ods/StdLayer/ktr/daily/", _
                          layer = "B", "/hcdw/edw/BasLayer/ktr/daily/", _
                          layer = "M", "/hcdw/edw/MidLayer/ktr/daily/", _
                          layer = "A", "/hcdw/edw/AppLayer/ktr/daily/") _
                          & trans
End Function

