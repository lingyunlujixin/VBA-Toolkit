VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************
'**
'** �����ı��ļ���������
'**
'**
'**
'**
'**
'*************************************************

Private objTS As TextStream    ' ����TextStream����

Private objAS As ADODB.Stream  ' ����ADODB.STREAM����

Public Function OpenFile(strFileName As String, strMode As String)

    Dim objFSO As FileSystemObject ' �����ļ�����
    
    Set objFSO = New FileSystemObject
    
    Set objTS = Nothing

    If strMode = "R" Then
        ' ��ȡ��ʽ���ļ�
        Set objTS = objFSO.OpenTextFile(strFileName, ForReading, True)
    End If
    If strMode = "W" Then
        ' д�뷽ʽ���ļ�
        Set objTS = objFSO.OpenTextFile(strFileName, ForWriting, True)
    End If
    If strMode = "APPEND" Then
        ' д�뷽ʽ���ļ�
        Set objTS = objFSO.OpenTextFile(strFileName, ForAppending, True)
    End If
    
        
End Function

Public Function CloseFile()

    ' �ر��ļ�
    
    If Not objTS Is Nothing Then
        objTS.Close
    End If
    
    ' utf8�ĳ���
    ' Set objAS = Nothing
    
End Function

Public Function GetLine() As String

    ' ���ж�ȡ�ļ�����
    GetLine = objTS.ReadLine

End Function

Public Function GetAll() As String

    ' ��ȡ�����ļ�����
    GetAll = tranUTF8(objTS.readAll)

End Function

Public Function OpenReadUTF8(ByVal pfile As String) As String
          
          objAS.Open
          
          objAS.LineSeparator = adCRLF
          
          objAS.Charset = "utf-8"
          
          objAS.Type = adTypeText
          
          objAS.LoadFromFile pfile
          
          objAS.Position = 0

          OpenReadUTF8 = objAS.ReadText()
          
          objAS.Close
          
  End Function

Public Property Get AtEndOfFile() As Boolean

    ' �ж��Ƿ��ѵ��ļ�ĩβ
    AtEndOfFile = objTS.AtEndOfStream()
    
End Property

Public Function WriteLine(StrData As String)

    ' ���ļ�д��һ������
    objTS.WriteLine (StrData)
    
End Function

Public Function SkipLines(intLines As Integer)
    Dim i As Integer
    ' ����ָ�������У���ָ���г�������������ָ����ĩβ��
    ' һ�����ļ�ĩβ�жϺ����Լ���ȡ����ʹ��
    For i = 1 To intLines
        If objTS.AtEndOfStream Then
            Exit For
        End If
        objTS.SkipLine
    Next i
End Function

Private Sub Class_Initialize()

    Set objAS = New ADODB.Stream

End Sub