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
'** 用于文本文件操作的类
'**
'**
'**
'**
'**
'*************************************************

Private objTS As TextStream    ' 定义TextStream对象

Private objAS As ADODB.Stream  ' 定义ADODB.STREAM对象

Public Function OpenFile(strFileName As String, strMode As String)

    Dim objFSO As FileSystemObject ' 定义文件对象
    
    Set objFSO = New FileSystemObject
    
    Set objTS = Nothing

    If strMode = "R" Then
        ' 读取方式打开文件
        Set objTS = objFSO.OpenTextFile(strFileName, ForReading, True)
    End If
    If strMode = "W" Then
        ' 写入方式打开文件
        Set objTS = objFSO.OpenTextFile(strFileName, ForWriting, True)
    End If
    If strMode = "APPEND" Then
        ' 写入方式打开文件
        Set objTS = objFSO.OpenTextFile(strFileName, ForAppending, True)
    End If
    
        
End Function

Public Function CloseFile()

    ' 关闭文件
    
    If Not objTS Is Nothing Then
        objTS.Close
    End If
    
    ' utf8的场合
    ' Set objAS = Nothing
    
End Function

Public Function GetLine() As String

    ' 按行读取文件数据
    GetLine = objTS.ReadLine

End Function

Public Function GetAll() As String

    ' 读取所有文件数据
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

    ' 判断是否已到文件末尾
    AtEndOfFile = objTS.AtEndOfStream()
    
End Property

Public Function WriteLine(StrData As String)

    ' 向文件写入一条数据
    objTS.WriteLine (StrData)
    
End Function

Public Function SkipLines(intLines As Integer)
    Dim i As Integer
    ' 跳到指定数据行，如指定行超过总行数，则指定到末尾行
    ' 一般结合文件末尾判断函数以及读取函数使用
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
