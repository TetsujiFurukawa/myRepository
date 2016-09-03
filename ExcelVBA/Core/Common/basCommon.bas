Attribute VB_Name = "basCommon"
Option Explicit
'********************************************************************************
'
' ���ʏ���
'
'********************************************************************************

'******************************************************************
' �K�{�`�F�b�N
'
' ���� strTarget�F�`�F�b�N�Ώۂ̒l
'
' �߂�l�FTrue�G���[����
'         False�G���[�Ȃ�
'******************************************************************
Public Function IsBlank(strTarget As String, strMes As String) As Boolean
  IsBlank = False
  ' �����̓`�F�b�N
  If Trim(strTarget) = "" Then
    MsgBox strMes & "����͂��Ă��������B", vbCritical, "�G���[���b�Z�[�W"
    IsBlank = True
    Exit Function
  End If

End Function

'******************************************************************
' �t�@�C���̃`�F�b�N
'
' ���� path�F�t�@�C����
'
' �߂�l�FTrue�G���[����
'         False�G���[�Ȃ�
'******************************************************************
Public Function FileExist(path As String, strMes As String) As Boolean
  FileExist = False
  ' ���݃`�F�b�N
  With CreateObject("Scripting.FileSystemObject")
    If Not .FileExists(path) Then
      MsgBox strMes & "�����݂��܂���B", vbCritical, "�G���[���b�Z�[�W"
      Exit Function
    End If
  End With
  
  FileExist = True
End Function

'******************************************************************
' �t�H���_�̃`�F�b�N
'
' ���� path�F�t�@�C����
'
' �߂�l�FTrue�G���[����
'         False�G���[�Ȃ�
'******************************************************************
Public Function FolderExist(path As String, strMes As String) As Boolean
  FolderExist = False
  ' ���݃`�F�b�N
  With CreateObject("Scripting.FileSystemObject")
    If Not .FolderExists(path) Then
      MsgBox strMes & "�t�H���_�����݂��܂���B", vbCritical, "�G���[���b�Z�[�W"
      Exit Function
    End If
  End With
  
  FolderExist = True
End Function


'******************************************************************
' �t�H���_�쐬
'
' ���� path�F�t�@�C����
'
' �߂�l�FTrue�G���[����
'         False�G���[�Ȃ�
'******************************************************************
Public Sub CreateDir(path As String)
  With CreateObject("Scripting.FileSystemObject")
    If Not .FolderExists(path) Then
      .CreateFolder (path)
    End If
  End With

End Sub

'************************************************************************
'
' Excel�V�[�g����ŏI�s�̈ʒu���擾
'
' ���� targerSheet�F�Ώ���Excel�V�[�g
'      lngStartRowNo�F���׍s�̊J�n�ʒu
'      lngCountColumnNo1�F�ŏI�s���J�E���g�����ʒu1
'      lngCountColumnNo2�F�ŏI�s���J�E���g�����ʒu2
'
' �߂�l �ŏI�s�̈ʒu�A 1�����f�[�^���Ȃ��ꍇ�Ȃǂ͊J�n�s�ʒu��Ԃ�
'************************************************************************
Public Function GetMaxRowNo(targerSheet As Excel.Worksheet, lngStartRowNo As Long, lngCountColumnNo As Long) As Long

  GetMaxRowNo = targerSheet.Cells(Rows.Count, lngCountColumnNo).End(xlUp).Row
    
  ' 1�����f�[�^���Ȃ��ꍇ�Ȃǂ͊J�n�s�ʒu��Ԃ�
  If GetMaxRowNo < lngStartRowNo Then
    GetMaxRowNo = lngStartRowNo
  End If
  
End Function

'************************************************************************
'
' Excel�V�[�g���̎w�肵���͈͂Ń\�[�g
'
' ���� targerSheet�F�Ώ���Excel�V�[�g
'      lngRowStartNo�F���׍s�̊J�n�s�ʒu
'      lngRowEndNo�F���׍s�̏I���s�ʒu
'      lngColStartNo�F���׍s�̊J�n��ʒu
'      lngColEndNo�F���׍s�̏I����ʒu
'      lngSortColNo�F�\�[�g���鍀�ڂ̗�ʒu
'
'************************************************************************
Public Sub SortSpecifiedRange(targerSheet As Excel.Worksheet, lngRowStartNo As Long, lngRowEndNo As Long, lngColStartNo As Long, lngColEndNo As Long, lngSortColNo As Long)

  targerSheet.Activate
  targerSheet.Range(Cells(lngRowStartNo, lngColStartNo), Cells(lngRowEndNo, lngColEndNo)) _
           .Sort Key1:=targerSheet.Cells(lngRowStartNo, lngSortColNo), order1:=xlAscending, DataOption1:=xlSortTextAsNumbers

End Sub

'************************************************************************
'
' Excel�V�[�g���̎w�肵���͈͂�z��Ŏ擾
'
' ���� targerSheet�F�Ώ���Excel�V�[�g
'      lngRowStartNo�F���׍s�̊J�n�s�ʒu
'      lngRowEndNo�F���׍s�̏I���s�ʒu
'      lngColStartNo�F���׍s�̊J�n��ʒu
'      lngColEndNo�F���׍s�̏I����ʒu
'
'************************************************************************
Public Function GetSpecifiedRange(targerSheet As Excel.Worksheet, lngRowStartNo As Long, lngRowEndNo As Long, lngColStartNo As Long, lngColEndNo As Long) As Variant
  Dim myArray As Variant
  myArray = targerSheet.Range(Cells(lngRowStartNo, lngColStartNo), Cells(lngRowEndNo, lngColEndNo))

  GetSpecifiedRange = myArray
  
  Set myArray = Nothing

End Function








'******************************************************************
' �����R�[�h���w�肵�ăt�@�C����Ǎ�
'
' ���� strFileName�F�t�@�C����
'      strCharSet�F�����R�[�h(EUC-JP�AShift_JIS)
'
' �ߒl ������̔z��
'******************************************************************
Public Function FileRead(strFileName As String, strCharSet As String, strLineSeparator) As String()
 
  Dim adoSt As New ADODB.Stream
  Dim lngCount As Long
  Dim res() As String, strBuf As String
  
  lngCount = 0
  
  With adoSt
    .Type = adTypeText
    .Charset = strCharSet
    .LineSeparator = strLineSeparator
    .Open
    .LoadFromFile (strFileName)
    
  Do While Not (.EOS)
    ReDim Preserve res(lngCount)
    strBuf = .ReadText(adReadLine)
    res(lngCount) = strBuf
    lngCount = lngCount + 1
  Loop
  
  .Close
  End With
  
  FileRead = res
  
  Set adoSt = Nothing
 
End Function

'******************************************************************
' �����R�[�h���w�肵�ăt�@�C����������
'
' ���� strFileName�F�t�@�C����
'      strOutputData()�F������̔z��
'      strCharSet�F�����R�[�h(EUC-JP�AShift_JIS)
'      strLineSeparator�F���s����
'
' �ߒl ������̔z��
'******************************************************************
Public Sub FileWrite(strFileName As String, strOutputData() As String, strCharSet As String, strLineSeparator)
 
  Dim adoSt As New ADODB.Stream
  Dim lngRowCount As Long
  
  With adoSt
    .Type = adTypeText
    .Charset = strCharSet
    .LineSeparator = strLineSeparator
    .Open
    
    For lngRowCount = 0 To UBound(strOutputData)
  
      .WriteText strOutputData(lngRowCount), adWriteLine
      
    Next lngRowCount
  
    .SaveToFile strFileName, adSaveCreateNotExist
    
    .Close
    
  End With
  
  Set adoSt = Nothing
 
End Sub

'************************************************************************
'
' Variant�^��String�^�ɕύX����
' (���s���Ɍ^����v���Ȃ��ƌ�����ꍇ�Ɏg���j
'
' ���� varValue�F�Ώۃf�[�^
'
' �߂�l ��L
'************************************************************************
Public Function ToString(varValue As Variant) As String

  ToString = varValue
  
End Function


'
'Public Sub CopySheet(strBookNameFrom As String, strSheetNameFrom As String, strBookNameTo As String, strSheetNameTo As String)
'
'  Application.DisplayAlerts = False
'  Workbooks(strBookNameTo).Worksheets(strBookNameTo).Delete
'  Application.DisplayAlerts = True
'  Workbooks(strBookNameFrom).Worksheets(strSheetNameFrom).Copy After:=Workbooks(strBookNameTo).Worksheets(strBookNameTo)
'
'End Sub



