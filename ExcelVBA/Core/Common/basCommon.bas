Attribute VB_Name = "basCommon"
Option Explicit
'********************************************************************************
'
' 共通処理
'
'********************************************************************************

'******************************************************************
' 必須チェック
'
' 引数 strTarget：チェック対象の値
'
' 戻り値：Trueエラーあり
'         Falseエラーなし
'******************************************************************
Public Function IsBlank(strTarget As String, strMes As String) As Boolean
  IsBlank = False
  ' 未入力チェック
  If Trim(strTarget) = "" Then
    MsgBox strMes & "を入力してください。", vbCritical, "エラーメッセージ"
    IsBlank = True
    Exit Function
  End If

End Function

'******************************************************************
' ファイルのチェック
'
' 引数 path：ファイル名
'
' 戻り値：Trueエラーあり
'         Falseエラーなし
'******************************************************************
Public Function FileExist(path As String, strMes As String) As Boolean
  FileExist = False
  ' 存在チェック
  With CreateObject("Scripting.FileSystemObject")
    If Not .FileExists(path) Then
      MsgBox strMes & "が存在しません。", vbCritical, "エラーメッセージ"
      Exit Function
    End If
  End With
  
  FileExist = True
End Function

'******************************************************************
' フォルダのチェック
'
' 引数 path：ファイル名
'
' 戻り値：Trueエラーあり
'         Falseエラーなし
'******************************************************************
Public Function FolderExist(path As String, strMes As String) As Boolean
  FolderExist = False
  ' 存在チェック
  With CreateObject("Scripting.FileSystemObject")
    If Not .FolderExists(path) Then
      MsgBox strMes & "フォルダが存在しません。", vbCritical, "エラーメッセージ"
      Exit Function
    End If
  End With
  
  FolderExist = True
End Function


'******************************************************************
' フォルダ作成
'
' 引数 path：ファイル名
'
' 戻り値：Trueエラーあり
'         Falseエラーなし
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
' Excelシートから最終行の位置を取得
'
' 引数 targerSheet：対処のExcelシート
'      lngStartRowNo：明細行の開始位置
'      lngCountColumnNo1：最終行をカウントする列位置1
'      lngCountColumnNo2：最終行をカウントする列位置2
'
' 戻り値 最終行の位置、 1件もデータがない場合などは開始行位置を返す
'************************************************************************
Public Function GetMaxRowNo(targerSheet As Excel.Worksheet, lngStartRowNo As Long, lngCountColumnNo As Long) As Long

  GetMaxRowNo = targerSheet.Cells(Rows.Count, lngCountColumnNo).End(xlUp).Row
    
  ' 1件もデータがない場合などは開始行位置を返す
  If GetMaxRowNo < lngStartRowNo Then
    GetMaxRowNo = lngStartRowNo
  End If
  
End Function

'************************************************************************
'
' Excelシート内の指定した範囲でソート
'
' 引数 targerSheet：対処のExcelシート
'      lngRowStartNo：明細行の開始行位置
'      lngRowEndNo：明細行の終了行位置
'      lngColStartNo：明細行の開始列位置
'      lngColEndNo：明細行の終了列位置
'      lngSortColNo：ソートする項目の列位置
'
'************************************************************************
Public Sub SortSpecifiedRange(targerSheet As Excel.Worksheet, lngRowStartNo As Long, lngRowEndNo As Long, lngColStartNo As Long, lngColEndNo As Long, lngSortColNo As Long)

  targerSheet.Activate
  targerSheet.Range(Cells(lngRowStartNo, lngColStartNo), Cells(lngRowEndNo, lngColEndNo)) _
           .Sort Key1:=targerSheet.Cells(lngRowStartNo, lngSortColNo), order1:=xlAscending, DataOption1:=xlSortTextAsNumbers

End Sub

'************************************************************************
'
' Excelシート内の指定した範囲を配列で取得
'
' 引数 targerSheet：対処のExcelシート
'      lngRowStartNo：明細行の開始行位置
'      lngRowEndNo：明細行の終了行位置
'      lngColStartNo：明細行の開始列位置
'      lngColEndNo：明細行の終了列位置
'
'************************************************************************
Public Function GetSpecifiedRange(targerSheet As Excel.Worksheet, lngRowStartNo As Long, lngRowEndNo As Long, lngColStartNo As Long, lngColEndNo As Long) As Variant
  Dim myArray As Variant
  myArray = targerSheet.Range(Cells(lngRowStartNo, lngColStartNo), Cells(lngRowEndNo, lngColEndNo))

  GetSpecifiedRange = myArray
  
  Set myArray = Nothing

End Function








'******************************************************************
' 文字コードを指定してファイルを読込
'
' 引数 strFileName：ファイル名
'      strCharSet：文字コード(EUC-JP、Shift_JIS)
'
' 戻値 文字列の配列
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
' 文字コードを指定してファイルを書込み
'
' 引数 strFileName：ファイル名
'      strOutputData()：文字列の配列
'      strCharSet：文字コード(EUC-JP、Shift_JIS)
'      strLineSeparator：改行文字
'
' 戻値 文字列の配列
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
' Variant型をString型に変更する
' (実行時に型が一致しないと言われる場合に使う）
'
' 引数 varValue：対象データ
'
' 戻り値 上記
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



