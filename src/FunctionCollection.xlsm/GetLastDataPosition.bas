Attribute VB_Name = "GetLastDataPosition"
Function getMaxRow(sht As Worksheet, targetCol As Long) As Long
  '������ɍŏI�s����������i���j
  'http://www.niji.or.jp/home/toru/notes/8.html
  
  getMaxRow = sht.Cells(sht.Rows.Count, targetCol).End(xlUp).Row

End Function
Function getMaxCol(sht As Worksheet, targetRow As Long) As Long
  '�������ɍŏI�����������i���j
  'http://www.niji.or.jp/home/toru/notes/8.html
  
  getMaxCol = sht.Cells(targetRow, sht.Columns.Count).End(xlToLeft).Column

End Function
