Attribute VB_Name = "Module3"
Sub Zadacha_�4()

 Dim fname As String, sFileName As String, sNewFileName As String '�������� ����������
 
 Dim h As Date '������ ���������� ���� , DATA
 
 Dim c As Date '������ ���������� ���� , DATA
 
  Dim lDaysCnt As Long
 
 c = #4/14/2020# ' �������� ������� ������
 
 
   
 
 MkDir "C:\Users\alexm\Desktop\�����\" '������ ����� �����
 
 fname = Dir("C:\Users\alexm\Desktop\1\" & "*.*") '������ ���������� ����� ���� ������




 
  
  Do While fname <> ""
  h = FileDateTime("C:\Users\alexm\Desktop\1\" & fname) ' ������� ���� �������� �����

 lDaysCnt = DateDiff("d", "4.14.2020", h) ' ������ ������� ����� ������
    
                                                        '������ ���� , ��� ����������� ���� ������
If lDaysCnt < 0 Then
 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' �������� , �� ��� �������
    sNewFileName = "C:\Users\alexm\Desktop\�����\" & fname ' �������� ���� �������
   
 
    FileCopy sFileName, sNewFileName ' �������� �����������
   End If
   fname = Dir
   
   Loop
 


End Sub

