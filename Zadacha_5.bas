Attribute VB_Name = "Module4"
Sub Zadacha_�5()

 Dim fname As String, sFileName As String, sNewFileName As String '�������� ����������
 
 Dim sz As Long
 
 MkDir "C:\Users\alexm\Desktop\�����\" '������ ����� �����
 
 fname = Dir("C:\Users\alexm\Desktop\1\" & "*.*") '������ ���������� ����� ���� ������


 
 
 

 
 Do While fname <> ""  ' �������� ����
  

 sz = FileLen("C:\Users\alexm\Desktop\1\" & fname) ' ������� ������ ����� � ��������� ��� � ����������
    
                                                        
If (sz > 100000) Then ' ������� �����������
 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' �������� , �� ��� �������
    sNewFileName = "C:\Users\alexm\Desktop\�����\" & fname ' �������� ���� �������
   
 
    FileCopy sFileName, sNewFileName ' �������� �����������
   End If
   fname = Dir
   
   Loop
 


End Sub

