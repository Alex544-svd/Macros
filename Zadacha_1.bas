Attribute VB_Name = "NewMacros"




Sub Zadacha_�1()



    Dim sFileName As String, sNewFileName As String '������ ����������
    Dim fname   As String
     
     
     fname = Dir("C:\Users\alexm\Desktop\1\" & "*txt") ' ������� Dir ������������ ����� ������ �� ���������� ���� ,
                                                        '���������� ����������
     
    
 
     MkDir "C:\Users\alexm\Desktop\2\" '������ �����
    
    Do While fname <> ""            '������ ���� , ��� ����������� ���� ������

 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' �������� , �� ��� �������
    sNewFileName = "C:\Users\alexm\Desktop\2\" & fname ' �������� ���� �������
   
 
    FileCopy sFileName, sNewFileName ' �������� �����������
   
   fname = Dir
   
   Loop

End Sub
