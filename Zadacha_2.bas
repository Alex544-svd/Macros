Attribute VB_Name = "Module1"
Sub Zadacha_�2()

  
    Dim fname   As String '������ ����������
     
     
     fname = Dir("C:\Users\alexm\Desktop\1\" & "*��� ��������.*") ' ������� Dir ������������ ����� ������ �� ���������� ���� ,
                                                        '���������� ����� , � ������ ������ "��� ��������"
     
   
 
     
    
    Do While fname <> ""            '������ ���� , ��� �������� ���� ������

 
     Kill "C:\Users\alexm\Desktop\1\" & fname ' �������� �����
   
 
    
   
   fname = Dir
   
   Loop

End Sub

