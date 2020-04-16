Attribute VB_Name = "Module1"
Sub Zadacha_№2()

  
    Dim fname   As String 'Создаю переменную
     
     
     fname = Dir("C:\Users\alexm\Desktop\1\" & "*для удаления.*") ' функция Dir осуществляет поиск файлов по указанному пути ,
                                                        'указанному имени , в данном случае "для удаления"
     
   
 
     
    
    Do While fname <> ""            'Создаю цикл , для удаления всех файлов

 
     Kill "C:\Users\alexm\Desktop\1\" & fname ' удаление файла
   
 
    
   
   fname = Dir
   
   Loop

End Sub

