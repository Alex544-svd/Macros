Attribute VB_Name = "NewMacros"




Sub Zadacha_№1()



    Dim sFileName As String, sNewFileName As String 'Создаю переменные
    Dim fname   As String
     
     
     fname = Dir("C:\Users\alexm\Desktop\1\" & "*txt") ' функция Dir осуществляет поиск файлов по указанному пути ,
                                                        'указанному расширению
     
    
 
     MkDir "C:\Users\alexm\Desktop\2\" 'Создаю папку
    
    Do While fname <> ""            'Создаю цикл , для копирования всех файлов

 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' указываю , то что копирую
    sNewFileName = "C:\Users\alexm\Desktop\2\" & fname ' указываю куда копирую
   
 
    FileCopy sFileName, sNewFileName ' операция копирований
   
   fname = Dir
   
   Loop

End Sub
