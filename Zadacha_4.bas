Attribute VB_Name = "Module3"
Sub Zadacha_№4()

 Dim fname As String, sFileName As String, sNewFileName As String 'объявляю переменные
 
 Dim h As Date 'создаю переменные типа , DATA
 
 Dim c As Date 'создаю переменные типа , DATA
 
  Dim lDaysCnt As Long
 
 c = #4/14/2020# ' указываю условие отбора
 
 
   
 
 MkDir "C:\Users\alexm\Desktop\Архив\" 'Создаю папку Архив
 
 fname = Dir("C:\Users\alexm\Desktop\1\" & "*.*") 'Строка возвращает имена всех файлов




 
  
  Do While fname <> ""
  h = FileDateTime("C:\Users\alexm\Desktop\1\" & fname) ' получаю дату создания файла

 lDaysCnt = DateDiff("d", "4.14.2020", h) ' нахожу разницу между датами
    
                                                        'Создаю цикл , для копирования всех файлов
If lDaysCnt < 0 Then
 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' указываю , то что копирую
    sNewFileName = "C:\Users\alexm\Desktop\Архив\" & fname ' указываю куда копирую
   
 
    FileCopy sFileName, sNewFileName ' операция копирований
   End If
   fname = Dir
   
   Loop
 


End Sub

