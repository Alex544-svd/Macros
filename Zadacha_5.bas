Attribute VB_Name = "Module4"
Sub Zadacha_є5()

 Dim fname As String, sFileName As String, sNewFileName As String 'объ€вл€ю переменные
 
 Dim sz As Long
 
 MkDir "C:\Users\alexm\Desktop\—клад\" '—оздаю папку —клад
 
 fname = Dir("C:\Users\alexm\Desktop\1\" & "*.*") '—трока возвращает имена всех файлов


 
 
 

 
 Do While fname <> ""  ' объ€вл€ю цикл
  

 sz = FileLen("C:\Users\alexm\Desktop\1\" & fname) ' получаю размер файла и записываю его в переменную
    
                                                        
If (sz > 100000) Then ' условие копировани€
 
    sFileName = "C:\Users\alexm\Desktop\1\" & fname  ' указываю , то что копирую
    sNewFileName = "C:\Users\alexm\Desktop\—клад\" & fname ' указываю куда копирую
   
 
    FileCopy sFileName, sNewFileName ' операци€ копирований
   End If
   fname = Dir
   
   Loop
 


End Sub

