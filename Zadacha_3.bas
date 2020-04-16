Attribute VB_Name = "Module2"



Sub Zadacha_¹3()
        
      Dim fname   As String
     
     
     fname = Dir("C:\Users\alexm\Desktop\1\" & "*.*")
      
      Open "C:\Users\alexm\Desktop\1\param.txt" For Output As #1
     
    
    Do While fname <> ""

 Print #1, fname


    
   
 
    
   
   fname = Dir
   
   Loop
   Close #1


       
    End Sub
    
  

