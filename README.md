# utl-avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel
Avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel 
    Avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel                                                
                                                                                                                    
    I did not want the numbers stored as text.                                                                      
                                                                                                                    
                                                                                                                    
    I could not get SAS ods excel, R XLconnect or R xlxs to eliminate the 'green                                    
    triange'. The numbers were stored as text.                                                                      
                                                                                                                    
    github                                                                                                          
    https://tinyurl.com/y4rt2jqq                                                                                    
    https://github.com/rogerjdeangelis/utl-avoid-storing-numbers-as-text-when-exporting-mixed-type-to-excel         
                                                                                                                    
    StackOverflow                                                                                                   
    https://tinyurl.com/y44uvdul                                                                                    
    https://stackoverflow.com/questions/57715415/convert-number-stored-as-text-in-excel-file-using-r                
                                                                                                                    
    R2evans profile                                                                                                 
    https://stackoverflow.com/users/3358272/r2evans                                                                 
                                                                                                                    
    *_                   _                                                                                          
    (_)_ __  _ __  _   _| |_                                                                                        
    | | '_ \| '_ \| | | | __|                                                                                       
    | | | | | |_) | |_| | |_                                                                                        
    |_|_| |_| .__/ \__,_|\__|                                                                                       
            |_|                                                                                                     
    ;                                                                                                               
                                                                                                                    
                                                                                                                    
    options validvarname=upcase;                                                                                    
    libname sd1 "d:/sd1";                                                                                           
    data sd1.have;                                                                                                  
       set sashelp.class(keep=name age sex);                                                                        
       if mod(_n_,3)=0 then name=put(99*_n_,5. -l);                                                                 
    run;quit;                                                                                                       
                                                                                                                    
                                                                                                                    
    SD1.HAVE total obs=19     |  RULES                                                                              
                              |                                                                                     
      NAME       SEX    AGE   |                                                                                     
                              |                                                                                     
      Alfred      M      14   |                                                                                     
      Alice       F      13   |                                                                                     
                              |                                                                                     
      297         F      13   |  * do not store as text                                                             
                              |    with green triangle                                                              
                                                                                                                    
      Carol       F      14   |                                                                                     
      Henry       M      14   |                                                                                     
                              |                                                                                     
      594         M      12   |  * do not store as text                                                             
                                                                                                                    
                                                                                                                    
    *            _               _                                                                                  
      ___  _   _| |_ _ __  _   _| |_                                                                                
     / _ \| | | | __| '_ \| | | | __|                                                                               
    | (_) | |_| | |_| |_) | |_| | |_                                                                                
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                               
                    |_|                                                                                             
    ;                                                                                                               
                                                                                                                    
    d:/xls/wantopn.xlsx                                                                                             
                                                                                                                    
          +--------------------------------------+                                                                  
          |     A      |    B       |     C      |                                                                  
          +--------------------------------------+                                                                  
       1  | NAME       |   SEX      |    AGE     |                                                                  
          +------------+------------+------------+                                                                  
       2  | ALFRED     |    M       |    14      |                                                                  
          +------------+------------+------------+                                                                  
       3  | ALICE      |    M       |    14      |                                                                  
          +------------+------------+------------+                                                                  
       4  | 297        |    M       |    14      |                                                                  
          +------------+------------+------------+                                                                  
           ...                                                                                                      
          +------------+------------+------------+                                                                  
       20 | 594        |    M       |    15      |                                                                  
          +------------+------------+------------+                                                                  
                                                                                                                    
    *                                                                                                               
     _ __  _ __ ___   ___ ___  ___ ___                                                                              
    | '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                             
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                             
    | .__/|_|  \___/ \___\___||___/___/                                                                             
    |_|                                                                                                             
    ;                                                                                                               
                                                                                                                    
    %utl_submit_r64('                                                                                               
    library(openxlsx);                                                                                              
    library(haven);                                                                                                 
    want<-read_sas("d:/sd1/have.sas7bdat");                                                                         
    wb = createWorkbook();                                                                                          
    addWorksheet(wb, "want");                                                                                       
    writeDataTable(wb, "want", want);                                                                               
    for (cn in seq_len(ncol(want))) {                                                                               
      for (rn in seq_len(nrow(want))) {                                                                             
        if (!is.numeric(want[rn,cn]) && !is.na(val <- as.numeric(as.character(want[rn,cn])))) {                     
          writeData(wb, "want", val, startCol = cn, startRow = 1L + rn)                                             
        };                                                                                                          
      };                                                                                                            
    };                                                                                                              
    saveWorkbook(wb, file = "d:/xls/wantopn.xlsx", overwrite = TRUE)                                                
    ');                                                                                                             
                                                                                                                    
