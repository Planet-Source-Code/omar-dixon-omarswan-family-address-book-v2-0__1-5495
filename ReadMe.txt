
************************************************************************
************************************************************************
*** Family Address Book v2.0                                         
*** Author    : Omar                                                 
*** Web Page  : http://www.omarswan.cjb.net                          
*** Em@il     : omarswan@yahoo.com                                   
*** Copyright 1997-1998 by SmileyOrange inc.                         
*** If you have any suggestions or questions, I'd be happy to hear it
***                                                                  
*** ***  JESUS IS LORD  ***                      
***                                                                  ************************************************************************
************************************************************************


************************ DESCRIPTION ***********************************
************************************************************************
*** This is a family address book program created using VB6. Its      
*** main purpose is to store personal address book information for    
*** each member of a family. The program uses two different types of 
*** users (1. Administrator) (2. User)                                
************************************************************************    
************************************************************************


************************** REQUIREMENTS ********************************
************************************************************************
*** Microsoft Windows Common Control 5.0(SP2) - COMCTL32.OCX         
*** Microsoft DAO 3.51 Object Library                                
*** Visual Basic 6 Runtime Files                                     ************************************************************************
************************************************************************


**************************** WARNING ***********************************
************************************************************************
***                                                                  
*** I take no responsibility for the use of, or the results from the 
*** use of this program. If you are not sure about it,"DON'T USE IT" 
***                       " USE AT YOUR OWN RISK "                   
***                                                                  
*** Note : I don't expect that there should be any problems..... (o:                                                                ***                            I Think                               
************************************************************************
************************************************************************


********************* THINGS YOU SHOULD KNOW ***************************
************************************************************************
***                                                                  
*** [1] The name of the database file is "Family.FM2"                
***                                                                  
*** [2] The program uses the file "Family2.ini" file to store the    
***     the PATH and NAME of the database. If the database is not    
***     found in the DATABASE PATH that is set in "Family2.ini" the  
***     program will ask you to specify the correct PATH or if you   
***     want RECREATE it.                                            
***                                                                  
***     *Note you will have to recreate the database if you are      
***      running the database for the very first time of if you have 
***      deleted it.                                                 
***                                                                  
*** [3] When the database file has been recreate the default values  
***     are as follows :                                             
***                                                                  
***     Login Name   :Admin                                          
***     Password     :Admin                                          
***     Access Level :Administrator                                                                                                    ***
*** Note 1 : The values for LOGIN NAME & PASSWORD are Case Sensitive   
*** Not  2 : If you are running the program for the ver first time you
***          have to create the database file and the select it.  
************************************************************************
************************************************************************


************************* TYPES OF USERS *******************************
************************************************************************
*** This Program uses two types of uses :                            
***                                                                  
*** [1] Administrator                                                
***      + Allowed to add and remove users.                          
***      + Allowed to edit other users profile.  
***      + Allowed to Loggout User using Fam-Fix2.exe                    
***                                                                  
*** [2] User                                                         
***      + Not allowed to add and remove users.                      
***      + Not allowed to edit other users profile.                  
***      + Not allowed to Loggout User using Fam-Fix2.exe      
***
************************************************************************
************************************************************************


************************* MISC INFO ************************************
************************************************************************
*** Never use another software or program to edit the contents of    
*** the database file, this will cause serious damages to the file.  
***               								 
*** The should at least be one (1) Administrator in the database     
***        									 
*** To log-out a user in the event that he/she did not log-out       
*** correctly, that user cannot login again until and Administrator  
*** uses the file "FAM-Fix2.exe" or the Menu Option "Logout User",   
*** to log-out that user. 							 
*** *Note this can only be done by an Administrator.                 
*** *Never log-out a user if he/she is still using the database,     
***  this might lead to data loss.                     		 
***        									 
*** When your are on the Search Form you can Delete or Print records
*** by simply Right-Mouse clicking on a record und the First Name Column       
*** a popup menu will appear. Note you have to click the search button
*** for the records to appear.                                                                                  
************************************************************************
************************************************************************


************************************************************************
************************************************************************
*** If you have any questions, problems, suggestions or you find a   
*** bug please contact me.        					 
*** Em@il : omarswan@yahoo.com 						 
*** URL   : http://www.omarswan.cjb.net 					 
************************************************************************
************************************************************************
.-=*-*-*-*-*-*-*-*-*-*-*-* TO GOD BE THE GLORY *-*-*-*-*-*-*-*-*-*-*-*=-.
