intReturn = WScript.CreateObject("WScript.Shell").Popup("Content", 20, "Header", 0 + 32)

'Return Value = WScript.CreateObject("WScript.Shell").Popup(Content,Seconds To Wait,Title,type answer + Icon Types)

'type answer
'    0   OK button. 
'    1   OK and Cancel buttons. 
'    2   Abort, Retry, and Ignore buttons. 
'    3   Yes, No, and Cancel buttons. 
'    4   Yes and No buttons. 
'    5   Retry and Cancel buttons. 
   
'Icon Types
   
'   Value Description 
'    16  "Stop Mark" icon. 
'    32  "Question Mark" icon. 
'    48  "Exclamation Mark" icon. 
'    64  "Information Mark" icon. 

'return value: 
'   1  OK button 
'   2  Cancel button 
'   3  Abort button 
'   4  Retry button 
'   5  Ignore button 
'   6  Yes button 
'   7  No button	

''''''''''''''''''''''''''''''''''''''
