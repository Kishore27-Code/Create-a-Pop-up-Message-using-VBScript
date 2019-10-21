# Create-a-Pop-up-Message-using-VBScript
The "Popup" method produces a pop-up message box which can display a message to a user for a specified amount of time. If the message time is omitted or set to zero, the pop-up will remain until the user dismisses the message. In addition, a title can be assigned to the pop-up message.  If it is omitted, the default is "Windows Script Host"


intReturn = WScript.CreateObject("WScript.Shell").Popup("Content", 20, "Header", 0 + 32)

'Return Value = WScript.CreateObject("WScript.Shell").Popup(Content,Seconds To Wait,Title,type answer + Icon Types)

Content
String value containing the text you want to appear in the pop-up message box.

Seconds To Wait
Optional. Numeric value indicating the maximum length of time (in seconds) you want the pop-up message box displayed.

Title
Optional. String value containing the text you want to appear as the title of the pop-up message box.

Icon Types
Optional. Numeric value indicating the type of buttons and icons you want in the pop-up message box. These determine how the message box is used.

type answer
Integer value indicating the number of the button the user clicked to dismiss the message box. This is the value returned by the Popup method.


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
