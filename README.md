# CheckExternal

## Purpose

This code with create a popup box whenever you send an email externally.  The popup will prompt you that you are about to send an external email and give you a yes/no/cancel for if you wish to continue.  The idea is, this will prevent you from accidentially sending emails externally which were meant for an internal reciever.  

![Picture of warning](https://github.com/IainShepherd/CheckExternal/blob/master/Warning%20Box.png)

The code can be edited such that it checks in more detail too if the feature is requested.  Such as for emails sent outside of a selection of people like a team.  This could be useful during crunch project times.  

## How to set up

Easiest is to copy the text from [CheckExternal.bas](https://github.com/IainShepherd/CheckExternal/blob/master/CheckExternal.bas) into your vba code.  To do that:
1.	Press Alt+F11
2.	Expand "Microsoft Outlook Objects"
3.	Click on "ThisOutlookSession"
4.	Paste the code in here
5.	Update your "HomeDomain" in the third line of the code to be your email domain
 
Now you'll need to make it so that macros can run else it won't do anything. 
 
1.	Go to File > Options
2.	Trust Center > Trust Center Settings
3.	Macro setting
4.	Now select "Notification for all macros"
5.	Ok it out
 
Finally, you'll need to close and re-open Outlook.  Make sure you save the VBA when it prompts you.  
 
Now try sending yourself an email.  Nothing should happen, it's internal! Then try sending someone external, like your personal account, an email, you should get a popup box!  You can click yes to send, or no/cancel to not send.  
