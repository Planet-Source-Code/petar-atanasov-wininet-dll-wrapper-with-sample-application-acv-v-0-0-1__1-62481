Wininet.dll wrapper with sample application (ACV v.0.0.1.)

I've tried to code something like link harvester using the standart MS Wininet control, 
but it occured to be impossible task. 
Under win 2003 (and after continios testing under XP and 2000 too) the application 
didn't terminate (becouse ot the ms wininet control) after exit and that's way 
I've had to find another way. 

Now with the use if wininet API ACV is more than twice faster and really more stable. 

ACV itself represents a small example use of the wininet.dll wrapper. 
I'll create an FTP sample application later in order to show how to use ftp functions 
of the wrapper as well as example code for some of the others functions. 

ACV uses regular expressions when user perform searches against given directory and also when 
stripping url to server name and adjacent folders. 

The deployed application uses dbmon for reference counting and debuging.

You can reach me at: ppa_info@hotmail.com
09/2005