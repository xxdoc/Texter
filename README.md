# Texter
VB6 program to text messages or file contents to cell phones.
Use Texter to send individual messages. It can also be used to check files for changes and send the file contents as a text. 
For example it could be used with the TempMonitor Server program to text messages when temperature sensors exceed set limits.
It does this by checking the archive bit on a file. If the archive bit is set then the contents have changed and the text is sent.
After the text is sent the archive bit is cleared, ready for the next message. TempMonitor Server saves temperature alarms in the 
file 'Status.txt' in the folder 'Users\Public Documents\TemperatureMonitor\Common'. Point Texter to this file in the Setup 
dialog.
