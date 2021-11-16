# BedrockBackup - Forked from NeoLich/AutoBedrockBackup
Simplification of the NeoLich's script 
Single script designed to run on demand (eg. manually or through task scheduler)

The script is written in vbs and with variables declared in it

The vbs file should be copied to the same directory as your server.

On running the script, it will check if the server executable is running 
  If it is it will give a 5 & 1 minute warning before stopping the server, copying the files, and restarting the server. 
  If it is not running it will copy the files and start the server. 
  
