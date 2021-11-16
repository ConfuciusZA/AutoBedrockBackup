# BedrockBackup - Forked from NeoLich/AutoBedrockBackup
Simplification of the NeoLich's script to backup your Windows based Minecraft Bedrock Server 

Single script designed to run on demand (eg. manually or through task scheduler)

The script is written in vbs

The vbs file should be copied to the same directory as your server.
By default, it will look for the 'bedrock_server.exe' as the server executable. 
If you have renamed yours, update the script accordingly. 

No paths are hardcoded, script will look for your 'worlds' folder and a 'backups' folder (will create the latter if it doesn't exist)

On running the script, it will check if the server executable is running 
  If running,  it will give a 5 & 1 minute warning before stopping the server, copying the files into a subfolder within your specificied backup folder, and restarting the server. 
  If not it will copy the files and start the server.   
