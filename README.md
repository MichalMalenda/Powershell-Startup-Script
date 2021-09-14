# Powershell-Startup-Script
## Everyone loves automating stuff but not many people don't know where to start...
With this startup script I will show you how, you can automate the startup process of your PC with something better than Windows built in Startup manager.

To make the script work it's always good to first take a look and adjust your settings as per instructions in the header of the file.
1. Once you save the file right click on it and select create shortcut.
    * Shortcut should appear in the same directory as the .ps1 file
    * Right click on the shortcut and select properties
    * In **Shortcut** tab you should see Target for the shortcut
    * Change the text in the Target to match this pattern <br />
        **"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -ExecusionPolicy Bypass "Path for your script"**
    * Click Ok to make sure the changes are saved
2. With the shortcut we have just created we can now move it to your startup folder.
    * Go to your **Start Menu** and type in **Run** and hit **ENTER**
    * In the next window type in **shell:startup** and hit **ENTER** again
    * In the newly opened window we can now paste our shortcut we created in previous step
3. Restart your PC and make sure everything is working as intended or you can just double click on the shortcut to see if everything is working as it should.