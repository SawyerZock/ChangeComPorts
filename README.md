This simple program uses registry keys to reassign com ports. I have published this for anyone else who runs into this same problem under similar restrictions as me. Changing COM ports programatically should not be done, this is the job of the driver. If you lack the access or authority to correctly set up device policy to accept any dynamic COM port however, then this is your only option. This code is currently configured to reassign the default windows COM port from COM1 to COM9, then change the COM port for a Topaz signature pad to COM1. The COM ports and registry keys are hardcoded, but could easily be improved to prompt for the COM port. Change the registry paths to whatever you need for your particular problem. I hope this is helpful to anyone else who ran into the same restrictions. Enjoy!
