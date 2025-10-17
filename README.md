# SWAP
SmartWindowsAudioProfiles (SWAP)

This software is not affiliated and is an independent software that needs EarTrumpet (WITH CLI COMMAND FEATURE) to work properly.
This version with CLI command can be found from my fork repo here https://github.com/dayeggpi/EarTrumpet (to compile it yourself) or attached in this SWAP repo pre-compiled for Windows (until EarTrumpet accepts my PR to merge the code for CLI command into the official repo: https://github.com/File-New-Project/EarTrumpet/pull/1765).


![Image](https://i.imgur.com/FdHoP3b.gif)

## Usage
Launch with `python swap.py`

You can also compile it to an exe by having the *.ico and *.spec file in same folder as swap.py and doing `pyinstaller swap.spec`. The output exe will be in the "dist" folder.
Then you simply execute the exe file to launch it.

Once launched, the app will ask you to provide the path where EarTrumpet (with CLI command) is installed.

Once done, click on "Test EarTrumpet" to ensure that the app is properly linked to SmartWindowsAudioProfiles.

Go to "Audio Devices" tab and Refresh Device List.

You can now go to"Profiles" and create a new profile then start adding rules.

The rules will be to set an Input audio device and an Output audio devices for a given running application.

Each rule is in two parts (input+output).

Once done, save and you can create another profile.

When you wish, you can then activate a profile by selecting it, and clicking on "Activate profile".

Alternatively, you can also activate a profile via command line as such : `SWAP.exe PROFILE_NAME` where PROFILE_NAME is the name of your profile.

![Image]()

## Config
A config.ini file will be generated to adjust some settings.

```
[App]
eartrumpet_path = C:/Program Files/EarTrumpet/EarTrumpet.exe
auto_save = True
```
adjust "eartrumpet_path" as per the path to EarTrumpet.exe (it can be as is if EarTrumpet.exe is in the PATH environement, or in same folder as SWAP.exe)

adjust "auto_save" to True or False to save automatically any changes done on profiles/rules.

Note: no need to adjust this file manually, all can be done via the GUI.

## Profiles
A audio_profiles.json file will be generated with your profiles and respective rules. 
You can programatically generate it as well following this format (this is an exemple of a profile named "PROFILE_NAME" with 1 rule (input+output) for the chrome.exe application:

```
{
  "profiles": {
    "PROFILE_NAME": {
      "rules": [
        {
          "app_name": "chrome.exe",
          "device_id": "VB-Audio Virtual Cable A\\Device\\CABLE-A In 16ch\\Render",
          "device_name": "VB-Audio Virtual Cable A",
          "item_id": "{0.0.0.00000000}.{43fac67b-ac3d-4287-a3aa-b4c5a678a3ec}",
          "name": "CABLE-A In 16ch",
          "direction": "Render"
        },
        {
          "app_name": "chrome.exe",
          "device_id": "VB-Audio Virtual Cable A\\Device\\CABLE-A Output\\Capture",
          "device_name": "VB-Audio Virtual Cable A",
          "item_id": "{0.0.1.00000000}.{23b82c5e-caa6-433a-b138-d72333cfea13}",
          "name": "CABLE-A Output",
          "direction": "Capture"
        }
      ]
    }
  }
}
```

Note: no need to adjust this file manually, all can be done via the GUI.
