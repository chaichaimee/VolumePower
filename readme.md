![NVDA Logo](https://www.nvaccess.org/files/nvda/documentation/userGuide/images/nvda.ico)

# VolumePower

Take command of your sound and your power button, and never lose NVDA's voice again.

**author:** chai chaimee

**url:** https://github.com/chaichaimee/VolumePower

## Introduction

VolumePower gives you quick, keyboard-only control over the sound you hear from NVDA and over your computer's power. You can raise or lower NVDA's speech volume, adjust the left and right ear levels independently, and shut down or restart Windows with a single key combination, protected by a short cancellable countdown.

VolumePower also acts as a guard for your audio setup. When NVDA starts, it can unmute Windows sound, restore a volume that is too low, switch back to the default output device, warn you if a playback device has disappeared, and keep trying to bring your voice synthesizer back if it failed to load. The result is that NVDA stays audible, so you are never left with a silent computer.

## Hot Keys

**NVDA+Shift+PageUp**

- Increase NVDA volume by 5%

**NVDA+Shift+PageDown**

- Decrease NVDA volume by 5%

**Alt+Windows+S**

- Single Tap : Shutdown computer
- Double Tap : Restart computer

**NVDA+Shift+NumPad Plus**

- Open the ear volume layer

**Inside the ear volume layer**

- Left Arrow : Select the left ear
- Right Arrow : Select the right ear
- Up Arrow : Increase the selected ear by 5
- Down Arrow : Decrease the selected ear by 5
- Enter : Make both ears equal
- Escape : Close the layer
- Any other key : Close the layer

## Features

### NVDA Volume Control

Press **NVDA+Shift+PageUp** or **NVDA+Shift+PageDown** to change the volume of your current voice synthesizer in steps of 5%. The volume stays between 0% and 100%, NVDA announces the new level, for example "55%", and the new level is saved so it is still in effect the next time NVDA starts.

### Left and Right Ear Volume

This feature lets you set the loudness of the left and right speaker separately. It changes the left and right channel levels of your Windows default playback device, which is useful when one ear needs more volume than the other.

**Step by step:**

1. Press **NVDA+Shift+NumPad Plus**. The layer opens and NVDA announces the left ear level, for example "Left 40".
2. Press **Left Arrow** or **Right Arrow** to choose the ear. NVDA announces the level of the ear you chose.
3. Press **Up Arrow** to raise the chosen ear by 5, or **Down Arrow** to lower it by 5. NVDA announces the new level, and adds "maximum" or "minimum" when you reach the end of the range.
4. Repeat steps 2 and 3 for the other ear if needed.
5. Press **Enter** to make both ears equal. Both ears are set to the louder of the two levels, and NVDA announces "Both ears" followed by the level.
6. Press **Escape** to close the layer. NVDA announces "Ear volume closed". Pressing any other key also closes the layer.

Each change is accompanied by a short beep that is heard only in the ear you are adjusting. The pitch of the beep rises as the level goes up, so you can judge the level by ear as well as by speech.

If Windows audio is muted, NVDA announces "Audio is muted". If the output device is not stereo, NVDA announces "The output device is not stereo". If the change cannot be applied, NVDA announces "Failed to change ear volume". In all of these cases the layer closes.

### Shutdown and Restart

Press **Alt+Windows+S** once to shut down the computer, or twice to restart it. When the action begins, NVDA announces "Shutdown" or "Restart". Programs that are not responding are closed so the process is not blocked.

If Windows refuses the request, NVDA announces the error and plays a warning beep.

### Countdown Confirmation Dialog

To protect you from shutting down by accident, a countdown dialog appears before the action is carried out. This is enabled by default and can be turned off in the add-on settings.

**Step by step:**

1. Press **Alt+Windows+S** once for Shutdown or twice for Restart.
2. A dialog opens and NVDA reads, for example, "Shutdown in 3. Press any key to cancel." A beep plays.
3. The countdown continues with "2..." and "1...". Each number has its own beep, and the pitch rises as the countdown approaches zero.
4. If you do nothing, the countdown ends and the computer shuts down or restarts.
5. To stop it, press any key or close the dialog at any point during the countdown. NVDA announces "Cancelled" and nothing happens.

If the countdown dialog is turned off in the settings, the action is carried out immediately.

### Audio Guard at Startup

Each time NVDA starts, VolumePower checks your sound setup in the background and repairs what it can:

1. It unmutes the default playback device. If the Windows volume is lower than your chosen minimum level, it raises it to your custom volume level.
2. It unmutes NVDA's own audio session, and raises it in the same way if it is below the minimum level.
3. If enabled, it unmutes every active playback device, not only the default one.
4. If enabled, it compares your active playback devices with the ones from your last session. If a device that was working before is now disabled or unavailable, NVDA warns you to check Sound settings and plays a beep.
5. If enabled, it switches NVDA back to the default audio output device and reloads your synthesizer.
6. If enabled, it plays a confirmation sound when the audio has been successfully turned on.

### Synthesizer Recovery

If your configured voice synthesizer fails to start with NVDA, VolumePower keeps trying to load it again. The wait between attempts gets longer each time, up to a maximum of 30 seconds, so NVDA stays responsive. You can limit the number of attempts or let it try until it succeeds. When the synthesizer is back, the confirmation sound plays if it is enabled.

### Prevent Windows from Sleeping

When this option is enabled, Windows will not go to sleep automatically while NVDA is running. The setting is released when the add-on is disabled or NVDA exits.

### Add-on Settings

Open the NVDA Settings dialog and choose the **VolumePower** category. The following options are available:

1. **Show countdown confirmation dialog before shutdown or restart** : on by default.
2. **Set custom volume level** : the volume used when Windows sound is found to be too low. The default is 20.
3. **Increase the volume if it is lower than** : the minimum level that triggers the volume restore. The default is 5.
4. **Repeat attempts to initialize the voice synthesizer driver** : on by default.
5. **Number of retries (0 - infinitely)** : shown only when the option above is checked. The default is 0, which means keep trying until it succeeds.
6. **Switch to the default audio output device** : on by default.
7. **Play sound when audio has been successfully turned on** : on by default.
8. **Unmute all playback devices, not only the default one** : on by default.
9. **Alert me if a playback device becomes disabled or unavailable** : on by default.
10. **Prevent Windows from going to sleep automatically** : off by default.

Your settings are saved when you press OK.

## Support Me

If this tool has made your life easier, consider fueling the next update with a small donation.

[![Support me](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Your support means the world. Let's build something great together

© 2026 Chai Chaimee NVDA Add-on Released under GNU GPL
