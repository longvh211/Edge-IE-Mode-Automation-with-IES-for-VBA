# Edge IE Mode Automation with IES for VBA
An alternative method to automate Edge IE Mode directly using VBA. This method is useful when the target environment is unable to perform the necessary windows upgrade to transist to Edge IE Mode Automation from the original Internet Explorer (see my StackOverflow's response below for further details here). This method is easy to setup and employ and once it is setup, coding for Edge IE Mode is just the same as coding for Internet Explorer.

**StackOverflow Response**

https://stackoverflow.com/questions/70619305/automating-edge-browser-using-vba-without-downloading-selenium/71994505?noredirect=1#comment128133297_71994505

**For Demo**

You can download the IES Framework Excel macro file (.xlsm) in this git for a demo or simply download the .zip package from the Release section. The codes are unlocked and can be viewed in the VBIDE screen. A demo example with instructions has also been prepared for your ease of teasting.

**For Installation**

Look for the core.bas module in the import folder. Alternatively, you can also copy the module found in the IES Framework Excel macro file. It is the same.

**Note**

This automation method only works for Edge IE Mode, not Edge entirely. For direct automation with Chromium-based browsers such as Chrome and Edge with VBA, see this git of mine instead:

https://github.com/longvh211/Chrome-Automation-with-CDP-for-VBA
