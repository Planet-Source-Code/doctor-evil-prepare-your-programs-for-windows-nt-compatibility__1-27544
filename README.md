<div align="center">

## Prepare Your Programs for Windows NT Compatibility\!


</div>

### Description

Windows XP is coming. With that, everyone will be moving to the NT platform. This article goes through some tips to prepare your programs for compatibility with the NT platform.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Doctor Evil](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/doctor-evil.md)
**Level**          |Beginner
**User Rating**    |4.6 (46 globes from 10 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/doctor-evil-prepare-your-programs-for-windows-nt-compatibility__1-27544/archive/master.zip)





### Source Code

<font face="Tahoma,Verdana,Arial">
<font size="+3"><b>Windows NT Compatibilty</b></font><br>
With the release of Windows XP, everyone will be moving to the NT platform. This means that you have to prepare your apps to
run on the NT platform today! You can no longer just blow off NT compatibility. Did you know that Microsoft is dropping
support for Windows 95 at the end of this year, and will drop Support for Windows 98 in 2 years? Windows 9x is a thing of
the past.<p>
<font size="+2"><b>New Visual Themes</b></font><br>
Microsoft has totally redesigned the User Interface (UI) in this version of Windows. The new Visual Themes are a
semi-problem. All of your VB apps will have the new titlebar, close/min/max buttons, and border, but the standard VB
controls (like the button) will not. Realize, however, that simply replacing your buttons with the XP look-alike ones is
<b>NOT a solution.</b> For one reason, it is not just the buttons that have been redesigned. Every control has been. Even
the frame. Also, if the user decides not to use the default theme, then your buttons will look strange because they use the
style of the Windows default theme, not the one the user is using. Be aware that this does not affect menus. The VB menus
have been drawing in the new XP style menu. (including the alpha-blended shadows)<br><b>UPDATE:</b> I have created a button control that actually calls the Theme APIs to draw the button. You can <a href="http://www.planetsourcecode.com/xq/ASP/txtCodeId.27673/lngWId.1/qx/vb/scripts/ShowCode.htm" target="_new">get it here.</a><p>
<font size="+2"><b>Application Compatibility</b></font><br>
Chances are, your app will <i>probabally</i> run on XP. If you have access to a Windows 2000 machine, test your app on it.
If it runs, chances are that it will run on XP.<br>Since XP is WinNT, if you have used any code in your projects that is
labeled "does not work in NT", it won't run in XP. Usually, 9x-only code doesn't work in NT due to the security
restrictions.<br>
<b>Some examples of what WON'T work:</b><br>
<bl><li>Anything that has to do with hiding your app from the CTRL-ALT-DEL list. The NT Task Manager shows ALL processes, no
matter what. Even system services show up. In fact, the RegisterServiceProcess API that many of you use to hide your apps
from the list doesn't exsist under NT.</li>
<li>Most code to shut down your computer. The app must get more token privliges before it calls the shutdown API. (assuming
the user has enough rights to shut down anyways) If you would like some good code to shut down a NT system, see <a
href="http://vbaccelerator.com/tips/vba0019.htm" target="_new">this page</a>.</li>
<li>Code that writes to the Registry. Specifically, to anything besides HKEY_CURRENT_USER. HKLM writes fail if the current
user is not an Administrator.</li>
<li>If you try to write to different areas of the hard drive. Many users can't <i>write</i> to many locations due to NTFS
restrictions.</li></bl>
<br>
Basically, if the user has Admin rights, chances are that your app will run. But if they are not, you might have
problems.<p>
<font size="+2"><b>Detecting The Operating System</b></font><br>
At a minimum, you should detect the OS that your program is running under. This way, if you are sure that your app won't run
under NT, you can simply display a message box informing the user that their OS is not supported and that they should go to
your website to download a new version of your software that does support NT, then exit your program. This prevents nasty VB
error message boxes that will only confuse the user. (or even GPFs) Here is some code that will detect if your app is
running under NT or not: (You can put this into a .MOD)<pre>
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Type OSVERSIONINFO
 OSVSize As Long
 dwVerMajor As Long
 dwVerMinor As Long
 dwBuildNumber As Long
 PlatformID As Long
 szCSDVersion As String * 128
End Type
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
 (lpVersionInformation As OSVERSIONINFO) As Long
Public Function RunningUnderNT() As Boolean
 Dim OSV As OSVERSIONINFO
 OSV.OSVSize = Len(OSV)
 If GetVersionEx(OSV) = 1 Then
 If OSV.PlatformID = VER_PLATFORM_WIN32_NT And Then
		RunningUnderNT = True
	 End If
 End If
End Function
</pre>
Then you can call the function during the startup of your app.<pre>
Private Sub Form_Load()
If RunningUnderNT Then
	MsgBox "Sorry. This program does not support your operating system." & vbCrLf & vbCrLf & "Please go to my website at
[address] to download the latest version of this program, which may support this version of Windows.", vbCritical, "Windows
Version Not Supported"
	End
End If
...
End Sub
</pre>
While your program is running, you can check the version of Windows before making any API calls that won't work in NT. You
can either state that the requested function is not available, or substitute NT-compatible calls in place of the 9x ones.<p>
<font size="+2"><b>The NTFS File System</b></font><br>
For those of you who have used a NT OS before, you know what NTFS is and what it does. To the rest of you, NTFS will be a
new suprise. While upgrading from Windows 9x/ME to XP, Setup will convert drives from FAT16/32 to NTFS. (You can choose not
to, but you really should anyways.) NTFS allows Administrators to set access privlages to certain folders and even certain
files. You can be granted anywhere from no access at all to full read/write access to a file/folder, and everything in
between. So, it is possible that a user can read a file (even your application and it's folder) but not write to it. The
probelm comes in when your program tries to write to the disk, or even if it tries to read a file. What if your program
writes a configuration file, but the current user is not allowed to write to the folder? Not only will your program error
out if you didn't put error handlers in, but the settings will never be saved. This is obviously not good. You should try
to save settings to the HKEY_<i>CURRENT_USER</i> section of the Registry, or allow the user to pick where they want to save
your configuration file to. Then save the the path to the config file in the Registry. This way, you can be sure that
everything works out OK. Also, be sure to <b><i>ALWAYS</i></b> put in error handlers in any sub/function that writes to the
Registry or the hard drive. Of course, it is a good idea to put error handlers in any sub or function to handle anything
that you didn't anticipate. I as a user can't stand when a VB app (or any other app for that matter) didn't use error
handlers, then comes up with that nasty little VB error message box, causing the program I'm using to close.<p>
 <p>
Remeber, NT <i>cannot</i> just be blown off anymore. You <b>have</b> to make your programs compatible with NT now. Good
luck, and happy coding!
</font>

