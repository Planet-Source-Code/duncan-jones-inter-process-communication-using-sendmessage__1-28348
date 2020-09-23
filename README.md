<div align="center">

## Inter process communication using Sendmessage


</div>

### Description

This articles shows how you can use windows messages to communicate between two (or more) applications.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Intermediate
**User Rating**    |4.9 (73 globes from 15 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-inter-process-communication-using-sendmessage__1-28348/archive/master.zip)





### Source Code

<p align=center>
  <font face="Arial">
  <strong>Inter process communication using registered messages from Visual Basic</strong>
  </font></p>
  <p align=justify><font face=Arial> One of the simplest ways to implement multi-tasking in Visual Basic is to create a seperate
  executable program to do each task and simply use the <b>Shell</b> command to run them as neccessary. The only problem with this is
  that once a program is running you need to communicate with it in order to control its operation.<br>
   One way of doing this is using the <b>RegisterWindowMessage</b> and <b>SendMessage</b> API calls to create your own particluar
  windows messages and to send them between windows thus allowing you to create two or more programs that <i>communicate</i> with each
  other.<br>
   In this example the server has the job of watching a printer queue and sending a message to every interested client whenever an event (job
  added, driver changed, job printed etc.) occurs.
  </font></p>
  <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow"> 1. Specifying your own unique messages</font></p>
  <p align=justify> <font face="Arial"> Windows communicate with each other by sending each other <a href="http://www.merrioncomputing.com/EventVB/WindowMessages.html">standard
   windows messages</a> such as <b>WM_CLOSE</b> to close and terminate the window.<br>
    There are a large number of standard messages which cover most of the standard operations that can be performed by and to different windows. However if you want to implement your
   own custom communication you need to create your own custom messages. This is done with the <b>RegisterWindowMessage</b> API call:
  </font></p>
  <p align=left bgcolor=Silver >
  <font color=Olive>'\\ Declaration to register custom messages<br></font>
  <font color=Blue>Private Declare Function</font> RegisterWindowMessage <font color=Blue>Lib</font> "user32" <font color=Blue>Alias</font> <br>
      "RegisterWindowMessageA"<font color=Blue> (ByVal</font> lpString <font color=Blue>As String) As Long</font>
  </font>
  </p>
  <p align=justify><font face=Arial> This API call takes an unique string and registers it as a defined windows message, returning a system wide unique identifier for that message
  as a result. Thereafter any call to <b>RegisterWindowMessage</b> in any application that specifies the same string will return the same unique message id.<br>
  Because this value is constant during each session it is safe to store it in a global variable to speed up execution thus:
  </font> </p>
  <p bgcolor=Silver >
  <font color=Blue>Public Const</font> MSG_CHANGENOTIFY = "MCL_PRINT_NOTIFY" <br>
  <br>
  <font color=Blue>Public Function</font> WM_MCL_CHANGENOTIFY() <font color=Blue>As Long</font> <br>
  <font color=Blue>Static</font> msg <font color=Blue>As Long</font> <br>
  <br>
  <font color=Blue>If</font> msg = 0 <font color=Blue>Then</font> <br>
     msg = RegisterWindowMessage(MSG_CHANGENOTIFY)<br>
  <font color=Blue>End If</font><br>
  <br>
  WM_MCL_CHANGENOTIFY = msg<br>
  <br>
  <font color=Blue>End Function</font><br>
  </p>
  <p align=justify><font face=Arial> <i>Since this message needs to be known to every application that is using it to communicate, it is a good idea to put this into a shared
  code module common to all projects.</i>
  </font></p>
  <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow"> 2. Creating windows to listen for these messages</font></p>
  <p align=justify><font face=Arial> To create a window in Visual Basic you usually use the form designer and add a new form to your project. However, since our communications
  window has no visible component nor interaction with the user, this is a bit excessive.<br>
   Instead we can use the <b>CreateWindowEx</b> API call to create a window solely for our communication:
  </font></p>
  <p bgcolor=Silver >
  <font color=Blue>Private Declare Function</font> CreateWindowEx <font color=Blue>Lib</font> "user32" <font color=Blue>Alias</font> "CreateWindowExA" <br>
     <font color=Blue>(ByVal</font> dwExStyle <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> lpClassName <font color=Blue>As String</font>, <font color=Olive>'\\ The window class, e.g. "STATIC","BUTTON" etc.</font><br>
     <font color=Blue>ByVal</font> lpWindowName <font color=Blue>As String</font>, <font color=Olive>'\\ The window's name (and caption if it has one)</font><br>
     <font color=Blue>ByVal</font> dwStyle <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> x <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> y <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> nWidth <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> nHeight <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> hWndParent <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> hMenu <font color=Blue>As Long</font>, <br>
     <font color=Blue>ByVal</font> hInstance <font color=Blue>As Long</font>, <br>
     lpParam <font color=Blue>As Any) As Long</font> <br>
  </p>
  <p align=justify><font face=Arial> If this call is successful, it returns an unique <b>window handle</b> which can be used to refer to that window. This can be used in <b>SendMessage</b>
  calls to send a message to it.
  </font></p>
  <p align=justify><font face=Arial> In a typical client/server communication you need to create one window for the client(s) and one window for the server. Again this can be done with a bit of code common to each application:
  </font></p>
  <p bgcolor=Silver >
  <font color=Blue>Public Const</font> WINDOWTITLE_CLIENT = "Merrion Computing IPC - Client" <br>
  <font color=Blue>Public Const</font> WINDOWTITLE_SERVER = "Merrion Computing IPC - Server"<br>
  <br>
  <font color=Blue>Public Function</font> CreateCommunicationWindow<font color=Blue>(ByVal</font> client <font color=Blue>As Boolean) As Long</font> <br>
 <br>
 <font color=Blue>Dim</font> hwndThis <font color=Blue>As Long</font> <br>
 <font color=Blue>Dim</font> sWindowTitle <font color=Blue>As String</font> <br>
 <br>
 <font color=Blue>If</font> client <font color=Blue>Then</font> <br>
    sWindowTitle = WINDOWTITLE_CLIENT <br>
 <font color=Blue>Else</font><br>
    sWindowTitle = WINDOWTITLE_SERVER <br>
 <font color=Blue>End If</font> <br>
 <br>
 hwndThis = CreateWindowEx(0, "STATIC", sWindowTitle, 0, 0, 0, 0, 0, 0, 0, <font color=Blue>App.hInstance, ByVal</font> 0&)<br>
 <br>
 CreateCommunicationWindow = hwndThis<br>
 <br>
  <font color=Blue>End Function</font> <br>
  </p>
  <p align=justify><font face=Arial> <i>Obviously for your own applications you should use different text for the WINDOWTITLE_CLIENT and WINDOWTITLE_SERVER than above to ensure that your window names are unique.</i>
  </font></p>
  <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow"> 3. Processing the custom messages</font></p>
  <p align=justify><font face=Arial> As it stands you have a custom message and have created a window to which you can send that message. However, as this message is entirely new to windows it does not do anything when it recieves it. To actually process the message you need to <a href="http://www.merrioncomputing.com/OnlineIssue2.htm">subclass</a> the window to intercept and react to the message yourself.<br>
   To subclass the window you create a procedure that processes windows messages and substitute this for the default message handling procedure of that window. Your procedure must have
 the same parameters and return type as the default window procedure:
  </font></p>
 <p bgcolor=Silver >
 <font color=Blue>Private Declare Function</font> CallWindowProc <font color=Blue>Lib</font>
  "user32" <font color=Blue>Alias</font> "CallWindowProcA" <font color=Blue>(ByVal</font> lpPrevWndFunc
  <font color=Blue>As Long, ByVal</font> hwnd <font color=Blue>As Long, ByVal</font> msg <font color=Blue>As Long, ByVal</font>
  wParam <font color=Blue>As Long, ByVal</font> lParam <font color=Blue>As Long) As Long</font>
</font><br>
 <br>
 <font color=Olive>
 '\\ --[VB_WindowProc]-----------------------<br>
 '\\ 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM); <br>
 '\\ Parameters: <br>
 '\\ hwnd - window handle receiving message <br>
 '\\ wMsg - The window message (WM_..etc.) <br>
 '\\ wParam - First message parameter <br>
 '\\ lParam - Second message parameter <br>
</font>
 <font color=Blue>Public Function</font> VB_WindowProc(<font color=Blue>ByVal</font> hwnd <font color=Blue>As Long, ByVal</font> wMsg <font color=Blue>As Long, ByVal</font> wParam <font color=Blue>As Long, ByVal</font> lParam <font color=Blue>As Long) As Long</font> <br>
  <br>
  <font color=Blue>If</font> wMsg = WM_MCL_CHANGENOTIFY <font color=Blue>Then</font> <br>
  <font color=Olive>   '\\Respond to the custom message here</font><br>
  <br>
  <font color=Blue>Else</font><br>
  <font color=Olive>   '\\Pass the message to the previous window procedure to handle it</font><br>
     VB_WindowProc = CallWindowProc(hOldProc, hwnd, wMsg, wParam, lParam)<br>
  <font color=Blue>End If</font><br>
  <br>
  <font color=Blue>End Function</font> <br>
 </p>
 <p align=justify><font face=Arial>You then need to inform Windows to substitute this procedure for the existing window procedure. To do this you call <b>SetWindowLong</b> to change the address
 of the procedure as stored in the <b>GWL_WINDPROC</b> index.
 </font></p>
 <p bgcolor=Silver >
  <font color=Blue>Public Const</font> GWL_WNDPROC = (-4) <br>
  <font color=Blue>Public Declare Function</font> SetWindowLongApi <font color=Blue>Lib</font> "user32" <font color=Blue>Alias</font> "SetWindowLongA"
 (<font color=Blue>ByVal</font> hwnd <font color=Blue>As Long, ByVal</font> nIndex <font color=Blue>As Long, ByVal</font> dwNewLong <font color=Blue>As Long) As Long</font> <br>
 <br>
 <font color=Olive>'\\ Use (after creating the window...)</font> <br>
 hOldProc = SetWindowLongApi(hwndThis, GWL_WNDPROC, <font color=Blue>AddressOf</font> VB_WindowProc) <br>
 </p>
 <p align=justify><font face=Arial> You keep the address of the previous window procedure address in <b>hOldProc</b> in order to pass on all the messages that you don't deal with for
 default processing. It is a good idea to set the window procedure back to this address before closing the window.
 </font></p>
 <p align=left><font face="Arial" style="BACKGROUND-COLOR: yellow"> 4. Sending the custom messages</font></p>
 <p align=justify><font face=Arial> There are two steps to sending the custom message to your server window: First you need to find the window handle of that window using the <b>FindWindowEx</b> API call then you need to send the message using the <b>SendMessage</b> API call.
 </font></p>
 <p bgcolor=Silver >
 <font color=Olive>'\\ Declarations</font> <br>
 <font color=Blue>Public Declare Function</font> SendMessageLong <font color=Blue>Lib</font> "user32" <font color=Blue>Alias</font> "SendMessageA"
 <font color=Blue>(ByVal</font> hwnd <font color=Blue>As Long, ByVal</font> wMsg <font color=Blue>As Long, ByVal</font> wParam
 <font color=Blue>As Long, ByVal</font> lParam <font color=Blue>As Long) As Long</font> <br>
 <font color=Blue>Public Declare Function</font> FindWindow <font color=Blue>Lib</font> "user32" <font color=Blue>Alias</font> "FindWindowA" (
 <font color=Blue>ByVal</font> lpClassName <font color=Blue>As String, ByVal</font> lpWindowName <font color=Blue>As String) As Long</font> <br>
 <br>
 <font color=Olive>'\\ use....</font> <br>
 <font color=Blue>Dim</font> hwndTarget <font color=Blue>As Long</font><br>
 <br>
 hwndTarget = FindWindow(vbNullString, WINDOWTITLE_SERVER)<br>
 <br>
 <font color=Blue>If</font> hwndTarget <> 0 <font color=Blue>Then</font> <br>
    <font color=Blue>Call</font> SendMessageLong(hwnd_Server, WM_MCL_CHANGENOTIFY, 0,0) <br>
 <font color=Blue>End If</font><br>
 </p>
 <p align=justify><font face=Arial color=Black> This will send the WM_MCL_CHANGENOTIFY message to the server window and return when it has been processed.
 </font></p>
 <!-- Source code to download -->
  <strong>Source Code</strong>
  <p align=justify>
  <font face=Arial color=Blue>The complete source code for these examples is available for download <a href="http://groups.yahoo.com/group/MerrionComputing/files/PrintWatchClient.zip
">here</a><br>
  You will be asked to register with <b>Yahoo!Groups</b> in order to access it.
  </font>
  </p>

