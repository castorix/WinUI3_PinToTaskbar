# WinUI3_PinToTaskbar

Test Pin/Unpin to taskbar with IPinManagerInterop, IPinnedList3, [IStartMenuPinnedList](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl/nn-shobjidl-istartmenupinnedlist)

 Tested on Windows 10 22H2, Windows App SDK 1.6.241114003

 From https://github.com/microsoft/WindowsAppSDK/issues/3917#issuecomment-2696821712, COM interfaces do not work on Windows 11

 So I added another method with a [WH_GETMESSAGE Hook](https://learn.microsoft.com/en-us/windows/win32/winmsg/about-hooks#wh_getmessage)
 to execute "Pin to Taskbar" context menu in Explorer address space.
 
 The original idea comes from [SysPin](https://www.technosys.net/products/utils/pintotaskbar), which calls WriteProcessMemory

 ![image](https://github.com/user-attachments/assets/4f20fc87-ff03-4603-b42f-b71191488744)

