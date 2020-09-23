<div align="center">

## Clean & Quick: EnumerateWindows


</div>

### Description

This code will enumerate all windows on the desktop including child windows, and children of children. There are several other entries here on PSC that perform a similar task, but they add functionality and interfaces that I didn't need... And so this is a more straightforward barebones approach; lighter & cleaner without worrying about treeviews or complicated code- A few API declarations, about a dozen lines of code and you're good to go.
 
### More Info
 
Just paste the code into a .bas module and call the EnumerateAllWindows function. In the EnumerateChildren function there is a place where you can do whatever you need to do with each window handle (right now the function prints to the immediate window for demo purposes).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Heydman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-heydman.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-heydman-clean-quick-enumeratewindows__1-13762/archive/master.zip)

### API Declarations

```
'API Declarations
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
```


### Source Code

```
'Call this function to begin the process of getting every window on the desktop
Public Sub EnumerateAllWindows()
Dim hWndDesktop As Long
 hWndDesktop = GetDesktopWindow()
 EnumerateChildren hWndDesktop
End Sub
Private Sub EnumerateChildren(hWndParent As Long)
Dim hWndChild As Long
 'Get the first child of hWndParent
 hWndChild = GetWindow(hWndParent, GW_CHILD Or GW_HWNDFIRST)
 Do While hWndChild <> 0
  ' At this point, hWndChild contains a child window handle of hWndParent.
  ' You could use GetWindowText here, for instance, to retrieve the title of the window.
  Debug.Print hWndParent, hWndChild
  'Now get any children for hWndChild
  EnumerateChildren hWndChild
  'And move on to the next window
  hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
 Loop
End Sub
```

