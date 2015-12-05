/*
	Name:		folderJump.ahk
    Based on dirMenu.ahk by Robert Ryan
  
  Author: busywiththingsalreadyforgotten
          busywiththingsalreadyforgotten@gmail.com

	Function:
		See the About box

  To do:
    - Add tree view support (for those pesky browse for folder dialogues)
    - "Add this path" support for standard windows explorer
    - Done. "Add this path" support for standard dialogs
    - Add error checking throughout script
    - add check that path exists before adding to folderjump.txt
    - add check that path exists before navigating ot a folder
    - change open folder in new window to open in directory opus
    - test whta happen when multiple diretory opus windows are open
*/

#NoEnv
#SingleInstance force
SetBatchLines -1
ListLines Off
SendMode Input
SetWorkingDir %A_ScriptDir%

versionNumber :=001
; ---------------------------------------------------------------------------------------------------------------------------
; Settings
; ---------------------------------------------------------------------------------------------------------------------------
openPathinDopus :=1 ; change this to 0 if you aren't using directory opus
dopusExecutable :="C:\Program Files\GPSoftware\Directory Opus\Dopus.exe"

; ---------------------------------------------------------------------------------------------------------------------------
; Autoexecute
; ---------------------------------------------------------------------------------------------------------------------------
MakeTrayMenu()
MakeGUI()
MakeCallTable()
LoadListView()
return

; ---------------------------------------------------------------------------------------------------------------------------
; Labels and hotkeys
; ---------------------------------------------------------------------------------------------------------------------------
GuiCall:
  Call[A_GuiControl].()
return

MenuCall:
  Call[A_ThisMenuItem].()
return

OpenFavorite:
  LV_GetText(Path, A_ThisMenuItemPos, 2)
  if (A_ThisHotkey = "+MButton") 
  {
    NewWindow(Path)
  }
  else 
  {
    WinWaitActive ahk_id %WinId%
    Call[WinGetClass("A")].(Path)
  }
return

GuiClose:
  Cancel()
return

+MButton::
  UpdateMenuItemStates()
  Menu Favorites, Show
return

#If Call.HasKey(ClassMouseOver())
#a::
MButton::
  WinId := WinMouseOver()
  WinActivate ahk_id %WinId%
  UpdateMenuItemStates()
  Menu Favorites, Show
return

; ---------------------------------------------------------------------------------------------------------------------------
; Initialisation
; ---------------------------------------------------------------------------------------------------------------------------
MakeCallTable()
{
  global Call := []
  
  Call["MyList"] := Func("UpdateButtons")
  Call["&Add"] := Func("Add")
  Call["&Remove"] := Func("Remove")
  Call["&Modify"] := Func("Modify")
  Call["&Separator"] := Func("Separator")
  Call["Move &Up"] := Func("MoveUp")
  Call["Move &Down"] := Func("MoveDown")
  Call["&OK"] := Func("OK")
  Call["&Cancel"] := Func("Cancel")
  Call["Re&vert"] := Func("LoadListView")
  Call["About"] := Func("About")
  Call["Edit Custom Menu"] := Func("ShowGUI")
  Call["Edit This Menu"] := Func("ShowGUI")
  Call["Add This Path"] := Func("AddThisPath")
  
  Call["Progman"] := Func("NewWindow")
  Call["CabinetWClass"] := Func("Explorer")
  Call["#32770"] := Func("Dialog")
  Call["ConsoleWindowClass"] := Func("Console")

  Call["dopus.lister"] := Func("DirectoryOpus")
}

LoadListView()
{
	if FileExist("folderJumpPaths.txt")
  {
		ReadFile()
	}
  else
  {
		FirstRun()
    GuiControl Disable, Re&vert
  }
}

; If there is no settings file, set up a a default custom menu
FirstRun()
{
  LV_Delete()
  LV_Add("Select Focus", "C:\", "C:\")
  LV_Add("", "My Documents", A_MyDocuments)
  LV_Add("", "---------------", "---------------")
  LV_Add("", "Windows", A_WinDir)
  LV_Add("", "Program Files", A_ProgramFiles)
  LV_ModifyCol(1, "AutoHdr")
  LV_ModifyCol(2, "AutoHdr")
  UpdateMenu()
}

; Read the ListView entries from the settings file
ReadFile()
{
  LV_Delete()
  Loop Read, folderJumpPaths.txt
  {
    StringSplit Item, A_LoopReadLine, %A_Tab%
	  LV_Add("", Item1, Item2)
  }
  LV_Modify(1, "Select Focus")
  LV_ModifyCol(1, "AutoHdr")
  LV_ModifyCol(2, "AutoHdr")
  UpdateMenu()
}

; Update the custom menu to match the ListView
UpdateMenu()
{
  Menu Favorites, Add
  Menu Favorites, DeleteAll
  Loop % LV_GetCount() 
  {
    LV_GetText(Name, A_Index, 1)
    if (Name = "---------------")
    {
      Menu Favorites, Add
    }
    else
    {
      Menu Favorites, Add, %Name%, OpenFavorite
    }
  }
  Menu Favorites, Add
  Menu Favorites, Add, Add This Path, MenuCall
  Menu Favorites, Add, Edit This Menu, MenuCall
}

; Enable / Disable menu items according to which dialog is visible
UpdateMenuItemStates()
{
  if ( isWindow_Dopus() )
  {
    Menu, Favorites, Enable, Add This Path
  }
  else if ( isWindow_Explorer() )
  {
    Menu, Favorites, Enable, Add This Path
  }
  else if ( isWindow_StandardFileDialog() )
  {
    Menu, Favorites, Enable, Add This Path
  }
  else if ( isWindow_SelectFolderDialog() )
  {
    Menu, Favorites, Disable, Add This Path
  }
  else 
  {
    ; unknown dialog
    Menu, Favorites, Disable, Add This Path
  }
}

; ---------------------------------------------------------------------------------------------------------------------------
; Window Test Methods
; ---------------------------------------------------------------------------------------------------------------------------
isWindow_Dopus()
{
  activeClass := ClassMouseOver()
  if ( activeClass == "dopus.lister")
  {
    return 1
  }
  return 0
}

isWindow_Explorer()
{
  activeClass := ClassMouseOver()
  if ( activeClass == "CabinetWClass")
  {
    return 1
  }
  return 0
}

isWindow_StandardFileDialog()
{
  activeClass := ClassMouseOver()
  if ( activeClass == "#32770")
  {
    if ControlExist("ToolbarWindow323", "A") 
    {
      return 1
    }
  }
  return 0
}

isWindow_SelectFolderDialog()
{
  activeClass := ClassMouseOver()
  if ( activeClass == "#32770")
  {
    if not ControlExist("ToolbarWindow323", "A") 
    {
      return 1
    }
  }
  return 0
}

; ---------------------------------------------------------------------------------------------------------------------------
; Menu Methods
; ---------------------------------------------------------------------------------------------------------------------------
NewWindow(Path)
{
  global
  if (openPathinDopus == 0)
  {
    Run Explorer.exe /n`,%Path%
  }
  else 
  {
    SetTitleMatchMode, 2
    IfWinExist ahk_class dopus.lister
    {
      WinActivate
    }
    else
    {
      Run, %dopusExecutable%
    }
    WinWaitActive, ahk_class dopus.lister
    DirectoryOpus(Path)
  }
}

Console(Path)
{
  Send cd /d %Path%{Enter}
}

Dialog(Path)
{
  WinGetTitle Title, A
  ;if not RegExMatch(Title, "i)save|open") {
  ;    NewWindow(Path)
  ;    return
  ;}
  ControlFocus Edit1, A
  ControlGetText OldText, Edit1, A
  ControlSetText Edit1, %Path%, A
  ControlSend Edit1, {Enter}, A
  ControlSetText Edit1, %OldText%, A
}

Explorer(Path)
{
  if RegExMatch(A_OSVersion, "WIN_VISTA|WIN_7")
  {
    Win7_Explorer(Path)
  }
  else
  {
    XP_Explorer(Path)
  }
}

Win7_Explorer(Path)
{
  if ControlExist("ToolbarWindow323", "A") 
  {
    ControlSetText, Path, ToolbarWindow323, A
  }
}

XP_Explorer(Path)
{
  if not ControlExist("Edit1", "A") 
  {
    PostMessage 0x111, 41477, 0, , A ; Show Address Bar
    while not ControlExist("Edit1", "A")
    {
      Sleep 0
    }
    PostMessage 0x111, 41477, 0, , A ; Hide Address Bar
  }
  ControlFocus Edit1, A
  ControlSetText Edit1, %Path%, A
  ControlSend Edit1, {Enter}, A
}

DirectoryOpus(Path)
{
  WinGetTitle Title, A
  ControlGetFocus, FileListControl, %Title%

  if (ErrorLevel)
  {
    showMessage("The target window doesn't exist or none of its controls has input focus")
  }
  else
  {
    ; get target path Edit control name
    /*
        There can be multiple filedisplay control (some of which are hidden)
        However they remain in numerical order if there is more than one present
        There path control edits are only ever name Edit1 and Edit2 (when present)
    */
    StringReplace, FileDisplayControlNumber, FileListControl, dopus.filedisplay,, All

    ; search through all controls and check if there is a visible filedisplay control with a smaller number
    PathControlNumber := 1
    WinGet, ActiveControlList, ControlList, A
    Loop, Parse, ActiveControlList, `n
    {
      controlname := SubStr(A_LoopField, 1 ,17)
      if (controlname == "dopus.filedisplay") && (A_LoopField != FileListControl)
      {
        controlname2 := SubStr(A_LoopField, 1 ,26)
        if (controlname2 != "dopus.filedisplaycontainer" )
        {
          ControlGet, ControlVisible, Visible,, %A_LoopField%
          if(ControlVisible == 1)
          {
            StringReplace, controlNumber, A_LoopField, dopus.filedisplay,, All
            if (controlNumber < FileDisplayControlNumber)
            {
              PathControlNumber += 1
            }
          }
        }
      }
    }
    ControlSetText Edit%PathControlNumber%, %Path%, A
    ControlSend Edit%PathControlNumber%, {Enter}, A
  }
}

; Add current path (in window or dialog) to folderJump folders
AddThisPath()
{
  ; check if in directory DirectoryOpus
  activeClass := ClassMouseOver() 
  if ( isWindow_Dopus() )
  {
    ; Dopus lister
    WinWaitActive, ahk_class %activeClass%
    WinGetTitle Title, A
    ControlGetFocus, FileListControl, %Title%
    if (ErrorLevel)
    {
      showMessage("The target window doesn't exist or none of its controls has input focus.")
    } 
    else 
    {
      ; get target path Edit control name
      /*
          There can be multiple filedisplay control (some of which are hidden)
          However they remain in numerical order if there is more than one present
          There path control edits are only ever name Edit1 and Edit2 (when present)
      */
      StringReplace, FileDisplayControlNumber, FileListControl, dopus.filedisplay,, All

      ; search through all controls and check if there is a visible filedisplay control with a smaller number
      PathControlNumber := 1
      WinGet, ActiveControlList, ControlList, A
      Loop, Parse, ActiveControlList, `n
      {
        controlname := SubStr(A_LoopField, 1 ,17)
        if (controlname == "dopus.filedisplay") && (A_LoopField != FileListControl)
        {
          controlname2 := SubStr(A_LoopField, 1 ,26)
          if (controlname2 != "dopus.filedisplaycontainer" )
          {
            ControlGet, ControlVisible, Visible,, %A_LoopField%
            if(ControlVisible == 1)
            {
              StringReplace, controlNumber, A_LoopField, dopus.filedisplay,, All
              if (controlNumber < FileDisplayControlNumber)
              {
                PathControlNumber += 1
              }
            }
          }
        }
      }

      ControlGetText, PathText, Edit%PathControlNumber%, A
      Gui +LastFound +OwnDialogs +AlwaysOnTop
      InputBox Name, Folder Name, Please Enter a name, , 250, 130
      if (ErrorLevel)
      {
        return
      }
      FileAppend % Name . A_Tab . PathText "`n", folderJumpPaths.txt
      ReadFile()
    }
  }
  else if ( isWindow_Explorer() )
  {
    ; Windows Explorer
    ControlGetText, PathText, ToolbarWindow323, A
    if (ErrorLevel) 
    {
      showMessage("AddThisPath Explorer ControlGetText error")
    } 
    else
    {
      Gui +LastFound +OwnDialogs +AlwaysOnTop
      InputBox Name, Folder Name, Please Enter a name, , 250, 130
      if (ErrorLevel) 
      {
        return
      }
      ; remove "Address: " prefix
      StringReplace, PathText, PathText,Address: ,, All
      FileAppend % Name . A_Tab . PathText "`n", folderJumpPaths.txt
      ReadFile()
    }

  }
  else if ( isWindow_StandardFileDialog() )
  {
    ; standard full file dialog
    WinWaitActive, ahk_class %activeClass%
    WinGetTitle Title, A
    ControlGetText, PathText, ToolbarWindow324, A
    if (ErrorLevel) 
    {
      showMessage("AddThisPath ControlGetText error")
    } 
    else
    {
      Gui +LastFound +OwnDialogs +AlwaysOnTop
      InputBox Name, Folder Name, Please Enter a name, , 250, 130
      if (ErrorLevel) 
      {
        return
      }
      ; remove "Address: " prefix
      StringReplace, PathText, PathText,Address: ,, All
      FileAppend % Name . A_Tab . PathText "`n", folderJumpPaths.txt
      ReadFile()
    }
  }
  else if ( isWindow_SelectFolderDialog() )
  {
    ; folder select dialog
    showMessage("Sorry, this function is not yet supported")
  } 
  else
  {
    showMessage("Unrecognised dialog class: %activeClass%")
  }
}

; Folder jump message box
showMessage(msgText)
{
  MsgBox, 262144, Folder Jump, %msgText%
}

; ---------------------------------------------------------------------------------------------------------------------------
; Edit Paths Dialog Methods
; ---------------------------------------------------------------------------------------------------------------------------
; Add a new entry to the ListView
Add()
{
  Gui +OwnDialogs
  
  FileSelectFolder Path, *C:\
  if (Path = "")
  {
    return
  }

  InputBox Name, Menu Name, Please Enter a name for the new entry:, , 250, 130
  if (ErrorLevel)
  {
    return
  }
  
  LV_Insert(LV_GetCount() ? LV_GetNext() : 1, "Select Focus", Name, Path)
  LV_ModifyCol(1, "AutoHdr")
  LV_ModifyCol(2, "AutoHdr")
  GuiControl Enable, Re&vert
}

; Remove an entry from the ListView
Remove()
{
  LV_Delete(LV_GetNext())
  GuiControl Enable, Re&vert
}

; Modify an existing entry in th ListView
Modify()
{
  Gui +OwnDialogs
  SelectedRow := LV_GetNext()
  LV_GetText(Name, SelectedRow, 1)
  LV_GetText(Path, SelectedRow, 2)
  
  FileSelectFolder NewPath, *%Path%
  if (NewPath = "")
  {
    return
  }
  
  InputBox NewName, Menu Name
         , Enter a name for the entry:, , 250, 130, , , , , %Name%
  if (ErrorLevel)
  { 
    return
  }

  LV_Modify(SelectedRow, "", NewName, NewPath)
  LV_ModifyCol(1, "AutoHdr")
  LV_ModifyCol(2, "AutoHdr")    
  GuiControl Enable, Re&vert
}

; Insert a separator line into the ListView
Separator()
{
  LV_Insert(LV_GetCount() 
          ? LV_GetNext() 
          : 1
          , "Select Focus", "---------------", "---------------")
  GuiControl Enable, Re&vert
}

; Move an entry down in the ListView
; LV_Modify is used to avoid flickering
MoveDown()
{
  SelectedRow := LV_GetNext()
  LV_GetText(ThisName, SelectedRow, 1)
  LV_GetText(ThisPath, SelectedRow, 2)
  
  LV_GetText(NextName, SelectedRow + 1, 1)
  LV_GetText(NextPath, SelectedRow + 1, 2)
  
  LV_Modify(SelectedRow, "", NextName, NextPath)
  LV_Modify(SelectedRow + 1, "Select Focus", ThisName, ThisPath)
  GuiControl Enable, Re&vert
}

; Move an entry up in the ListView
; LV_Modify is used to avoid flickering
MoveUp()
{
  SelectedRow := LV_GetNext()
  LV_GetText(ThisName, SelectedRow, 1)
  LV_GetText(ThisPath, SelectedRow, 2)
  
  LV_GetText(PriorName, SelectedRow - 1, 1)
  LV_GetText(PriorPath, SelectedRow - 1, 2)
  
  LV_Modify(SelectedRow, "", PriorName, PriorPath)
  LV_Modify(SelectedRow - 1, "Select Focus", ThisName, ThisPath)
  GuiControl Enable, Re&vert
}

; Save the ListView entries to the settings file
OK()
{
  Gui Cancel    
  FileDelete folderJumpPaths.txt
  
  Loop % LV_GetCount() 
  {
    LV_GetText(Name, A_Index, 1)
    LV_GetText(Path, A_Index, 2)
    FileAppend % Name . A_Tab . Path "`n", folderJumpPaths.txt
  }
  FileSetAttrib +H, folderJumpPaths.txt
  GuiControl Disable, Re&vert
  UpdateMenu()
}

Cancel()
{
  Gui Cancel
  LoadListView()
}

; This is called anytime the listview changes
UpdateButtons()
{
  Critical

  TotalNumberOfRows := LV_GetCount()
  
  ; Make sure there is always one selected row
  SelectedRow := LV_GetNext(0, "Focused")
  LV_Modify(SelectedRow, "Select")

  GuiControl % (TotalNumberOfRows = 0) ? "Disable" : "Enable", &Remove
  GuiControl % (SelectedRow <= 1) ? "Disable" : "Enable", Move &Up
  GuiControl % (SelectedRow = TotalNumberOfRows) ? "Disable" : "Enable", Move &Down
}

; ---------------------------------------------------------------------------------------------------------------------------
; Functions
; ---------------------------------------------------------------------------------------------------------------------------
WinMouseOver()
{
  MouseGetPos, , , WinId
  return WinId
}

WinGetClass(WinTitle = "", WinText = "", ExTitle = "", ExText = "")
{
  WinGetClass out, % WinTitle, % WinText, % ExTitle, % ExText
  return out
}

ClassMouseOver()
{
  return WinGetClass("ahk_id" WinMouseOver())
}

; Determine if a particular control exists in a particular window
ControlExist(Cntrl = "", WTitle = "", WText = "", ExTitle = "", ExText = "")
{
  ControlGet out, Enabled,, % Cntrl, % WTitle, % WText, % ExTitle, % ExText
  return not ErrorLevel
}

; ---------------------------------------------------------------------------------------------------------------------------
; GUI
; ---------------------------------------------------------------------------------------------------------------------------
About()
{
  global
  MsgBox, , About FolderJump,
  ( LTrim
      FolderJump gives you easy access to your favorite folders. This script is based on DirMenu.ahk by Robert Ryan. However it has been modified specifically to work with Directory Opus 11 and Windows 10 and there are a few extra features added to boot!

      Clicking middle mouse button (or pressing win + a) while hovering over certain window types will bring up a custom menu of your favorite folders. Upon selecting a favorite, the script will instantly switch to that folder within the active window. 
      
      Holding down the Shift key while clicking the middle mouse button will bring up the menu regardless of which window the mouse is over. The folder in this case will be shown in a new Explorer window.

      By default, The following window types are supported:
      - Standard file-open or file-save dialogs
      - Explorer windows
      - Console (command prompt) windows
      - The Desktop
      - Directory Opus 11

      A GUI is also included to reorder or change the folder paths 

      Version: %versionNumber%

      busywiththingsalreadyforgotten@gmail.com
  )
}

ShowGUI()
{
  Gui Show, , Folder Jump Folders
}

MakeGUI()
{
  global
  
  Gui , Add, ListView
      , xm w350 h240 Count15 -Multi NoSortHdr AltSubmit vMyList gGuiCall
      , Name|Path
  Gui, Add, Button, x+10 w75 r1 gGuiCall, &Add
  Gui, Add, Button, w75 r1 gGuiCall, &Remove
  Gui, Add, Button, W75 r1 gGuiCall, &Modify
  Gui, Add, Button, w75 r1 gGuiCall, &Separator
  Gui, Add, Button, w75 r1 gGuiCall, Move &Up
  Gui, Add, Button, w75 r1 gGuiCall, Move &Down
  Gui, Add, Button, xm+85 w75 r1 gGuiCall Default, &OK
  Gui, Add, Button, x+20 w75 r1 gGuiCall, &Cancel
  Gui, Add, Button, x+20 w75 r1 gGuiCall, Re&vert
}

MakeTrayMenu()
{
	Menu Default Menu, Standard
	Menu Tray, NoStandard
	Menu Tray, Add, About, MenuCall
	Menu Tray, Add
	Menu Tray, Add, Default Menu, :Default Menu
  Menu Tray, Add
	Menu Tray, Add, Edit Custom Menu, MenuCall
	Menu Tray, Default, Edit Custom Menu
  Menu, Tray, Icon, folderJump.ico 
}