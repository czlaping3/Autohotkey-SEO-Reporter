


#Requires AutoHotkey v2.0+
#Include WinClipAPI.ahk
#Include WinClip.ahk
#Include AhkSoup.ahk


; Hot key to exit gracefully
^u::
{
  ExitApp
}

; Main Gui
App := Gui("+Resize")
AppBtn := App.Add("Button", "default", "&Go")
AppBtn.OnEvent("Click", Search)
AppMainEdit := App.Add("Edit", "WantTab W400 R15")
AppMainEdit.Value := "bakeries san antonio tx `r`nbirthday cakes san antonio tx `r`nwedding cakes san antonio tx`r`nbar mitzvah cakes san antonio tx `r`nquinceanera cake san antonio tx"
App.Show()
App.OnEvent("Close", OnClose)

OnClose(ThisGui) 
{
  ; The ExitApp command terminates the script entirely.
  ExitApp
}

Search(*)
{
  results_map := Map()
  ; Choose results file
  fl := "Search-Result.txt"
  searches := SplitContents(AppMainEdit.Value)
  For search in searches
  {
    ; Open Chrome
    OpenCHrome()
    Sleep 3000
    SendInput search 
    SendInput "{ENTER}"
    Sleep 3000
    ; Copies text
    CopyText()
    Sleep 500
    ; Get HTML from Clipboard
    html := GetHTML()
    ; Parse HTML
    rlt := ParseHTML(html)
    results_map[search] := rlt   
    Sleep 500 
    SendInput "^w"      
  }
  ; Write results to file
  WriteToFile(results_map,fl)

}

SplitContents(text)
{
  ; Get Edit Contents
  ;MsgBox  text   AppMainEdit.Value
  searches := StrSplit(text,"`n")  ; Split contents of main edit
  return searches
}

; Opens chrome
OpenChrome()
{
  Run "chrome.exe https://www.google.com"
}

; Retrieves HTML from clipboard
GetHTML()
{
  clpbrd := WinClip()
  html := clpbrd.GetHTML()
  return html
}

; Copies Text/HTML to clipboard
CopyText()
{
  WinActivate(WinGetTitle("A"))
  SendInput "^a"
  Sleep 500
  SendInput "^c"
  Sleep 500
}

; Parses HTML
ParseHTML(html)
{ 
  rlt := []
  ; Iniate AHKSoup instance
  page := AhkSoup()
  page.Open(html)
  elements := page.GetElementsByClassName("yuRUbf")
  ; Iterate through elements
  for element in elements
  {
    elementHTML := element.innerHTML
    ;MsgBox elementHTML
    element_tree := AhkSoup()
    element_tree.Open(elementHTML)
    title :=  element_tree.GetElementByTagName("h3")
    anchor :=  element_tree.GetElementByTagName("a")
    ;MsgBox title.Text "   " StrReplace(anchor.Text,title.Text)
    trt := RegExMatch(anchor.outerHTML, "(?:href\s*=\s*`")([^`"]+)(?:`")",&m_href,1)        ;`".*`"\b", &m_href,1)
    ;MsgBox m_href[1]
    ; Save results to result array
    rlt.Push([title.Text,m_href[1]])
  }
  return rlt
}

; Gets the length of the longest Google search result. Used in formatting report.
GetMaxKeyLength(m_map)
{
  max_length := 0
  for key, values in m_map
  {
    if StrLen(key) > max_length
    {
       max_length := StrLen(key)
    }  
  }  
  return max_length
}

; Left pads a string with spaces. Used to format report.
PadStringLeft(m_string,num)
{
  tmp_string := m_string
  loop num
  {
    if StrLen(tmp_string) < num
    {
      tmp_string := tmp_string " "
    }
    else
    {
      break 
    }
  }
  return tmp_string
}

; Writes results to file
WriteToFile(results_map,fl)
{
  results_map2 := Map()
  ; Populate results_map2 for report
  for key, values in results_map
  {
    line := key "`n"
    FileAppend line, fl
    for value in values
    {
      line := A_Index ". " value[1] "`n"
      FileAppend line, fl
      if not results_map2.Has(value[1])
      {
        results_map2[value[1]] := []
        results_map2[value[1]].Push([key, A_Index])
      }
      else
      {
        results_map2[value[1]].Push([key, A_Index])  
      } 
    }
    FileAppend "`n", fl    
  }
  ; Get longest Google result
  max_key_length := GetMaxKeyLength(results_map2)
  ; Write table refernce table
  for key, values in results_map
  {
    line := A_Index "`. " key "`n"
    FileAppend line, fl
  }
   FileAppend "`n", fl

  ; Create results table header
  line := "----------------------------------------------------------------------------------------------------------------------------------------------------------- `n"
  FileAppend line, fl
  line := PadStringLeft("Name", max_key_length)
  for key, values in results_map
  {
    line := line "`t"  A_Index   
  }
  line := line "`n"
  FileAppend line, fl
  line := "----------------------------------------------------------------------------------------------------------------------------------------------------------- `n"
  FileAppend line, fl

  ; Create results table 
  for key2, values2 in results_map2
  {
    line := PadStringLeft(key2, max_key_length) "`t" 
    for key, values in results_map
    {
      key_found := 0
      for value in values2
      {
        if value[1] = key
        {
          line := line value[2] "`t"
          key_found := 1
        }     
      }
      if key_found = 0  
      {
        line := line "`.`t"
      }  
    }
    line := line "`n" 
    FileAppend line, fl
  }
}
