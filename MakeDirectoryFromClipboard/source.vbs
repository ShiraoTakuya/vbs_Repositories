Dim strClip
strClip = CreateObject("htmlfile").ParentWindow.ClipboardData.Getdata("text")
arClip = Split(strClip, vbCrLf)

Set oFSO = CreateObject("Scripting.FileSystemObject")
For Each str In arClip
  oFSO.CreateFolder str
Next