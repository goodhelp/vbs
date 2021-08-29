'建立字典
Dim Dict : Set Dict = CreateObject("Scripting.Dictionary")
 
'添加键值对
Dict.Add "Key1", "Item1"
Dict.Add "Key2", "Item2"
Dict.Add "Key3", "Item3"
 
'字典中键值对数量
WScript.Echo "字典中现有键值对数量: " & Dict.Count '让一个脚本在屏幕上显示文本信息
 
WScript.Echo 
 
'检查指定键是否存在
If Dict.Exists("Key1") Then
 WScript.Echo "Key1 存在!"
Else
 WScript.Echo "Key1 不存在!"
End If
 
If Dict.Exists("Keyn") Then
 WScript.Echo "Keyn 存在!"
Else
 WScript.Echo "Keyn 不存在!"
End If
 
WScript.Echo 
 
'遍历字典
Sub TraverseDict
 Dim DictKeys, DictItems, Counter
 DictKeys = Dict.Keys
 DictItems = Dict.Items 'Items返回一个包含所有Item值的数组
 For Counter = 0 To Dict.Count - 1 'Count返回Dictionary对象键数目
 WScript.Echo "键: " & DictKeys(Counter) & "值: " & DictItems(Counter)
 Next
End Sub
 
TraverseDict
 
WScript.Echo 
 
'在一个键值对中，修改键或修改值
Dict.Key("Key2") = "Keyx"
Dict.Item("Key1") = "Itemx"
TraverseDict
 
WScript.Echo 
 
'删除指定键
Dict.Remove("Key3")
TraverseDict
 
WScript.Echo 
 
'删除全部键
Dict.RemoveAll
WScript.Echo "字典中现有键值对数量: " & Dict.Count