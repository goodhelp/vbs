'�����ֵ�
Dim Dict : Set Dict = CreateObject("Scripting.Dictionary")
 
'��Ӽ�ֵ��
Dict.Add "Key1", "Item1"
Dict.Add "Key2", "Item2"
Dict.Add "Key3", "Item3"
 
'�ֵ��м�ֵ������
WScript.Echo "�ֵ������м�ֵ������: " & Dict.Count '��һ���ű�����Ļ����ʾ�ı���Ϣ
 
WScript.Echo 
 
'���ָ�����Ƿ����
If Dict.Exists("Key1") Then
 WScript.Echo "Key1 ����!"
Else
 WScript.Echo "Key1 ������!"
End If
 
If Dict.Exists("Keyn") Then
 WScript.Echo "Keyn ����!"
Else
 WScript.Echo "Keyn ������!"
End If
 
WScript.Echo 
 
'�����ֵ�
Sub TraverseDict
 Dim DictKeys, DictItems, Counter
 DictKeys = Dict.Keys
 DictItems = Dict.Items 'Items����һ����������Itemֵ������
 For Counter = 0 To Dict.Count - 1 'Count����Dictionary�������Ŀ
 WScript.Echo "��: " & DictKeys(Counter) & "ֵ: " & DictItems(Counter)
 Next
End Sub
 
TraverseDict
 
WScript.Echo 
 
'��һ����ֵ���У��޸ļ����޸�ֵ
Dict.Key("Key2") = "Keyx"
Dict.Item("Key1") = "Itemx"
TraverseDict
 
WScript.Echo 
 
'ɾ��ָ����
Dict.Remove("Key3")
TraverseDict
 
WScript.Echo 
 
'ɾ��ȫ����
Dict.RemoveAll
WScript.Echo "�ֵ������м�ֵ������: " & Dict.Count