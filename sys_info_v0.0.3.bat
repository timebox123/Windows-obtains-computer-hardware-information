' 2>nul 3>nul&cls&@echo off
'&rem ��ȡ����ϵͳ��Ӳ��������Ϣ
'&title ��ȡ����ϵͳ��Ӳ��������Ϣ
'&cscript /nologo -e:vbscript "%~fs0">"%USERPROFILE%\Desktop\computer-info.txt"
'&exit
   
On Error Resume Next
' Set fso=CreateObject("Scripting.Filesystemobject")
' Set ws=CreateObject("WScript.Shell")
Set wmi=GetObject("winmgmts:\\.\root\cimv2")
    
WSH.echo "---------------ϵͳ-------------"
Set query=wmi.ExecQuery("Select * from Win32_ComputerSystem")
For each item in query
    WSH.echo "��ǰ�û�=" & item.UserName
    WSH.echo "������=" & item.Workgroup
    WSH.echo "��=" & item.Domain
    WSH.echo "�������=" & item.Name
    WSH.echo "ϵͳ����=" & item.SystemType
Next
    
Set query=wmi.ExecQuery("Select * from Win32_OperatingSystem")
For each item in query
    WSH.echo "ϵͳ=" & item.Caption & "[" & item.Version & "]"
    WSH.echo "��ʼ��װ����=" & item.InstallDate
    visiblemem=item.TotalVisibleMemorySize
    virtualmem=item.TotalVirtualMemorySize
Next
    
Set query=wmi.ExecQuery("Select * from Win32_ComputerSystemProduct")
For each item in query
    WSH.echo "������=" & item.Vendor
    WSH.echo "�ͺ�=" & item.Name
    WSH.echo ""
Next
    
WSH.echo "---------------����BIOS-------------"
Set query=wmi.ExecQuery("Select * from Win32_BaseBoard")
For each item in query
    WSH.echo "������=" & item.Manufacturer
    WSH.echo "���к�=" & item.SerialNumber
    WSH.echo ""
Next

Set query=wmi.ExecQuery("Select * from Win32_BIOS")
For each item in query
    WSH.echo "����=" & item.Name
    WSH.echo "bios������=" & item.Manufacturer
    WSH.echo "��������=" & item.ReleaseDate
    WSH.echo "�汾=" & item.SMBIOSBIOSVersion
    WSH.echo ""
Next
    
WSH.echo "---------------CPU-------------"
Set query=wmi.ExecQuery("Select * from WIN32_PROCESSOR")
For each item in query
    WSH.echo "���=" & item.DeviceID
    WSH.echo "����=" & item.Name
    WSH.echo "����=" & item.NumberOfCores
    WSH.echo "�߳�=" & item.NumberOfLogicalProcessors
    WSH.echo ""
Next
    
WSH.echo "---------------�ڴ�-------------"
WSH.echo "�������ڴ�=" & FormatNumber(visiblemem/1048576,2,True) & " GB"
WSH.echo "�������ڴ�=" & FormatNumber(virtualmem/1048576,2,True) & " GB"
Set query=wmi.ExecQuery("Select * from Win32_PhysicalMemory")
For each item in query
    WSH.echo "���=" & item.Tag
    WSH.echo "����=" & FormatSize(item.Capacity)
    WSH.echo "��Ƶ=" & item.Speed
    WSH.echo "������=" & item.Manufacturer
    WSH.echo ""
Next
    
WSH.echo "--------------Ӳ��-------------"
Set query=wmi.ExecQuery("Select * from Win32_DiskDrive")
For each item in query
    WSH.echo "����=" & item.Caption
    WSH.echo "�ӿ�=" & item.InterfaceType
    WSH.echo "����=" & FormatSize(item.Size)
    WSH.echo "������=" & item.Partitions
    WSH.echo ""
Next
    
Set query=wmi.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3 or DriveType=2")
For each item in query
    WSH.echo item.Caption & Chr(9) & item.FileSystem & Chr(9) & FormatSize(item.Size) & Chr(9) & FormatSize(item.FreeSpace)
Next
    
WSH.echo "--------------����-------------"
Set query=wmi.ExecQuery("Select * from Win32_NetworkAdapter Where NetConnectionID !=null and not Name like '%Virtual%'")
For each item in query
    WSH.echo "����=" & item.Name
    WSH.echo "������=" & item.NetConnectionID
    WSH.echo "MAC=" & item.MACAddress
    Set query2=wmi.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where Index=" & item.Index)
    For each item2 in query2
        If typeName(item2.IPAddress) <> "Null" Then
            WSH.echo "IP=" & item2.IPAddress(0)
        End If
    Next
    WSH.echo ""
Next
    
WSH.echo "--------------��ʾ-------------"
Set query=wmi.ExecQuery("Select * from Win32_VideoController")
For each item in query
    WSH.echo "����=" & item.Name
    WSH.echo "�Դ�=" & FormatSize(Abs(item.AdapterRAM))
    WSH.echo "��ǰˢ����=" & item.CurrentRefreshRate
    WSH.echo "ˮƽ�ֱ���=" & item.CurrentHorizontalResolution
    WSH.echo "��ֱ�ֱ���=" & item.CurrentVerticalResolution
    WSH.echo ""
Next
    
WSH.echo "--------------����-------------"
Set query=wmi.ExecQuery("Select * from WIN32_SoundDevice")
For each item in query
    WSH.echo item.Name
    WSH.echo ""
Next
    
WSH.echo "--------------��ӡ��-------------"
Set query=wmi.ExecQuery("Select * from Win32_Printer")
For each item in query
    If item.Default =True Then
        WSH.echo item.Name & "(Ĭ��)"
    Else
        WSH.echo item.Name
    End If
    WSH.echo ""
Next
    
Function FormatSize(byVal t)
    If t >= 1099511627776 Then
        FormatSize = FormatNumber(t/1099511627776, 2, true) & " TB"
    ElseIf t >= 1073741824 Then
        FormatSize = FormatNumber(t/1073741824, 2, true) & " GB"
    ElseIf t >= 1048576 Then
        FormatSize = FormatNumber(t/1048576, 2, true) & " MB"
    ElseIf t >= 1024 Then
        FormatSize = FormatNumber(t/1024, 2, true) & " KB"
    Else
        FormatSize = t & " B"    
    End If
End Function