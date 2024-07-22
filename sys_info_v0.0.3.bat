' 2>nul 3>nul&cls&@echo off
'&rem 获取本机系统及硬件配置信息
'&title 获取本机系统及硬件配置信息
'&cscript /nologo -e:vbscript "%~fs0">"%USERPROFILE%\Desktop\computer-info.txt"
'&exit
   
On Error Resume Next
' Set fso=CreateObject("Scripting.Filesystemobject")
' Set ws=CreateObject("WScript.Shell")
Set wmi=GetObject("winmgmts:\\.\root\cimv2")
    
WSH.echo "---------------系统-------------"
Set query=wmi.ExecQuery("Select * from Win32_ComputerSystem")
For each item in query
    WSH.echo "当前用户=" & item.UserName
    WSH.echo "工作组=" & item.Workgroup
    WSH.echo "域=" & item.Domain
    WSH.echo "计算机名=" & item.Name
    WSH.echo "系统类型=" & item.SystemType
Next
    
Set query=wmi.ExecQuery("Select * from Win32_OperatingSystem")
For each item in query
    WSH.echo "系统=" & item.Caption & "[" & item.Version & "]"
    WSH.echo "初始安装日期=" & item.InstallDate
    visiblemem=item.TotalVisibleMemorySize
    virtualmem=item.TotalVirtualMemorySize
Next
    
Set query=wmi.ExecQuery("Select * from Win32_ComputerSystemProduct")
For each item in query
    WSH.echo "制造商=" & item.Vendor
    WSH.echo "型号=" & item.Name
    WSH.echo ""
Next
    
WSH.echo "---------------主板BIOS-------------"
Set query=wmi.ExecQuery("Select * from Win32_BaseBoard")
For each item in query
    WSH.echo "制造商=" & item.Manufacturer
    WSH.echo "序列号=" & item.SerialNumber
    WSH.echo ""
Next

Set query=wmi.ExecQuery("Select * from Win32_BIOS")
For each item in query
    WSH.echo "名称=" & item.Name
    WSH.echo "bios制造商=" & item.Manufacturer
    WSH.echo "发布日期=" & item.ReleaseDate
    WSH.echo "版本=" & item.SMBIOSBIOSVersion
    WSH.echo ""
Next
    
WSH.echo "---------------CPU-------------"
Set query=wmi.ExecQuery("Select * from WIN32_PROCESSOR")
For each item in query
    WSH.echo "序号=" & item.DeviceID
    WSH.echo "名称=" & item.Name
    WSH.echo "核心=" & item.NumberOfCores
    WSH.echo "线程=" & item.NumberOfLogicalProcessors
    WSH.echo ""
Next
    
WSH.echo "---------------内存-------------"
WSH.echo "总物理内存=" & FormatNumber(visiblemem/1048576,2,True) & " GB"
WSH.echo "总虚拟内存=" & FormatNumber(virtualmem/1048576,2,True) & " GB"
Set query=wmi.ExecQuery("Select * from Win32_PhysicalMemory")
For each item in query
    WSH.echo "序号=" & item.Tag
    WSH.echo "容量=" & FormatSize(item.Capacity)
    WSH.echo "主频=" & item.Speed
    WSH.echo "制造商=" & item.Manufacturer
    WSH.echo ""
Next
    
WSH.echo "--------------硬盘-------------"
Set query=wmi.ExecQuery("Select * from Win32_DiskDrive")
For each item in query
    WSH.echo "名称=" & item.Caption
    WSH.echo "接口=" & item.InterfaceType
    WSH.echo "容量=" & FormatSize(item.Size)
    WSH.echo "分区数=" & item.Partitions
    WSH.echo ""
Next
    
Set query=wmi.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3 or DriveType=2")
For each item in query
    WSH.echo item.Caption & Chr(9) & item.FileSystem & Chr(9) & FormatSize(item.Size) & Chr(9) & FormatSize(item.FreeSpace)
Next
    
WSH.echo "--------------网卡-------------"
Set query=wmi.ExecQuery("Select * from Win32_NetworkAdapter Where NetConnectionID !=null and not Name like '%Virtual%'")
For each item in query
    WSH.echo "名称=" & item.Name
    WSH.echo "连接名=" & item.NetConnectionID
    WSH.echo "MAC=" & item.MACAddress
    Set query2=wmi.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where Index=" & item.Index)
    For each item2 in query2
        If typeName(item2.IPAddress) <> "Null" Then
            WSH.echo "IP=" & item2.IPAddress(0)
        End If
    Next
    WSH.echo ""
Next
    
WSH.echo "--------------显示-------------"
Set query=wmi.ExecQuery("Select * from Win32_VideoController")
For each item in query
    WSH.echo "名称=" & item.Name
    WSH.echo "显存=" & FormatSize(Abs(item.AdapterRAM))
    WSH.echo "当前刷新率=" & item.CurrentRefreshRate
    WSH.echo "水平分辨率=" & item.CurrentHorizontalResolution
    WSH.echo "垂直分辨率=" & item.CurrentVerticalResolution
    WSH.echo ""
Next
    
WSH.echo "--------------声卡-------------"
Set query=wmi.ExecQuery("Select * from WIN32_SoundDevice")
For each item in query
    WSH.echo item.Name
    WSH.echo ""
Next
    
WSH.echo "--------------打印机-------------"
Set query=wmi.ExecQuery("Select * from Win32_Printer")
For each item in query
    If item.Default =True Then
        WSH.echo item.Name & "(默认)"
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