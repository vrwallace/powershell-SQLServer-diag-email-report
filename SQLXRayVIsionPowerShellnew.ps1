####################################
# Program: SQLXRayVIsionPowerShell.ps1
# By: Von Wallace
# To run add the following to the login script
# powershell.exe �Noninteractive �Noprofile �Command "C:\support\SQLXRayVIsionPowerShell.ps1"
# Program queries SQL server for Info
###################################


$Version = "1.03ps"
#Install-Module -Name SqlServer
Import-Module SqlServer
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$smtpserver = "smtp.office365.com"
$smtpport = "587"
$smtpfrom = "someone@somewhere.net"
    
$smtpto = "someone@somewhere.net"
    
$sendusername = "someone@somewhere.net"
$sendpassword =  "emailpassword"

$strcomputer = "localhost\SQLEXPRESS"
$user = "dbuser"
$password = "dbpassword"


try {

    # I am defining website url in a variable
    $url = "http://checkip.dyndns.com" 
    # Creating a new .Net Object names a System.Net.Webclient
    $webclient = New-Object System.Net.WebClient
    # In this new webdownlader object we are telling $webclient to download the
    # url $url 
    $Ip = $webclient.DownloadString($url)
    # Just a simple text manuplation to get the ipadress form downloaded URL
    # If you want to know what it contain try to see the variable $Ip
    $Ip2 = $Ip.ToString()
    $ip3 = $Ip2.Split(" ")
    $ip4 = $ip3[5]
    $ip5 = $ip4.replace("</body>", "")
    $externalIP = $ip5.replace("</html>", "")
    
    $adOpenStatic = 3
    $adLockOptimistic = 3


    #$connectionstring = -join ("Provider=SQLOLEDB;Data Source=", $strcomputer, ";", "Integrated Security=SSPI;Initial Catalog=Master;")
    $connectionstring = -join ("Provider=SQLOLEDB;Data Source=", $strcomputer, ";user Id=", $user, ";Password=", $password, ";Initial Catalog=Master;")
	
    $messagelog = ""

    $t1 = "<td>"
    $t2 = "</td>"
    $outputlineprocess = ""

    $htmltitle = -join ("<html><title>", -join ("SQL Server XRay Vision Report for ", $strcomputer), " </title><body><A NAME=0><h2>", -join ("SQL Server XRay Vision Report for ", $strcomputer), " as of ", (get-date), "</h2>", "`r`n`r`n")
    $htmlend = -join ("<h4>[Written by Von Wallace : End of Report!]</h4></body></html>")

    $htmltableend = -join ("</table><br>")

    $htmlheadingsprocess = -join ("<table border=1>", "<THEAD><TR><TH SCOPE=col>[sp_who2] SPID</TH><TH SCOPE=col>Login</TH><TH SCOPE=col>Host Name</TH><TH SCOPE=col>DB Name</TH><TH SCOPE=col>Status</TH><TH SCOPE=col>Command</TH><TH SCOPE=col>CPU Time</TH><TH SCOPE=col>TSQL Statement</TH></tr></thead>")


    $htmlheadingslock = -join ("<table border=1>", "<THEAD><TR><TH SCOPE=col>[sp_lock] SPID</TH><TH SCOPE=col>DBID</TH><TH SCOPE=col>Object ID</TH><TH SCOPE=col>IndId</TH><TH SCOPE=col>Type</TH><TH SCOPE=col>Resource</TH><TH SCOPE=col>Mode</TH><TH SCOPE=col>Status</TH></tr></thead>")

    $htmlheadingslock2 = -join ("<table border=1>", "<THEAD><TR><TH SCOPE=col>[sys.dm_os_waiting_tasks] blocking_session_id</TH><TH SCOPE=col>wait_duration_ms</TH><TH SCOPE=col>session_id</TH></tr></thead>")


    $htmlheadingssysprocess = -join ("<table border=1>", "<THEAD><TR><TH SCOPE=col>[sysprocesses] SPID</TH><TH SCOPE=col>KPID</TH><TH SCOPE=col>Blocked</TH><TH SCOPE=col>Wait Type</TH><TH SCOPE=col>Wait Time</TH><TH SCOPE=col>Last Wait Type</TH><TH SCOPE=col>DBID</TH><TH SCOPE=col>CPU</TH><TH SCOPE=col>Physical I/O</TH><TH SCOPE=col>Mem Usage</TH><TH SCOPE=col>Login Time</TH><TH SCOPE=col>Last Batch</TH><TH SCOPE=col>Open Tran</TH><TH SCOPE=col>Status</TH><TH SCOPE=col>Host Name</TH><TH SCOPE=col>Program Name</TH><TH SCOPE=col>Host Process</TH><TH SCOPE=col>CMD</TH><TH SCOPE=col>Net Address</TH><TH SCOPE=col>Net Library</TH><TH SCOPE=col>Login Name</TH></tr></thead>")
    $htmlheadingfrag = -join ("<table border=1>", "<THEAD><TR><TH SCOPE=col>[Table Fragmentation] Table Name</TH><TH SCOPE=col>Index Name</TH><TH SCOPE=col>Average Fragmentation in Percent</TH></tr></thead>")


    $outputlineprocess = -join ($outputlineprocess, $htmltitle, $htmlheadingsprocess)


    #' making the connection to your sql server
    #' change yourservername to match your server
    $objConnection = new-Object  -com "ADODB.Connection"
    $objRecordSet = new-Object  -com "ADODB.Recordset"


    # Opens or creates an OLE 2.0 Automation object
    $objConnection.Open($connectionstring)
    # The query goes here
    $objRecordSet.Open("sp_who2", $objConnection, $adOpenStatic, $adLockOptimistic)

    $objRecordSet.MoveFirst()


    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
    $sqlserver = New-Object ('Microsoft.SqlServer.Management.Smo.Server') ($strcomputer) 

    $sqlserver.ConnectionContext.LoginSecure = 0
    $sqlserver.ConnectionContext.Login = $user
    $sqlserver.ConnectionContext.Password = $password

    $sqlserver.ConnectionContext.Connect()


    While (!($objRecordset.EOF)) {
       
       

        $spid = $objRecordset.Fields.Item("SPID").value
        #;inputbuffersql = InputBuffer(spid, SQLServer)
        
        $errorQR = 0   
        try {
            $QueryResult = $sqlserver.ConnectionContext.ExecuteWithResults( -join ("dbcc inputbuffer(", $spid.Trim(), ")"))
        }
        catch {$errorQR = 1}

      
        if ($errorqr -eq 0) {
            $inputbuffersql = $QueryResult.tables[0].eventinfo
        }
        else { $inputbuffersql = "Error"}

        
        $outputlineprocess = -join ($outputlineprocess, "<tr>", $t1, $spid, $t2, $t1, $objRecordset.Fields.Item("Login").value, $t2, $t1, $objRecordset.Fields.Item("HostName").value, $t2, $t1, $objRecordset.Fields.Item("DBName").value, $t2, $t1, $objRecordset.Fields.Item("Status").value, $t2, $t1, $objRecordset.Fields.Item("Command").value, $t2, $t1, $objRecordset.Fields.Item("CPUTime").value, $t2, $t1, $inputbuffersql, "</tr>")

        $QueryResult = ""
        $objRecordset.MoveNext()
        
        
    }

    $sqlserver.ConnectionContext.DisConnect()
    $SQLServer = ""


    $outputlineprocess = -join ($outputlineprocess, $htmltableend)

    
    $objRecordSet.Close()
   

    #;'=================================================================================

    $outputlinelocks = ""

   
    $objRecordSet.Open("sp_lock", $objConnection, $adOpenStatic, $adLockOptimistic)

    $objRecordSet.MoveFirst()



    While (!($objRecordset.EOF)) {    
        $outputlinelocks = -join ($outputlinelocks, "<tr>", $t1, $objRecordset.Fields.Item("spid").value, $t2, $t1, $objRecordset.Fields.Item("dbid").value, $t2, $t1, $objRecordset.Fields.Item("objid").value, $t2, $t1, $objRecordset.Fields.Item("indid").value, $t2, $t1, $objRecordset.Fields.Item("Type").value, $t2, $t1, $objRecordset.Fields.Item("Resource").value, $t2, $t1, $objRecordset.Fields.Item("mode").value, $t2, $t1, $objRecordset.Fields.Item("status").value, "</tr>")


        $objRecordset.MoveNext()

    }

    $outputlinelocks = -join ($outputlinelocks, $htmltableend)

    $objRecordSet.Close()

    #;****************************************************************************************

    $outputlinelocks2 = ""

   
    $objRecordSet.Open("SELECT blocking_session_id, wait_duration_ms, session_id FROM sys.dm_os_waiting_tasks WHERE blocking_session_id IS NOT NULL", $objConnection, $adOpenStatic, $adLockOptimistic)

    try {
        $objRecordSet.MoveFirst()
        While (!($objRecordset.EOF)) {    
            $outputlinelocks2 = -join ($outputlinelocks2, "<tr>", $t1, $objRecordset.Fields.Item("blocking_session_id").value, $t2, $t1, $objRecordset.Fields.Item("wait_duration_ms").value, $t2, $t1, $objRecordset.Fields.Item("session_id").value, $t2, "</tr>")


            $objRecordset.MoveNext()

        }

    }
    catch {
        
    }

    
    $outputlinelocks2 = -join ($outputlinelocks2, $htmltableend)

    $objRecordSet.Close()



   
    #;****************************************************************************************

    $outputlinesysprocess = ""

   
    $objRecordSet.Open("select spid,kpid,blocked,waittype,waittime,lastwaittype,dbid,cpu,physical_io,memusage,login_time,last_batch,open_tran,status,hostname,program_name,hostprocess,cmd,net_address,net_library,loginame  from sysprocesses", $objConnection, $adOpenStatic, $adLockOptimistic)

    $objRecordSet.MoveFirst()



    While (!($objRecordset.EOF)) {
        $outputlinesysprocess = -join ($outputlinesysprocess, "<tr>", $t1, $objRecordset.Fields.Item("spid").value, $t2, $t1, $objRecordset.Fields.Item("kpid").value, $t2, $t1, $objRecordset.Fields.Item("blocked").value, $t2, $t1, $objRecordset.Fields.Item("waittype").value, $t2, $t1, $objRecordset.Fields.Item("waittime").value, $t2, $t1, $objRecordset.Fields.Item("lastwaittype").value, $t2, $t1, $objRecordset.Fields.Item("dbid").value, $t2, $t1, $objRecordset.Fields.Item("cpu").value, $t2, $t1, $objRecordset.Fields.Item("physical_io").value, $t2, $t1, $objRecordset.Fields.Item("memusage").value, $t2, $t1, $objRecordset.Fields.Item("login_time").value, $t2, $t1, $objRecordset.Fields.Item("last_batch").value, $t2, $t1, $objRecordset.Fields.Item("open_tran").value, $t2, $t1, $objRecordset.Fields.Item("status").value, $t2, $t1, $objRecordset.Fields.Item("hostname").value, $t2, $t1, $objRecordset.Fields.Item("program_name").value, $t2, $t1, $objRecordset.Fields.Item("hostprocess").value, $t2, $t1, $objRecordset.Fields.Item("cmd").value, $t2, $t1, $objRecordset.Fields.Item("net_address").value, $t2, $t1, $objRecordset.Fields.Item("net_library").value, $t2, $t1, $objRecordset.Fields.Item("loginame").value, $t2, "</tr>")


        $objRecordset.MoveNext()

    }



    $outputlinesysprocess = -join ($outputlinesysprocess, $htmltableend)

    $objRecordSet.Close()

   
    #;%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


    $outputlinefrag = ""
   
    $objRecordSet.Open("SELECT    OBJECT_NAME(i.object_id) AS TableName,i.name AS IndexName,indexstats.avg_fragmentation_in_percent FROM    sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'DETAILED') indexstats INNER JOIN sys.indexes i ON i.object_id = indexstats.object_id AND i.index_id = indexstats.index_id WHERE indexstats.avg_fragmentation_in_percent > 20", $objConnection, $adOpenStatic, $adLockOptimistic)

    $objRecordSet.RecordCount
    If ($objRecordSet.RecordCount -lt 1) {
        $outputlinefrag = "<br>"
    }
    else
    {$objRecordSet.MoveFirst()}


    While (!($objRecordset.EOF)) {
        $outputlinefrag = -join ($outputlinefrag, "<tr>", $t1, $objRecordset.Fields.Item("TableName").value, $t2, $t1, $objRecordset.Fields.Item("Indexname").value, $t2, $t1, $objRecordset.Fields.Item("avg_fragmentation_in_percent").value, "</tr>")


        $objRecordset.MoveNext()

    }
    $outputlinefrag = -join ($outputlinefrag, $htmltableend)



    
    $objRecordSet.Close()

    $objConnection.Close()

    $objrecordset = ""
    $objConnection = ""

    
    $outputcomputerinfo = "<br>"
    $outputlinecomputerinfo = -join ($outputlineprocess, $htmlheadingslock, $outputlinelocks,$htmlheadingslock2 ,$outputlinelocks2, $htmlheadingssysprocess, $outputlinesysprocess, $htmlheadingfrag, $outputlinefrag, $outputcomputerinfo)

    $reportoutput = "
    <style>
    TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
    TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
    TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
    </style>
    "

    $reportoutput = -join ($reportoutput, $outputlinecomputerinfo, $htmlend)

    
    $subject = -join ("SQL Server XRay Vision Report for ", $strcomputer, " ", (get-date))

    
    $message = new-object Net.Mail.MailMessage;
    
    $message.From = $smtpfrom;
    $message.To.Add($smtpto);
    $message.Subject = $subject ;
    $message.IsBodyHTML = $true
    $message.Body = $reportoutput
    

    $smtp = new-object Net.Mail.SmtpClient($smtpserver, $smtpport);
    $smtp.EnableSSL = $true;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($sendUsername, $sendPassword);
    $smtp.send($message);
    write-host "Mail Sent to "  $smtpto ; 


}
catch {
    $ErrorMessage = $Error[0].tostring() + $error[0].InvocationInfo.PositionMessage
    
    write-host $ErrorMessage
        
    $message = new-object Net.Mail.MailMessage;
      
    $message.From = $smtpfrom;
    $message.To.Add($smtpto);
    $message.Subject = $strcomputer + -join ("SQL Server XRay Vision Report for ", $strcomputer, "  Error ", (get-date)) ;
    $message.IsBodyHTML = $true
    $message.Body = $ErrorMessage + "<br>Version: " + $Version + "<br>External IP Address: " + $externalIP  
        
        
    $smtp = new-object Net.Mail.SmtpClient($smtpserver, $smtpport);
    $smtp.EnableSSL = $true;  
    $smtp.Credentials = New-Object System.Net.NetworkCredential($sendUsername, $sendPassword);
    $smtp.send($message);
    write-host "Mail Sent to "  $smtpto ; 
}
        

