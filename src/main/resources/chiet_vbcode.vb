Option Explicit
Function action
Dim objConnection
Dim strConnectionString
Dim strSQL
Dim objCommand
strConnectionString = "Provider=MSDASQL;DSN=EX_WINCC_DATA;UID=;PWD=;"
Set objConnection = CreateObject("ADODB.Connection")
objConnection.ConnectionString = strConnectionString
objConnection.Open

Dim tag_NewTag
tag_NewTag = HMIRuntime.Tags("NewTag").Read
strSQL = "INSERT INTO z_tag_NewTag (tag_value, created) values(" & tag_NewTag & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_NewTag = Nothing
Dim tag_please_wait_refre
tag_please_wait_refre = HMIRuntime.Tags("please_wait_refre").Read
strSQL = "INSERT INTO z_tag_please_wait_refre (tag_value, created) values(" & tag_please_wait_refre & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_please_wait_refre = Nothing
Dim tag_seconds
tag_seconds = HMIRuntime.Tags("seconds").Read
strSQL = "INSERT INTO z_tag_seconds (tag_value, created) values(" & tag_seconds & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_seconds = Nothing
Dim tag_date
tag_date = HMIRuntime.Tags("date").Read
strSQL = "INSERT INTO z_tag_date (tag_value, created) values(" & tag_date & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date = Nothing
Dim tag_time
tag_time = HMIRuntime.Tags("time").Read
strSQL = "INSERT INTO z_tag_time (tag_value, created) values(" & tag_time & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_time = Nothing
Dim tag_exit
tag_exit = HMIRuntime.Tags("exit").Read
strSQL = "INSERT INTO z_tag_exit (tag_value, created) values(" & tag_exit & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_exit = Nothing
Dim tag_ACK
tag_ACK = HMIRuntime.Tags("ACK").Read
strSQL = "INSERT INTO z_tag_ACK (tag_value, created) values(" & tag_ACK & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ACK = Nothing
Dim tag_Actual_fill
tag_Actual_fill = HMIRuntime.Tags("Actual_fill").Read
strSQL = "INSERT INTO z_tag_Actual_fill (tag_value, created) values(" & tag_Actual_fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Actual_fill = Nothing
Dim tag_MachineIsON_fill
tag_MachineIsON_fill = HMIRuntime.Tags("MachineIsON_fill").Read
strSQL = "INSERT INTO z_tag_MachineIsON_fill (tag_value, created) values(" & tag_MachineIsON_fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_fill = Nothing
Dim tag_B2105_start
tag_B2105_start = HMIRuntime.Tags("B2105_start").Read
strSQL = "INSERT INTO z_tag_B2105_start (tag_value, created) values(" & tag_B2105_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2105_start = Nothing
Dim tag_B2105_freq
tag_B2105_freq = HMIRuntime.Tags("B2105_freq").Read
strSQL = "INSERT INTO z_tag_B2105_freq (tag_value, created) values(" & tag_B2105_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2105_freq = Nothing
Dim tag_B2110_start
tag_B2110_start = HMIRuntime.Tags("B2110_start").Read
strSQL = "INSERT INTO z_tag_B2110_start (tag_value, created) values(" & tag_B2110_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2110_start = Nothing
Dim tag_B2110_freq
tag_B2110_freq = HMIRuntime.Tags("B2110_freq").Read
strSQL = "INSERT INTO z_tag_B2110_freq (tag_value, created) values(" & tag_B2110_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2110_freq = Nothing
Dim tag_B2111_start
tag_B2111_start = HMIRuntime.Tags("B2111_start").Read
strSQL = "INSERT INTO z_tag_B2111_start (tag_value, created) values(" & tag_B2111_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2111_start = Nothing
Dim tag_B2111_freq
tag_B2111_freq = HMIRuntime.Tags("B2111_freq").Read
strSQL = "INSERT INTO z_tag_B2111_freq (tag_value, created) values(" & tag_B2111_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2111_freq = Nothing
Dim tag_B2115_start
tag_B2115_start = HMIRuntime.Tags("B2115_start").Read
strSQL = "INSERT INTO z_tag_B2115_start (tag_value, created) values(" & tag_B2115_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2115_start = Nothing
Dim tag_B2115_freq
tag_B2115_freq = HMIRuntime.Tags("B2115_freq").Read
strSQL = "INSERT INTO z_tag_B2115_freq (tag_value, created) values(" & tag_B2115_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2115_freq = Nothing
Dim tag_B2207_start
tag_B2207_start = HMIRuntime.Tags("B2207_start").Read
strSQL = "INSERT INTO z_tag_B2207_start (tag_value, created) values(" & tag_B2207_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2207_start = Nothing
Dim tag_B2207_freq
tag_B2207_freq = HMIRuntime.Tags("B2207_freq").Read
strSQL = "INSERT INTO z_tag_B2207_freq (tag_value, created) values(" & tag_B2207_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2207_freq = Nothing
Dim tag_B2210_start
tag_B2210_start = HMIRuntime.Tags("B2210_start").Read
strSQL = "INSERT INTO z_tag_B2210_start (tag_value, created) values(" & tag_B2210_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2210_start = Nothing
Dim tag_B2210_freq
tag_B2210_freq = HMIRuntime.Tags("B2210_freq").Read
strSQL = "INSERT INTO z_tag_B2210_freq (tag_value, created) values(" & tag_B2210_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2210_freq = Nothing
Dim tag_B2213_start
tag_B2213_start = HMIRuntime.Tags("B2213_start").Read
strSQL = "INSERT INTO z_tag_B2213_start (tag_value, created) values(" & tag_B2213_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2213_start = Nothing
Dim tag_B2213_freq
tag_B2213_freq = HMIRuntime.Tags("B2213_freq").Read
strSQL = "INSERT INTO z_tag_B2213_freq (tag_value, created) values(" & tag_B2213_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2213_freq = Nothing
Dim tag_B2215_start
tag_B2215_start = HMIRuntime.Tags("B2215_start").Read
strSQL = "INSERT INTO z_tag_B2215_start (tag_value, created) values(" & tag_B2215_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2215_start = Nothing
Dim tag_B2215_freq
tag_B2215_freq = HMIRuntime.Tags("B2215_freq").Read
strSQL = "INSERT INTO z_tag_B2215_freq (tag_value, created) values(" & tag_B2215_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2215_freq = Nothing
Dim tag_B2218_101_start
tag_B2218_101_start = HMIRuntime.Tags("B2218_101_start").Read
strSQL = "INSERT INTO z_tag_B2218_101_start (tag_value, created) values(" & tag_B2218_101_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218_101_start = Nothing
Dim tag_B2218__101_freq
tag_B2218__101_freq = HMIRuntime.Tags("B2218__101_freq").Read
strSQL = "INSERT INTO z_tag_B2218__101_freq (tag_value, created) values(" & tag_B2218__101_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218__101_freq = Nothing
Dim tag_B2218_121_start
tag_B2218_121_start = HMIRuntime.Tags("B2218_121_start").Read
strSQL = "INSERT INTO z_tag_B2218_121_start (tag_value, created) values(" & tag_B2218_121_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218_121_start = Nothing
Dim tag_B2218__121_freq
tag_B2218__121_freq = HMIRuntime.Tags("B2218__121_freq").Read
strSQL = "INSERT INTO z_tag_B2218__121_freq (tag_value, created) values(" & tag_B2218__121_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218__121_freq = Nothing
Dim tag_B2218_141_start
tag_B2218_141_start = HMIRuntime.Tags("B2218_141_start").Read
strSQL = "INSERT INTO z_tag_B2218_141_start (tag_value, created) values(" & tag_B2218_141_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218_141_start = Nothing
Dim tag_B2218__141_freq
tag_B2218__141_freq = HMIRuntime.Tags("B2218__141_freq").Read
strSQL = "INSERT INTO z_tag_B2218__141_freq (tag_value, created) values(" & tag_B2218__141_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2218__141_freq = Nothing
Dim tag_B2222_start
tag_B2222_start = HMIRuntime.Tags("B2222_start").Read
strSQL = "INSERT INTO z_tag_B2222_start (tag_value, created) values(" & tag_B2222_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2222_start = Nothing
Dim tag_B2222_freq
tag_B2222_freq = HMIRuntime.Tags("B2222_freq").Read
strSQL = "INSERT INTO z_tag_B2222_freq (tag_value, created) values(" & tag_B2222_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2222_freq = Nothing
Dim tag_B2225_start
tag_B2225_start = HMIRuntime.Tags("B2225_start").Read
strSQL = "INSERT INTO z_tag_B2225_start (tag_value, created) values(" & tag_B2225_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2225_start = Nothing
Dim tag_B2225_freq
tag_B2225_freq = HMIRuntime.Tags("B2225_freq").Read
strSQL = "INSERT INTO z_tag_B2225_freq (tag_value, created) values(" & tag_B2225_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2225_freq = Nothing
Dim tag_B2229_start
tag_B2229_start = HMIRuntime.Tags("B2229_start").Read
strSQL = "INSERT INTO z_tag_B2229_start (tag_value, created) values(" & tag_B2229_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2229_start = Nothing
Dim tag_B2229_freq
tag_B2229_freq = HMIRuntime.Tags("B2229_freq").Read
strSQL = "INSERT INTO z_tag_B2229_freq (tag_value, created) values(" & tag_B2229_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2229_freq = Nothing
Dim tag_B2305_start
tag_B2305_start = HMIRuntime.Tags("B2305_start").Read
strSQL = "INSERT INTO z_tag_B2305_start (tag_value, created) values(" & tag_B2305_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2305_start = Nothing
Dim tag_B2305_freq
tag_B2305_freq = HMIRuntime.Tags("B2305_freq").Read
strSQL = "INSERT INTO z_tag_B2305_freq (tag_value, created) values(" & tag_B2305_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2305_freq = Nothing
Dim tag_B2309_start
tag_B2309_start = HMIRuntime.Tags("B2309_start").Read
strSQL = "INSERT INTO z_tag_B2309_start (tag_value, created) values(" & tag_B2309_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2309_start = Nothing
Dim tag_B2309_freq
tag_B2309_freq = HMIRuntime.Tags("B2309_freq").Read
strSQL = "INSERT INTO z_tag_B2309_freq (tag_value, created) values(" & tag_B2309_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2309_freq = Nothing
Dim tag_B2312_start
tag_B2312_start = HMIRuntime.Tags("B2312_start").Read
strSQL = "INSERT INTO z_tag_B2312_start (tag_value, created) values(" & tag_B2312_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2312_start = Nothing
Dim tag_B2312_freq
tag_B2312_freq = HMIRuntime.Tags("B2312_freq").Read
strSQL = "INSERT INTO z_tag_B2312_freq (tag_value, created) values(" & tag_B2312_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2312_freq = Nothing
Dim tag_B2404_start
tag_B2404_start = HMIRuntime.Tags("B2404_start").Read
strSQL = "INSERT INTO z_tag_B2404_start (tag_value, created) values(" & tag_B2404_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2404_start = Nothing
Dim tag_B2404_freq
tag_B2404_freq = HMIRuntime.Tags("B2404_freq").Read
strSQL = "INSERT INTO z_tag_B2404_freq (tag_value, created) values(" & tag_B2404_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2404_freq = Nothing
Dim tag_B2410_start
tag_B2410_start = HMIRuntime.Tags("B2410_start").Read
strSQL = "INSERT INTO z_tag_B2410_start (tag_value, created) values(" & tag_B2410_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2410_start = Nothing
Dim tag_B2410_freq
tag_B2410_freq = HMIRuntime.Tags("B2410_freq").Read
strSQL = "INSERT INTO z_tag_B2410_freq (tag_value, created) values(" & tag_B2410_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2410_freq = Nothing
Dim tag_B2415_start
tag_B2415_start = HMIRuntime.Tags("B2415_start").Read
strSQL = "INSERT INTO z_tag_B2415_start (tag_value, created) values(" & tag_B2415_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2415_start = Nothing
Dim tag_B2415_freq
tag_B2415_freq = HMIRuntime.Tags("B2415_freq").Read
strSQL = "INSERT INTO z_tag_B2415_freq (tag_value, created) values(" & tag_B2415_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2415_freq = Nothing
Dim tag_B2504_start
tag_B2504_start = HMIRuntime.Tags("B2504_start").Read
strSQL = "INSERT INTO z_tag_B2504_start (tag_value, created) values(" & tag_B2504_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2504_start = Nothing
Dim tag_B2504_freq
tag_B2504_freq = HMIRuntime.Tags("B2504_freq").Read
strSQL = "INSERT INTO z_tag_B2504_freq (tag_value, created) values(" & tag_B2504_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2504_freq = Nothing
Dim tag_B2507_start
tag_B2507_start = HMIRuntime.Tags("B2507_start").Read
strSQL = "INSERT INTO z_tag_B2507_start (tag_value, created) values(" & tag_B2507_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2507_start = Nothing
Dim tag_B2507_freq
tag_B2507_freq = HMIRuntime.Tags("B2507_freq").Read
strSQL = "INSERT INTO z_tag_B2507_freq (tag_value, created) values(" & tag_B2507_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2507_freq = Nothing
Dim tag_B2513_start
tag_B2513_start = HMIRuntime.Tags("B2513_start").Read
strSQL = "INSERT INTO z_tag_B2513_start (tag_value, created) values(" & tag_B2513_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2513_start = Nothing
Dim tag_B2513_freq
tag_B2513_freq = HMIRuntime.Tags("B2513_freq").Read
strSQL = "INSERT INTO z_tag_B2513_freq (tag_value, created) values(" & tag_B2513_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2513_freq = Nothing
Dim tag_B2515_start
tag_B2515_start = HMIRuntime.Tags("B2515_start").Read
strSQL = "INSERT INTO z_tag_B2515_start (tag_value, created) values(" & tag_B2515_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2515_start = Nothing
Dim tag_B2515_freq
tag_B2515_freq = HMIRuntime.Tags("B2515_freq").Read
strSQL = "INSERT INTO z_tag_B2515_freq (tag_value, created) values(" & tag_B2515_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2515_freq = Nothing
Dim tag_B2607_start
tag_B2607_start = HMIRuntime.Tags("B2607_start").Read
strSQL = "INSERT INTO z_tag_B2607_start (tag_value, created) values(" & tag_B2607_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2607_start = Nothing
Dim tag_B2607_freq
tag_B2607_freq = HMIRuntime.Tags("B2607_freq").Read
strSQL = "INSERT INTO z_tag_B2607_freq (tag_value, created) values(" & tag_B2607_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2607_freq = Nothing
Dim tag_B2610_start
tag_B2610_start = HMIRuntime.Tags("B2610_start").Read
strSQL = "INSERT INTO z_tag_B2610_start (tag_value, created) values(" & tag_B2610_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2610_start = Nothing
Dim tag_B2610_freq
tag_B2610_freq = HMIRuntime.Tags("B2610_freq").Read
strSQL = "INSERT INTO z_tag_B2610_freq (tag_value, created) values(" & tag_B2610_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2610_freq = Nothing
Dim tag_B2612_start
tag_B2612_start = HMIRuntime.Tags("B2612_start").Read
strSQL = "INSERT INTO z_tag_B2612_start (tag_value, created) values(" & tag_B2612_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2612_start = Nothing
Dim tag_B2612_freq
tag_B2612_freq = HMIRuntime.Tags("B2612_freq").Read
strSQL = "INSERT INTO z_tag_B2612_freq (tag_value, created) values(" & tag_B2612_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2612_freq = Nothing
Dim tag_B2614_101_start
tag_B2614_101_start = HMIRuntime.Tags("B2614_101_start").Read
strSQL = "INSERT INTO z_tag_B2614_101_start (tag_value, created) values(" & tag_B2614_101_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_101_start = Nothing
Dim tag_B2614_101_freq
tag_B2614_101_freq = HMIRuntime.Tags("B2614_101_freq").Read
strSQL = "INSERT INTO z_tag_B2614_101_freq (tag_value, created) values(" & tag_B2614_101_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_101_freq = Nothing
Dim tag_B2614_121_start
tag_B2614_121_start = HMIRuntime.Tags("B2614_121_start").Read
strSQL = "INSERT INTO z_tag_B2614_121_start (tag_value, created) values(" & tag_B2614_121_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_121_start = Nothing
Dim tag_B2614_121_freq
tag_B2614_121_freq = HMIRuntime.Tags("B2614_121_freq").Read
strSQL = "INSERT INTO z_tag_B2614_121_freq (tag_value, created) values(" & tag_B2614_121_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_121_freq = Nothing
Dim tag_B2614_141_start
tag_B2614_141_start = HMIRuntime.Tags("B2614_141_start").Read
strSQL = "INSERT INTO z_tag_B2614_141_start (tag_value, created) values(" & tag_B2614_141_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_141_start = Nothing
Dim tag_B2614_141_freq
tag_B2614_141_freq = HMIRuntime.Tags("B2614_141_freq").Read
strSQL = "INSERT INTO z_tag_B2614_141_freq (tag_value, created) values(" & tag_B2614_141_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2614_141_freq = Nothing
Dim tag_B2616_start
tag_B2616_start = HMIRuntime.Tags("B2616_start").Read
strSQL = "INSERT INTO z_tag_B2616_start (tag_value, created) values(" & tag_B2616_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2616_start = Nothing
Dim tag_B2616_freq
tag_B2616_freq = HMIRuntime.Tags("B2616_freq").Read
strSQL = "INSERT INTO z_tag_B2616_freq (tag_value, created) values(" & tag_B2616_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2616_freq = Nothing
Dim tag_B2620_start
tag_B2620_start = HMIRuntime.Tags("B2620_start").Read
strSQL = "INSERT INTO z_tag_B2620_start (tag_value, created) values(" & tag_B2620_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2620_start = Nothing
Dim tag_B2620_freq
tag_B2620_freq = HMIRuntime.Tags("B2620_freq").Read
strSQL = "INSERT INTO z_tag_B2620_freq (tag_value, created) values(" & tag_B2620_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2620_freq = Nothing
Dim tag_B2623_start
tag_B2623_start = HMIRuntime.Tags("B2623_start").Read
strSQL = "INSERT INTO z_tag_B2623_start (tag_value, created) values(" & tag_B2623_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2623_start = Nothing
Dim tag_B2623_freq
tag_B2623_freq = HMIRuntime.Tags("B2623_freq").Read
strSQL = "INSERT INTO z_tag_B2623_freq (tag_value, created) values(" & tag_B2623_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2623_freq = Nothing
Dim tag_B2706_start
tag_B2706_start = HMIRuntime.Tags("B2706_start").Read
strSQL = "INSERT INTO z_tag_B2706_start (tag_value, created) values(" & tag_B2706_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2706_start = Nothing
Dim tag_B2706_freq
tag_B2706_freq = HMIRuntime.Tags("B2706_freq").Read
strSQL = "INSERT INTO z_tag_B2706_freq (tag_value, created) values(" & tag_B2706_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2706_freq = Nothing
Dim tag_B2709_start
tag_B2709_start = HMIRuntime.Tags("B2709_start").Read
strSQL = "INSERT INTO z_tag_B2709_start (tag_value, created) values(" & tag_B2709_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2709_start = Nothing
Dim tag_B2709_freq
tag_B2709_freq = HMIRuntime.Tags("B2709_freq").Read
strSQL = "INSERT INTO z_tag_B2709_freq (tag_value, created) values(" & tag_B2709_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2709_freq = Nothing
Dim tag_B2712_start
tag_B2712_start = HMIRuntime.Tags("B2712_start").Read
strSQL = "INSERT INTO z_tag_B2712_start (tag_value, created) values(" & tag_B2712_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2712_start = Nothing
Dim tag_B2712_freq
tag_B2712_freq = HMIRuntime.Tags("B2712_freq").Read
strSQL = "INSERT INTO z_tag_B2712_freq (tag_value, created) values(" & tag_B2712_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2712_freq = Nothing
Dim tag_B2715_start
tag_B2715_start = HMIRuntime.Tags("B2715_start").Read
strSQL = "INSERT INTO z_tag_B2715_start (tag_value, created) values(" & tag_B2715_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2715_start = Nothing
Dim tag_B2715_freq
tag_B2715_freq = HMIRuntime.Tags("B2715_freq").Read
strSQL = "INSERT INTO z_tag_B2715_freq (tag_value, created) values(" & tag_B2715_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2715_freq = Nothing
Dim tag_B2719_start
tag_B2719_start = HMIRuntime.Tags("B2719_start").Read
strSQL = "INSERT INTO z_tag_B2719_start (tag_value, created) values(" & tag_B2719_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2719_start = Nothing
Dim tag_B2719_freq
tag_B2719_freq = HMIRuntime.Tags("B2719_freq").Read
strSQL = "INSERT INTO z_tag_B2719_freq (tag_value, created) values(" & tag_B2719_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2719_freq = Nothing
Dim tag_B2723_start
tag_B2723_start = HMIRuntime.Tags("B2723_start").Read
strSQL = "INSERT INTO z_tag_B2723_start (tag_value, created) values(" & tag_B2723_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2723_start = Nothing
Dim tag_B2723_freq
tag_B2723_freq = HMIRuntime.Tags("B2723_freq").Read
strSQL = "INSERT INTO z_tag_B2723_freq (tag_value, created) values(" & tag_B2723_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2723_freq = Nothing
Dim tag_B2725_start
tag_B2725_start = HMIRuntime.Tags("B2725_start").Read
strSQL = "INSERT INTO z_tag_B2725_start (tag_value, created) values(" & tag_B2725_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2725_start = Nothing
Dim tag_B2725_freq
tag_B2725_freq = HMIRuntime.Tags("B2725_freq").Read
strSQL = "INSERT INTO z_tag_B2725_freq (tag_value, created) values(" & tag_B2725_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_B2725_freq = Nothing
Dim tag_G2105_start
tag_G2105_start = HMIRuntime.Tags("G2105_start").Read
strSQL = "INSERT INTO z_tag_G2105_start (tag_value, created) values(" & tag_G2105_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2105_start = Nothing
Dim tag_G2105_freq
tag_G2105_freq = HMIRuntime.Tags("G2105_freq").Read
strSQL = "INSERT INTO z_tag_G2105_freq (tag_value, created) values(" & tag_G2105_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2105_freq = Nothing
Dim tag_G2109_start
tag_G2109_start = HMIRuntime.Tags("G2109_start").Read
strSQL = "INSERT INTO z_tag_G2109_start (tag_value, created) values(" & tag_G2109_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2109_start = Nothing
Dim tag_G2109_freq
tag_G2109_freq = HMIRuntime.Tags("G2109_freq").Read
strSQL = "INSERT INTO z_tag_G2109_freq (tag_value, created) values(" & tag_G2109_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2109_freq = Nothing
Dim tag_G2112_start
tag_G2112_start = HMIRuntime.Tags("G2112_start").Read
strSQL = "INSERT INTO z_tag_G2112_start (tag_value, created) values(" & tag_G2112_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2112_start = Nothing
Dim tag_G2112_freq
tag_G2112_freq = HMIRuntime.Tags("G2112_freq").Read
strSQL = "INSERT INTO z_tag_G2112_freq (tag_value, created) values(" & tag_G2112_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2112_freq = Nothing
Dim tag_G2116_start
tag_G2116_start = HMIRuntime.Tags("G2116_start").Read
strSQL = "INSERT INTO z_tag_G2116_start (tag_value, created) values(" & tag_G2116_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2116_start = Nothing
Dim tag_G2116_freq
tag_G2116_freq = HMIRuntime.Tags("G2116_freq").Read
strSQL = "INSERT INTO z_tag_G2116_freq (tag_value, created) values(" & tag_G2116_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2116_freq = Nothing
Dim tag_G2118_start
tag_G2118_start = HMIRuntime.Tags("G2118_start").Read
strSQL = "INSERT INTO z_tag_G2118_start (tag_value, created) values(" & tag_G2118_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2118_start = Nothing
Dim tag_G2118_freq
tag_G2118_freq = HMIRuntime.Tags("G2118_freq").Read
strSQL = "INSERT INTO z_tag_G2118_freq (tag_value, created) values(" & tag_G2118_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2118_freq = Nothing
Dim tag_G2203_start
tag_G2203_start = HMIRuntime.Tags("G2203_start").Read
strSQL = "INSERT INTO z_tag_G2203_start (tag_value, created) values(" & tag_G2203_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2203_start = Nothing
Dim tag_G2203_freq
tag_G2203_freq = HMIRuntime.Tags("G2203_freq").Read
strSQL = "INSERT INTO z_tag_G2203_freq (tag_value, created) values(" & tag_G2203_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2203_freq = Nothing
Dim tag_G2206_start
tag_G2206_start = HMIRuntime.Tags("G2206_start").Read
strSQL = "INSERT INTO z_tag_G2206_start (tag_value, created) values(" & tag_G2206_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2206_start = Nothing
Dim tag_G2206_freq
tag_G2206_freq = HMIRuntime.Tags("G2206_freq").Read
strSQL = "INSERT INTO z_tag_G2206_freq (tag_value, created) values(" & tag_G2206_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2206_freq = Nothing
Dim tag_G2214_start
tag_G2214_start = HMIRuntime.Tags("G2214_start").Read
strSQL = "INSERT INTO z_tag_G2214_start (tag_value, created) values(" & tag_G2214_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2214_start = Nothing
Dim tag_G2214_freq
tag_G2214_freq = HMIRuntime.Tags("G2214_freq").Read
strSQL = "INSERT INTO z_tag_G2214_freq (tag_value, created) values(" & tag_G2214_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2214_freq = Nothing
Dim tag_G2306_start
tag_G2306_start = HMIRuntime.Tags("G2306_start").Read
strSQL = "INSERT INTO z_tag_G2306_start (tag_value, created) values(" & tag_G2306_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2306_start = Nothing
Dim tag_G2306_freq
tag_G2306_freq = HMIRuntime.Tags("G2306_freq").Read
strSQL = "INSERT INTO z_tag_G2306_freq (tag_value, created) values(" & tag_G2306_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2306_freq = Nothing
Dim tag_G2311_start
tag_G2311_start = HMIRuntime.Tags("G2311_start").Read
strSQL = "INSERT INTO z_tag_G2311_start (tag_value, created) values(" & tag_G2311_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2311_start = Nothing
Dim tag_G2311_freq
tag_G2311_freq = HMIRuntime.Tags("G2311_freq").Read
strSQL = "INSERT INTO z_tag_G2311_freq (tag_value, created) values(" & tag_G2311_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2311_freq = Nothing
Dim tag_G2404_start
tag_G2404_start = HMIRuntime.Tags("G2404_start").Read
strSQL = "INSERT INTO z_tag_G2404_start (tag_value, created) values(" & tag_G2404_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2404_start = Nothing
Dim tag_G2404_freq
tag_G2404_freq = HMIRuntime.Tags("G2404_freq").Read
strSQL = "INSERT INTO z_tag_G2404_freq (tag_value, created) values(" & tag_G2404_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2404_freq = Nothing
Dim tag_G2406_start
tag_G2406_start = HMIRuntime.Tags("G2406_start").Read
strSQL = "INSERT INTO z_tag_G2406_start (tag_value, created) values(" & tag_G2406_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2406_start = Nothing
Dim tag_G2406_freq
tag_G2406_freq = HMIRuntime.Tags("G2406_freq").Read
strSQL = "INSERT INTO z_tag_G2406_freq (tag_value, created) values(" & tag_G2406_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2406_freq = Nothing
Dim tag_G2408_start
tag_G2408_start = HMIRuntime.Tags("G2408_start").Read
strSQL = "INSERT INTO z_tag_G2408_start (tag_value, created) values(" & tag_G2408_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2408_start = Nothing
Dim tag_G2408_freq
tag_G2408_freq = HMIRuntime.Tags("G2408_freq").Read
strSQL = "INSERT INTO z_tag_G2408_freq (tag_value, created) values(" & tag_G2408_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2408_freq = Nothing
Dim tag_G2411_start
tag_G2411_start = HMIRuntime.Tags("G2411_start").Read
strSQL = "INSERT INTO z_tag_G2411_start (tag_value, created) values(" & tag_G2411_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2411_start = Nothing
Dim tag_G2411_freq
tag_G2411_freq = HMIRuntime.Tags("G2411_freq").Read
strSQL = "INSERT INTO z_tag_G2411_freq (tag_value, created) values(" & tag_G2411_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2411_freq = Nothing
Dim tag_G2413_start
tag_G2413_start = HMIRuntime.Tags("G2413_start").Read
strSQL = "INSERT INTO z_tag_G2413_start (tag_value, created) values(" & tag_G2413_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2413_start = Nothing
Dim tag_G2413_freq
tag_G2413_freq = HMIRuntime.Tags("G2413_freq").Read
strSQL = "INSERT INTO z_tag_G2413_freq (tag_value, created) values(" & tag_G2413_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2413_freq = Nothing
Dim tag_G2502_start
tag_G2502_start = HMIRuntime.Tags("G2502_start").Read
strSQL = "INSERT INTO z_tag_G2502_start (tag_value, created) values(" & tag_G2502_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2502_start = Nothing
Dim tag_G2502_freq
tag_G2502_freq = HMIRuntime.Tags("G2502_freq").Read
strSQL = "INSERT INTO z_tag_G2502_freq (tag_value, created) values(" & tag_G2502_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2502_freq = Nothing
Dim tag_G2505_start
tag_G2505_start = HMIRuntime.Tags("G2505_start").Read
strSQL = "INSERT INTO z_tag_G2505_start (tag_value, created) values(" & tag_G2505_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2505_start = Nothing
Dim tag_G2505_freq
tag_G2505_freq = HMIRuntime.Tags("G2505_freq").Read
strSQL = "INSERT INTO z_tag_G2505_freq (tag_value, created) values(" & tag_G2505_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2505_freq = Nothing
Dim tag_G2506_start
tag_G2506_start = HMIRuntime.Tags("G2506_start").Read
strSQL = "INSERT INTO z_tag_G2506_start (tag_value, created) values(" & tag_G2506_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2506_start = Nothing
Dim tag_G2506_freq
tag_G2506_freq = HMIRuntime.Tags("G2506_freq").Read
strSQL = "INSERT INTO z_tag_G2506_freq (tag_value, created) values(" & tag_G2506_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2506_freq = Nothing
Dim tag_G2508_start
tag_G2508_start = HMIRuntime.Tags("G2508_start").Read
strSQL = "INSERT INTO z_tag_G2508_start (tag_value, created) values(" & tag_G2508_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2508_start = Nothing
Dim tag_G2508_freq
tag_G2508_freq = HMIRuntime.Tags("G2508_freq").Read
strSQL = "INSERT INTO z_tag_G2508_freq (tag_value, created) values(" & tag_G2508_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2508_freq = Nothing
Dim tag_G2510_start
tag_G2510_start = HMIRuntime.Tags("G2510_start").Read
strSQL = "INSERT INTO z_tag_G2510_start (tag_value, created) values(" & tag_G2510_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2510_start = Nothing
Dim tag_G2510_freq
tag_G2510_freq = HMIRuntime.Tags("G2510_freq").Read
strSQL = "INSERT INTO z_tag_G2510_freq (tag_value, created) values(" & tag_G2510_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2510_freq = Nothing
Dim tag_G2513_start
tag_G2513_start = HMIRuntime.Tags("G2513_start").Read
strSQL = "INSERT INTO z_tag_G2513_start (tag_value, created) values(" & tag_G2513_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2513_start = Nothing
Dim tag_G2513_freq
tag_G2513_freq = HMIRuntime.Tags("G2513_freq").Read
strSQL = "INSERT INTO z_tag_G2513_freq (tag_value, created) values(" & tag_G2513_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2513_freq = Nothing
Dim tag_G2604_start
tag_G2604_start = HMIRuntime.Tags("G2604_start").Read
strSQL = "INSERT INTO z_tag_G2604_start (tag_value, created) values(" & tag_G2604_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2604_start = Nothing
Dim tag_G2604_freq
tag_G2604_freq = HMIRuntime.Tags("G2604_freq").Read
strSQL = "INSERT INTO z_tag_G2604_freq (tag_value, created) values(" & tag_G2604_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2604_freq = Nothing
Dim tag_G2607_start
tag_G2607_start = HMIRuntime.Tags("G2607_start").Read
strSQL = "INSERT INTO z_tag_G2607_start (tag_value, created) values(" & tag_G2607_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2607_start = Nothing
Dim tag_G2607_freq
tag_G2607_freq = HMIRuntime.Tags("G2607_freq").Read
strSQL = "INSERT INTO z_tag_G2607_freq (tag_value, created) values(" & tag_G2607_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2607_freq = Nothing
Dim tag_G2609_start
tag_G2609_start = HMIRuntime.Tags("G2609_start").Read
strSQL = "INSERT INTO z_tag_G2609_start (tag_value, created) values(" & tag_G2609_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2609_start = Nothing
Dim tag_G2609_freq
tag_G2609_freq = HMIRuntime.Tags("G2609_freq").Read
strSQL = "INSERT INTO z_tag_G2609_freq (tag_value, created) values(" & tag_G2609_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2609_freq = Nothing
Dim tag_G2210_start
tag_G2210_start = HMIRuntime.Tags("G2210_start").Read
strSQL = "INSERT INTO z_tag_G2210_start (tag_value, created) values(" & tag_G2210_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2210_start = Nothing
Dim tag_G2210_freq
tag_G2210_freq = HMIRuntime.Tags("G2210_freq").Read
strSQL = "INSERT INTO z_tag_G2210_freq (tag_value, created) values(" & tag_G2210_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2210_freq = Nothing
Dim tag_G2217_start
tag_G2217_start = HMIRuntime.Tags("G2217_start").Read
strSQL = "INSERT INTO z_tag_G2217_start (tag_value, created) values(" & tag_G2217_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2217_start = Nothing
Dim tag_G2217_freq
tag_G2217_freq = HMIRuntime.Tags("G2217_freq").Read
strSQL = "INSERT INTO z_tag_G2217_freq (tag_value, created) values(" & tag_G2217_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2217_freq = Nothing
Dim tag_G2219_start
tag_G2219_start = HMIRuntime.Tags("G2219_start").Read
strSQL = "INSERT INTO z_tag_G2219_start (tag_value, created) values(" & tag_G2219_start & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2219_start = Nothing
Dim tag_G2219_freq
tag_G2219_freq = HMIRuntime.Tags("G2219_freq").Read
strSQL = "INSERT INTO z_tag_G2219_freq (tag_value, created) values(" & tag_G2219_freq & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_G2219_freq = Nothing
Dim tag_ActualSpeed_fill
tag_ActualSpeed_fill = HMIRuntime.Tags("ActualSpeed_fill").Read
strSQL = "INSERT INTO z_tag_ActualSpeed_fill (tag_value, created) values(" & tag_ActualSpeed_fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualSpeed_fill = Nothing
Dim tag_MachineOperation_fill
tag_MachineOperation_fill = HMIRuntime.Tags("MachineOperation_fill").Read
strSQL = "INSERT INTO z_tag_MachineOperation_fill (tag_value, created) values(" & tag_MachineOperation_fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineOperation_fill = Nothing
Dim tag_MachineIsON_label
tag_MachineIsON_label = HMIRuntime.Tags("MachineIsON_label").Read
strSQL = "INSERT INTO z_tag_MachineIsON_label (tag_value, created) values(" & tag_MachineIsON_label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_label = Nothing
Dim tag_ActualSpeed_label
tag_ActualSpeed_label = HMIRuntime.Tags("ActualSpeed_label").Read
strSQL = "INSERT INTO z_tag_ActualSpeed_label (tag_value, created) values(" & tag_ActualSpeed_label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualSpeed_label = Nothing
Dim tag_MachineOperation_label
tag_MachineOperation_label = HMIRuntime.Tags("MachineOperation_label").Read
strSQL = "INSERT INTO z_tag_MachineOperation_label (tag_value, created) values(" & tag_MachineOperation_label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineOperation_label = Nothing
Dim tag_MachineIsON_Ins
tag_MachineIsON_Ins = HMIRuntime.Tags("MachineIsON_Ins").Read
strSQL = "INSERT INTO z_tag_MachineIsON_Ins (tag_value, created) values(" & tag_MachineIsON_Ins & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_Ins = Nothing
Dim tag_ActualSpeed_Ins
tag_ActualSpeed_Ins = HMIRuntime.Tags("ActualSpeed_Ins").Read
strSQL = "INSERT INTO z_tag_ActualSpeed_Ins (tag_value, created) values(" & tag_ActualSpeed_Ins & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualSpeed_Ins = Nothing
Dim tag_MachineOperation_Ins
tag_MachineOperation_Ins = HMIRuntime.Tags("MachineOperation_Ins").Read
strSQL = "INSERT INTO z_tag_MachineOperation_Ins (tag_value, created) values(" & tag_MachineOperation_Ins & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineOperation_Ins = Nothing
Dim tag_MachineIsON_Past
tag_MachineIsON_Past = HMIRuntime.Tags("MachineIsON_Past").Read
strSQL = "INSERT INTO z_tag_MachineIsON_Past (tag_value, created) values(" & tag_MachineIsON_Past & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_Past = Nothing
Dim tag_MachineOperation_Wash
tag_MachineOperation_Wash = HMIRuntime.Tags("MachineOperation_Wash").Read
strSQL = "INSERT INTO z_tag_MachineOperation_Wash (tag_value, created) values(" & tag_MachineOperation_Wash & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineOperation_Wash = Nothing
Dim tag_MachineIsON_Packer
tag_MachineIsON_Packer = HMIRuntime.Tags("MachineIsON_Packer").Read
strSQL = "INSERT INTO z_tag_MachineIsON_Packer (tag_value, created) values(" & tag_MachineIsON_Packer & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_Packer = Nothing
Dim tag_MachineIsON_UnPacker
tag_MachineIsON_UnPacker = HMIRuntime.Tags("MachineIsON_UnPacker").Read
strSQL = "INSERT INTO z_tag_MachineIsON_UnPacker (tag_value, created) values(" & tag_MachineIsON_UnPacker & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MachineIsON_UnPacker = Nothing
Dim tag_Start_Time
tag_Start_Time = HMIRuntime.Tags("Start_Time").Read
strSQL = "INSERT INTO z_tag_Start_Time (tag_value, created) values(" & tag_Start_Time & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Start_Time = Nothing
Dim tag_ActualTime_Unpacker
tag_ActualTime_Unpacker = HMIRuntime.Tags("ActualTime_Unpacker").Read
strSQL = "INSERT INTO z_tag_ActualTime_Unpacker (tag_value, created) values(" & tag_ActualTime_Unpacker & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Unpacker = Nothing
Dim tag_ActualTime_Wash
tag_ActualTime_Wash = HMIRuntime.Tags("ActualTime_Wash").Read
strSQL = "INSERT INTO z_tag_ActualTime_Wash (tag_value, created) values(" & tag_ActualTime_Wash & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Wash = Nothing
Dim tag_ActualTime_MPI
tag_ActualTime_MPI = HMIRuntime.Tags("ActualTime_MPI").Read
strSQL = "INSERT INTO z_tag_ActualTime_MPI (tag_value, created) values(" & tag_ActualTime_MPI & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_MPI = Nothing
Dim tag_ActualTime_Fill
tag_ActualTime_Fill = HMIRuntime.Tags("ActualTime_Fill").Read
strSQL = "INSERT INTO z_tag_ActualTime_Fill (tag_value, created) values(" & tag_ActualTime_Fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Fill = Nothing
Dim tag_ActualTime_Label
tag_ActualTime_Label = HMIRuntime.Tags("ActualTime_Label").Read
strSQL = "INSERT INTO z_tag_ActualTime_Label (tag_value, created) values(" & tag_ActualTime_Label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Label = Nothing
Dim tag_ActualTime_Packer
tag_ActualTime_Packer = HMIRuntime.Tags("ActualTime_Packer").Read
strSQL = "INSERT INTO z_tag_ActualTime_Packer (tag_value, created) values(" & tag_ActualTime_Packer & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Packer = Nothing
Dim tag_ActualTime_Pasteu
tag_ActualTime_Pasteu = HMIRuntime.Tags("ActualTime_Pasteu").Read
strSQL = "INSERT INTO z_tag_ActualTime_Pasteu (tag_value, created) values(" & tag_ActualTime_Pasteu & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Pasteu = Nothing
Dim tag_StopTime_Fill
tag_StopTime_Fill = HMIRuntime.Tags("StopTime_Fill").Read
strSQL = "INSERT INTO z_tag_StopTime_Fill (tag_value, created) values(" & tag_StopTime_Fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Fill = Nothing
Dim tag_StopTime_Label
tag_StopTime_Label = HMIRuntime.Tags("StopTime_Label").Read
strSQL = "INSERT INTO z_tag_StopTime_Label (tag_value, created) values(" & tag_StopTime_Label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Label = Nothing
Dim tag_StopTime_Wash
tag_StopTime_Wash = HMIRuntime.Tags("StopTime_Wash").Read
strSQL = "INSERT INTO z_tag_StopTime_Wash (tag_value, created) values(" & tag_StopTime_Wash & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Wash = Nothing
Dim tag_StopTime_MPI
tag_StopTime_MPI = HMIRuntime.Tags("StopTime_MPI").Read
strSQL = "INSERT INTO z_tag_StopTime_MPI (tag_value, created) values(" & tag_StopTime_MPI & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_MPI = Nothing
Dim tag_StopTime_unpacker
tag_StopTime_unpacker = HMIRuntime.Tags("StopTime_unpacker").Read
strSQL = "INSERT INTO z_tag_StopTime_unpacker (tag_value, created) values(" & tag_StopTime_unpacker & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_unpacker = Nothing
Dim tag_StopTime_packer
tag_StopTime_packer = HMIRuntime.Tags("StopTime_packer").Read
strSQL = "INSERT INTO z_tag_StopTime_packer (tag_value, created) values(" & tag_StopTime_packer & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_packer = Nothing
Dim tag_Capacity_Unpacker
tag_Capacity_Unpacker = HMIRuntime.Tags("Capacity_Unpacker").Read
strSQL = "INSERT INTO z_tag_Capacity_Unpacker (tag_value, created) values(" & tag_Capacity_Unpacker & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Unpacker = Nothing
Dim tag_Capacity_Wash
tag_Capacity_Wash = HMIRuntime.Tags("Capacity_Wash").Read
strSQL = "INSERT INTO z_tag_Capacity_Wash (tag_value, created) values(" & tag_Capacity_Wash & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Wash = Nothing
Dim tag_Capacity_MPI
tag_Capacity_MPI = HMIRuntime.Tags("Capacity_MPI").Read
strSQL = "INSERT INTO z_tag_Capacity_MPI (tag_value, created) values(" & tag_Capacity_MPI & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_MPI = Nothing
Dim tag_Capacity_Fill
tag_Capacity_Fill = HMIRuntime.Tags("Capacity_Fill").Read
strSQL = "INSERT INTO z_tag_Capacity_Fill (tag_value, created) values(" & tag_Capacity_Fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Fill = Nothing
Dim tag_Capacity_Label
tag_Capacity_Label = HMIRuntime.Tags("Capacity_Label").Read
strSQL = "INSERT INTO z_tag_Capacity_Label (tag_value, created) values(" & tag_Capacity_Label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Label = Nothing
Dim tag_Capacity_Packer
tag_Capacity_Packer = HMIRuntime.Tags("Capacity_Packer").Read
strSQL = "INSERT INTO z_tag_Capacity_Packer (tag_value, created) values(" & tag_Capacity_Packer & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Packer = Nothing
Dim tag_Capacity_Pasteu
tag_Capacity_Pasteu = HMIRuntime.Tags("Capacity_Pasteu").Read
strSQL = "INSERT INTO z_tag_Capacity_Pasteu (tag_value, created) values(" & tag_Capacity_Pasteu & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Pasteu = Nothing
Dim tag_StopTime_Pasteu
tag_StopTime_Pasteu = HMIRuntime.Tags("StopTime_Pasteu").Read
strSQL = "INSERT INTO z_tag_StopTime_Pasteu (tag_value, created) values(" & tag_StopTime_Pasteu & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Pasteu = Nothing
Dim tag_P2201_ON
tag_P2201_ON = HMIRuntime.Tags("P2201_ON").Read
strSQL = "INSERT INTO z_tag_P2201_ON (tag_value, created) values(" & tag_P2201_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2201_ON = Nothing
Dim tag_P2202_ON
tag_P2202_ON = HMIRuntime.Tags("P2202_ON").Read
strSQL = "INSERT INTO z_tag_P2202_ON (tag_value, created) values(" & tag_P2202_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2202_ON = Nothing
Dim tag_P2203_ON
tag_P2203_ON = HMIRuntime.Tags("P2203_ON").Read
strSQL = "INSERT INTO z_tag_P2203_ON (tag_value, created) values(" & tag_P2203_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2203_ON = Nothing
Dim tag_P2204_ON
tag_P2204_ON = HMIRuntime.Tags("P2204_ON").Read
strSQL = "INSERT INTO z_tag_P2204_ON (tag_value, created) values(" & tag_P2204_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2204_ON = Nothing
Dim tag_P2205_ON
tag_P2205_ON = HMIRuntime.Tags("P2205_ON").Read
strSQL = "INSERT INTO z_tag_P2205_ON (tag_value, created) values(" & tag_P2205_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2205_ON = Nothing
Dim tag_P2209_ON
tag_P2209_ON = HMIRuntime.Tags("P2209_ON").Read
strSQL = "INSERT INTO z_tag_P2209_ON (tag_value, created) values(" & tag_P2209_ON & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_P2209_ON = Nothing
Dim tag_Automatic_Loader
tag_Automatic_Loader = HMIRuntime.Tags("Automatic_Loader").Read
strSQL = "INSERT INTO z_tag_Automatic_Loader (tag_value, created) values(" & tag_Automatic_Loader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Automatic_Loader = Nothing
Dim tag_Automatic_Unloader
tag_Automatic_Unloader = HMIRuntime.Tags("Automatic_Unloader").Read
strSQL = "INSERT INTO z_tag_Automatic_Unloader (tag_value, created) values(" & tag_Automatic_Unloader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Automatic_Unloader = Nothing
Dim tag_Output_Cap_Filler
tag_Output_Cap_Filler = HMIRuntime.Tags("Output_Cap_Filler").Read
strSQL = "INSERT INTO z_tag_Output_Cap_Filler (tag_value, created) values(" & tag_Output_Cap_Filler & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Output_Cap_Filler = Nothing
Dim tag_Output_Cap_Label
tag_Output_Cap_Label = HMIRuntime.Tags("Output_Cap_Label").Read
strSQL = "INSERT INTO z_tag_Output_Cap_Label (tag_value, created) values(" & tag_Output_Cap_Label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Output_Cap_Label = Nothing
Dim tag_Output_Cap_MPI
tag_Output_Cap_MPI = HMIRuntime.Tags("Output_Cap_MPI").Read
strSQL = "INSERT INTO z_tag_Output_Cap_MPI (tag_value, created) values(" & tag_Output_Cap_MPI & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Output_Cap_MPI = Nothing
Dim tag_BITONOFF_5s
tag_BITONOFF_5s = HMIRuntime.Tags("BITONOFF_5s").Read
strSQL = "INSERT INTO z_tag_BITONOFF_5s (tag_value, created) values(" & tag_BITONOFF_5s & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BITONOFF_5s = Nothing
Dim tag_TimeON_report
tag_TimeON_report = HMIRuntime.Tags("TimeON_report").Read
strSQL = "INSERT INTO z_tag_TimeON_report (tag_value, created) values(" & tag_TimeON_report & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TimeON_report = Nothing
Dim tag_TimeOFF_report
tag_TimeOFF_report = HMIRuntime.Tags("TimeOFF_report").Read
strSQL = "INSERT INTO z_tag_TimeOFF_report (tag_value, created) values(" & tag_TimeOFF_report & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TimeOFF_report = Nothing
Dim tag_DateTime_LastStart
tag_DateTime_LastStart = HMIRuntime.Tags("DateTime_LastStart").Read
strSQL = "INSERT INTO z_tag_DateTime_LastStart (tag_value, created) values(" & tag_DateTime_LastStart & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_DateTime_LastStart = Nothing
Dim tag_Ten_NVH
tag_Ten_NVH = HMIRuntime.Tags("Ten_NVH").Read
strSQL = "INSERT INTO z_tag_Ten_NVH (tag_value, created) values(" & tag_Ten_NVH & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Ten_NVH = Nothing
Dim tag_Bit_name
tag_Bit_name = HMIRuntime.Tags("Bit_name").Read
strSQL = "INSERT INTO z_tag_Bit_name (tag_value, created) values(" & tag_Bit_name & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Bit_name = Nothing
Dim tag_BIT_STOP
tag_BIT_STOP = HMIRuntime.Tags("BIT_STOP").Read
strSQL = "INSERT INTO z_tag_BIT_STOP (tag_value, created) values(" & tag_BIT_STOP & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BIT_STOP = Nothing
Dim tag_PASSWORD
tag_PASSWORD = HMIRuntime.Tags("PASSWORD").Read
strSQL = "INSERT INTO z_tag_PASSWORD (tag_value, created) values(" & tag_PASSWORD & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PASSWORD = Nothing
Dim tag_0005
tag_0005 = HMIRuntime.Tags("0005").Read
strSQL = "INSERT INTO z_tag_0005 (tag_value, created) values(" & tag_0005 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0005 = Nothing
Dim tag_0007
tag_0007 = HMIRuntime.Tags("0007").Read
strSQL = "INSERT INTO z_tag_0007 (tag_value, created) values(" & tag_0007 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0007 = Nothing
Dim tag_0096
tag_0096 = HMIRuntime.Tags("0096").Read
strSQL = "INSERT INTO z_tag_0096 (tag_value, created) values(" & tag_0096 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0096 = Nothing
Dim tag_0109
tag_0109 = HMIRuntime.Tags("0109").Read
strSQL = "INSERT INTO z_tag_0109 (tag_value, created) values(" & tag_0109 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0109 = Nothing
Dim tag_0110
tag_0110 = HMIRuntime.Tags("0110").Read
strSQL = "INSERT INTO z_tag_0110 (tag_value, created) values(" & tag_0110 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0110 = Nothing
Dim tag_0111
tag_0111 = HMIRuntime.Tags("0111").Read
strSQL = "INSERT INTO z_tag_0111 (tag_value, created) values(" & tag_0111 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0111 = Nothing
Dim tag_0112
tag_0112 = HMIRuntime.Tags("0112").Read
strSQL = "INSERT INTO z_tag_0112 (tag_value, created) values(" & tag_0112 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0112 = Nothing
Dim tag_0113
tag_0113 = HMIRuntime.Tags("0113").Read
strSQL = "INSERT INTO z_tag_0113 (tag_value, created) values(" & tag_0113 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0113 = Nothing
Dim tag_0114
tag_0114 = HMIRuntime.Tags("0114").Read
strSQL = "INSERT INTO z_tag_0114 (tag_value, created) values(" & tag_0114 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0114 = Nothing
Dim tag_0146
tag_0146 = HMIRuntime.Tags("0146").Read
strSQL = "INSERT INTO z_tag_0146 (tag_value, created) values(" & tag_0146 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0146 = Nothing
Dim tag_0164
tag_0164 = HMIRuntime.Tags("0164").Read
strSQL = "INSERT INTO z_tag_0164 (tag_value, created) values(" & tag_0164 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0164 = Nothing
Dim tag_0167
tag_0167 = HMIRuntime.Tags("0167").Read
strSQL = "INSERT INTO z_tag_0167 (tag_value, created) values(" & tag_0167 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0167 = Nothing
Dim tag_0183
tag_0183 = HMIRuntime.Tags("0183").Read
strSQL = "INSERT INTO z_tag_0183 (tag_value, created) values(" & tag_0183 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0183 = Nothing
Dim tag_0203
tag_0203 = HMIRuntime.Tags("0203").Read
strSQL = "INSERT INTO z_tag_0203 (tag_value, created) values(" & tag_0203 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203 = Nothing
Dim tag_0206
tag_0206 = HMIRuntime.Tags("0206").Read
strSQL = "INSERT INTO z_tag_0206 (tag_value, created) values(" & tag_0206 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206 = Nothing
Dim tag_0223
tag_0223 = HMIRuntime.Tags("0223").Read
strSQL = "INSERT INTO z_tag_0223 (tag_value, created) values(" & tag_0223 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0223 = Nothing
Dim tag_0224
tag_0224 = HMIRuntime.Tags("0224").Read
strSQL = "INSERT INTO z_tag_0224 (tag_value, created) values(" & tag_0224 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0224 = Nothing
Dim tag_0231
tag_0231 = HMIRuntime.Tags("0231").Read
strSQL = "INSERT INTO z_tag_0231 (tag_value, created) values(" & tag_0231 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0231 = Nothing
Dim tag_0232
tag_0232 = HMIRuntime.Tags("0232").Read
strSQL = "INSERT INTO z_tag_0232 (tag_value, created) values(" & tag_0232 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0232 = Nothing
Dim tag_0297
tag_0297 = HMIRuntime.Tags("0297").Read
strSQL = "INSERT INTO z_tag_0297 (tag_value, created) values(" & tag_0297 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0297 = Nothing
Dim tag_0299
tag_0299 = HMIRuntime.Tags("0299").Read
strSQL = "INSERT INTO z_tag_0299 (tag_value, created) values(" & tag_0299 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0299 = Nothing
Dim tag_0301
tag_0301 = HMIRuntime.Tags("0301").Read
strSQL = "INSERT INTO z_tag_0301 (tag_value, created) values(" & tag_0301 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301 = Nothing
Dim tag_0342
tag_0342 = HMIRuntime.Tags("0342").Read
strSQL = "INSERT INTO z_tag_0342 (tag_value, created) values(" & tag_0342 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0342 = Nothing
Dim tag_0347
tag_0347 = HMIRuntime.Tags("0347").Read
strSQL = "INSERT INTO z_tag_0347 (tag_value, created) values(" & tag_0347 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0347 = Nothing
Dim tag_0349
tag_0349 = HMIRuntime.Tags("0349").Read
strSQL = "INSERT INTO z_tag_0349 (tag_value, created) values(" & tag_0349 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0349 = Nothing
Dim tag_0351
tag_0351 = HMIRuntime.Tags("0351").Read
strSQL = "INSERT INTO z_tag_0351 (tag_value, created) values(" & tag_0351 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0351 = Nothing
Dim tag_0354
tag_0354 = HMIRuntime.Tags("0354").Read
strSQL = "INSERT INTO z_tag_0354 (tag_value, created) values(" & tag_0354 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0354 = Nothing
Dim tag_0356
tag_0356 = HMIRuntime.Tags("0356").Read
strSQL = "INSERT INTO z_tag_0356 (tag_value, created) values(" & tag_0356 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0356 = Nothing
Dim tag_0388
tag_0388 = HMIRuntime.Tags("0388").Read
strSQL = "INSERT INTO z_tag_0388 (tag_value, created) values(" & tag_0388 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0388 = Nothing
Dim tag_0416
tag_0416 = HMIRuntime.Tags("0416").Read
strSQL = "INSERT INTO z_tag_0416 (tag_value, created) values(" & tag_0416 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0416 = Nothing
Dim tag_0426
tag_0426 = HMIRuntime.Tags("0426").Read
strSQL = "INSERT INTO z_tag_0426 (tag_value, created) values(" & tag_0426 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0426 = Nothing
Dim tag_0566
tag_0566 = HMIRuntime.Tags("0566").Read
strSQL = "INSERT INTO z_tag_0566 (tag_value, created) values(" & tag_0566 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0566 = Nothing
Dim tag_Warm0001
tag_Warm0001 = HMIRuntime.Tags("Warm0001").Read
strSQL = "INSERT INTO z_tag_Warm0001 (tag_value, created) values(" & tag_Warm0001 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0001 = Nothing
Dim tag_Warm0065
tag_Warm0065 = HMIRuntime.Tags("Warm0065").Read
strSQL = "INSERT INTO z_tag_Warm0065 (tag_value, created) values(" & tag_Warm0065 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0065 = Nothing
Dim tag_Warm0069
tag_Warm0069 = HMIRuntime.Tags("Warm0069").Read
strSQL = "INSERT INTO z_tag_Warm0069 (tag_value, created) values(" & tag_Warm0069 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0069 = Nothing
Dim tag_Warm0099
tag_Warm0099 = HMIRuntime.Tags("Warm0099").Read
strSQL = "INSERT INTO z_tag_Warm0099 (tag_value, created) values(" & tag_Warm0099 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0099 = Nothing
Dim tag_Warm0100
tag_Warm0100 = HMIRuntime.Tags("Warm0100").Read
strSQL = "INSERT INTO z_tag_Warm0100 (tag_value, created) values(" & tag_Warm0100 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0100 = Nothing
Dim tag_Warm0101
tag_Warm0101 = HMIRuntime.Tags("Warm0101").Read
strSQL = "INSERT INTO z_tag_Warm0101 (tag_value, created) values(" & tag_Warm0101 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0101 = Nothing
Dim tag_Warm0104
tag_Warm0104 = HMIRuntime.Tags("Warm0104").Read
strSQL = "INSERT INTO z_tag_Warm0104 (tag_value, created) values(" & tag_Warm0104 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Warm0104 = Nothing
Dim tag_Note0006
tag_Note0006 = HMIRuntime.Tags("Note0006").Read
strSQL = "INSERT INTO z_tag_Note0006 (tag_value, created) values(" & tag_Note0006 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Note0006 = Nothing
Dim tag_Text5
tag_Text5 = HMIRuntime.Tags("Text5").Read
strSQL = "INSERT INTO z_tag_Text5 (tag_value, created) values(" & tag_Text5 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text5 = Nothing
Dim tag_Text7
tag_Text7 = HMIRuntime.Tags("Text7").Read
strSQL = "INSERT INTO z_tag_Text7 (tag_value, created) values(" & tag_Text7 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text7 = Nothing
Dim tag_Text16
tag_Text16 = HMIRuntime.Tags("Text16").Read
strSQL = "INSERT INTO z_tag_Text16 (tag_value, created) values(" & tag_Text16 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text16 = Nothing
Dim tag_Text174
tag_Text174 = HMIRuntime.Tags("Text174").Read
strSQL = "INSERT INTO z_tag_Text174 (tag_value, created) values(" & tag_Text174 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text174 = Nothing
Dim tag_Text175
tag_Text175 = HMIRuntime.Tags("Text175").Read
strSQL = "INSERT INTO z_tag_Text175 (tag_value, created) values(" & tag_Text175 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text175 = Nothing
Dim tag_Text176
tag_Text176 = HMIRuntime.Tags("Text176").Read
strSQL = "INSERT INTO z_tag_Text176 (tag_value, created) values(" & tag_Text176 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text176 = Nothing
Dim tag_Text226
tag_Text226 = HMIRuntime.Tags("Text226").Read
strSQL = "INSERT INTO z_tag_Text226 (tag_value, created) values(" & tag_Text226 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text226 = Nothing
Dim tag_Text236
tag_Text236 = HMIRuntime.Tags("Text236").Read
strSQL = "INSERT INTO z_tag_Text236 (tag_value, created) values(" & tag_Text236 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text236 = Nothing
Dim tag_Text237
tag_Text237 = HMIRuntime.Tags("Text237").Read
strSQL = "INSERT INTO z_tag_Text237 (tag_value, created) values(" & tag_Text237 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text237 = Nothing
Dim tag_Text350
tag_Text350 = HMIRuntime.Tags("Text350").Read
strSQL = "INSERT INTO z_tag_Text350 (tag_value, created) values(" & tag_Text350 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text350 = Nothing
Dim tag_Text555
tag_Text555 = HMIRuntime.Tags("Text555").Read
strSQL = "INSERT INTO z_tag_Text555 (tag_value, created) values(" & tag_Text555 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text555 = Nothing
Dim tag_Text563
tag_Text563 = HMIRuntime.Tags("Text563").Read
strSQL = "INSERT INTO z_tag_Text563 (tag_value, created) values(" & tag_Text563 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text563 = Nothing
Dim tag_Text631
tag_Text631 = HMIRuntime.Tags("Text631").Read
strSQL = "INSERT INTO z_tag_Text631 (tag_value, created) values(" & tag_Text631 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text631 = Nothing
Dim tag_Text644
tag_Text644 = HMIRuntime.Tags("Text644").Read
strSQL = "INSERT INTO z_tag_Text644 (tag_value, created) values(" & tag_Text644 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text644 = Nothing
Dim tag_Text645
tag_Text645 = HMIRuntime.Tags("Text645").Read
strSQL = "INSERT INTO z_tag_Text645 (tag_value, created) values(" & tag_Text645 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text645 = Nothing
Dim tag_Text646
tag_Text646 = HMIRuntime.Tags("Text646").Read
strSQL = "INSERT INTO z_tag_Text646 (tag_value, created) values(" & tag_Text646 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text646 = Nothing
Dim tag_Text647
tag_Text647 = HMIRuntime.Tags("Text647").Read
strSQL = "INSERT INTO z_tag_Text647 (tag_value, created) values(" & tag_Text647 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text647 = Nothing
Dim tag_Text654
tag_Text654 = HMIRuntime.Tags("Text654").Read
strSQL = "INSERT INTO z_tag_Text654 (tag_value, created) values(" & tag_Text654 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text654 = Nothing
Dim tag_Text655
tag_Text655 = HMIRuntime.Tags("Text655").Read
strSQL = "INSERT INTO z_tag_Text655 (tag_value, created) values(" & tag_Text655 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text655 = Nothing
Dim tag_Text660
tag_Text660 = HMIRuntime.Tags("Text660").Read
strSQL = "INSERT INTO z_tag_Text660 (tag_value, created) values(" & tag_Text660 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text660 = Nothing
Dim tag_Text661
tag_Text661 = HMIRuntime.Tags("Text661").Read
strSQL = "INSERT INTO z_tag_Text661 (tag_value, created) values(" & tag_Text661 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text661 = Nothing
Dim tag_Text662
tag_Text662 = HMIRuntime.Tags("Text662").Read
strSQL = "INSERT INTO z_tag_Text662 (tag_value, created) values(" & tag_Text662 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text662 = Nothing
Dim tag_Text663
tag_Text663 = HMIRuntime.Tags("Text663").Read
strSQL = "INSERT INTO z_tag_Text663 (tag_value, created) values(" & tag_Text663 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text663 = Nothing
Dim tag_Text674
tag_Text674 = HMIRuntime.Tags("Text674").Read
strSQL = "INSERT INTO z_tag_Text674 (tag_value, created) values(" & tag_Text674 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text674 = Nothing
Dim tag_Text675
tag_Text675 = HMIRuntime.Tags("Text675").Read
strSQL = "INSERT INTO z_tag_Text675 (tag_value, created) values(" & tag_Text675 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text675 = Nothing
Dim tag_Text676
tag_Text676 = HMIRuntime.Tags("Text676").Read
strSQL = "INSERT INTO z_tag_Text676 (tag_value, created) values(" & tag_Text676 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text676 = Nothing
Dim tag_Text677
tag_Text677 = HMIRuntime.Tags("Text677").Read
strSQL = "INSERT INTO z_tag_Text677 (tag_value, created) values(" & tag_Text677 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text677 = Nothing
Dim tag_Text684
tag_Text684 = HMIRuntime.Tags("Text684").Read
strSQL = "INSERT INTO z_tag_Text684 (tag_value, created) values(" & tag_Text684 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text684 = Nothing
Dim tag_Text685
tag_Text685 = HMIRuntime.Tags("Text685").Read
strSQL = "INSERT INTO z_tag_Text685 (tag_value, created) values(" & tag_Text685 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text685 = Nothing
Dim tag_Text690
tag_Text690 = HMIRuntime.Tags("Text690").Read
strSQL = "INSERT INTO z_tag_Text690 (tag_value, created) values(" & tag_Text690 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text690 = Nothing
Dim tag_Text691
tag_Text691 = HMIRuntime.Tags("Text691").Read
strSQL = "INSERT INTO z_tag_Text691 (tag_value, created) values(" & tag_Text691 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text691 = Nothing
Dim tag_Text692
tag_Text692 = HMIRuntime.Tags("Text692").Read
strSQL = "INSERT INTO z_tag_Text692 (tag_value, created) values(" & tag_Text692 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text692 = Nothing
Dim tag_Text693
tag_Text693 = HMIRuntime.Tags("Text693").Read
strSQL = "INSERT INTO z_tag_Text693 (tag_value, created) values(" & tag_Text693 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text693 = Nothing
Dim tag_Text704
tag_Text704 = HMIRuntime.Tags("Text704").Read
strSQL = "INSERT INTO z_tag_Text704 (tag_value, created) values(" & tag_Text704 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text704 = Nothing
Dim tag_Text705
tag_Text705 = HMIRuntime.Tags("Text705").Read
strSQL = "INSERT INTO z_tag_Text705 (tag_value, created) values(" & tag_Text705 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text705 = Nothing
Dim tag_Text706
tag_Text706 = HMIRuntime.Tags("Text706").Read
strSQL = "INSERT INTO z_tag_Text706 (tag_value, created) values(" & tag_Text706 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text706 = Nothing
Dim tag_Text707
tag_Text707 = HMIRuntime.Tags("Text707").Read
strSQL = "INSERT INTO z_tag_Text707 (tag_value, created) values(" & tag_Text707 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text707 = Nothing
Dim tag_Text710
tag_Text710 = HMIRuntime.Tags("Text710").Read
strSQL = "INSERT INTO z_tag_Text710 (tag_value, created) values(" & tag_Text710 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text710 = Nothing
Dim tag_Text724
tag_Text724 = HMIRuntime.Tags("Text724").Read
strSQL = "INSERT INTO z_tag_Text724 (tag_value, created) values(" & tag_Text724 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text724 = Nothing
Dim tag_Text726
tag_Text726 = HMIRuntime.Tags("Text726").Read
strSQL = "INSERT INTO z_tag_Text726 (tag_value, created) values(" & tag_Text726 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text726 = Nothing
Dim tag_Text732
tag_Text732 = HMIRuntime.Tags("Text732").Read
strSQL = "INSERT INTO z_tag_Text732 (tag_value, created) values(" & tag_Text732 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text732 = Nothing
Dim tag_Text755
tag_Text755 = HMIRuntime.Tags("Text755").Read
strSQL = "INSERT INTO z_tag_Text755 (tag_value, created) values(" & tag_Text755 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text755 = Nothing
Dim tag_Text800
tag_Text800 = HMIRuntime.Tags("Text800").Read
strSQL = "INSERT INTO z_tag_Text800 (tag_value, created) values(" & tag_Text800 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text800 = Nothing
Dim tag_Text802
tag_Text802 = HMIRuntime.Tags("Text802").Read
strSQL = "INSERT INTO z_tag_Text802 (tag_value, created) values(" & tag_Text802 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text802 = Nothing
Dim tag_Text804
tag_Text804 = HMIRuntime.Tags("Text804").Read
strSQL = "INSERT INTO z_tag_Text804 (tag_value, created) values(" & tag_Text804 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text804 = Nothing
Dim tag_Text1024
tag_Text1024 = HMIRuntime.Tags("Text1024").Read
strSQL = "INSERT INTO z_tag_Text1024 (tag_value, created) values(" & tag_Text1024 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1024 = Nothing
Dim tag_Text1025
tag_Text1025 = HMIRuntime.Tags("Text1025").Read
strSQL = "INSERT INTO z_tag_Text1025 (tag_value, created) values(" & tag_Text1025 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1025 = Nothing
Dim tag_Text1401
tag_Text1401 = HMIRuntime.Tags("Text1401").Read
strSQL = "INSERT INTO z_tag_Text1401 (tag_value, created) values(" & tag_Text1401 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1401 = Nothing
Dim tag_Text1406
tag_Text1406 = HMIRuntime.Tags("Text1406").Read
strSQL = "INSERT INTO z_tag_Text1406 (tag_value, created) values(" & tag_Text1406 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1406 = Nothing
Dim tag_Text1481
tag_Text1481 = HMIRuntime.Tags("Text1481").Read
strSQL = "INSERT INTO z_tag_Text1481 (tag_value, created) values(" & tag_Text1481 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1481 = Nothing
Dim tag_Text1503
tag_Text1503 = HMIRuntime.Tags("Text1503").Read
strSQL = "INSERT INTO z_tag_Text1503 (tag_value, created) values(" & tag_Text1503 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1503 = Nothing
Dim tag_Text1504
tag_Text1504 = HMIRuntime.Tags("Text1504").Read
strSQL = "INSERT INTO z_tag_Text1504 (tag_value, created) values(" & tag_Text1504 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1504 = Nothing
Dim tag_Text1566
tag_Text1566 = HMIRuntime.Tags("Text1566").Read
strSQL = "INSERT INTO z_tag_Text1566 (tag_value, created) values(" & tag_Text1566 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1566 = Nothing
Dim tag_Text1572
tag_Text1572 = HMIRuntime.Tags("Text1572").Read
strSQL = "INSERT INTO z_tag_Text1572 (tag_value, created) values(" & tag_Text1572 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1572 = Nothing
Dim tag_Text1622
tag_Text1622 = HMIRuntime.Tags("Text1622").Read
strSQL = "INSERT INTO z_tag_Text1622 (tag_value, created) values(" & tag_Text1622 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1622 = Nothing
Dim tag_Text1624
tag_Text1624 = HMIRuntime.Tags("Text1624").Read
strSQL = "INSERT INTO z_tag_Text1624 (tag_value, created) values(" & tag_Text1624 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1624 = Nothing
Dim tag_Text1841
tag_Text1841 = HMIRuntime.Tags("Text1841").Read
strSQL = "INSERT INTO z_tag_Text1841 (tag_value, created) values(" & tag_Text1841 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1841 = Nothing
Dim tag_Text1847
tag_Text1847 = HMIRuntime.Tags("Text1847").Read
strSQL = "INSERT INTO z_tag_Text1847 (tag_value, created) values(" & tag_Text1847 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1847 = Nothing
Dim tag_Text1863
tag_Text1863 = HMIRuntime.Tags("Text1863").Read
strSQL = "INSERT INTO z_tag_Text1863 (tag_value, created) values(" & tag_Text1863 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1863 = Nothing
Dim tag_Text1890
tag_Text1890 = HMIRuntime.Tags("Text1890").Read
strSQL = "INSERT INTO z_tag_Text1890 (tag_value, created) values(" & tag_Text1890 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Text1890 = Nothing
Dim tag_T05
tag_T05 = HMIRuntime.Tags("T05").Read
strSQL = "INSERT INTO z_tag_T05 (tag_value, created) values(" & tag_T05 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T05 = Nothing
Dim tag_T07
tag_T07 = HMIRuntime.Tags("T07").Read
strSQL = "INSERT INTO z_tag_T07 (tag_value, created) values(" & tag_T07 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T07 = Nothing
Dim tag_T121
tag_T121 = HMIRuntime.Tags("T121").Read
strSQL = "INSERT INTO z_tag_T121 (tag_value, created) values(" & tag_T121 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T121 = Nothing
Dim tag_T122
tag_T122 = HMIRuntime.Tags("T122").Read
strSQL = "INSERT INTO z_tag_T122 (tag_value, created) values(" & tag_T122 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T122 = Nothing
Dim tag_T123
tag_T123 = HMIRuntime.Tags("T123").Read
strSQL = "INSERT INTO z_tag_T123 (tag_value, created) values(" & tag_T123 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T123 = Nothing
Dim tag_T157
tag_T157 = HMIRuntime.Tags("T157").Read
strSQL = "INSERT INTO z_tag_T157 (tag_value, created) values(" & tag_T157 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T157 = Nothing
Dim tag_T172
tag_T172 = HMIRuntime.Tags("T172").Read
strSQL = "INSERT INTO z_tag_T172 (tag_value, created) values(" & tag_T172 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T172 = Nothing
Dim tag_T180
tag_T180 = HMIRuntime.Tags("T180").Read
strSQL = "INSERT INTO z_tag_T180 (tag_value, created) values(" & tag_T180 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T180 = Nothing
Dim tag_T183
tag_T183 = HMIRuntime.Tags("T183").Read
strSQL = "INSERT INTO z_tag_T183 (tag_value, created) values(" & tag_T183 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T183 = Nothing
Dim tag_T186
tag_T186 = HMIRuntime.Tags("T186").Read
strSQL = "INSERT INTO z_tag_T186 (tag_value, created) values(" & tag_T186 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T186 = Nothing
Dim tag_T191
tag_T191 = HMIRuntime.Tags("T191").Read
strSQL = "INSERT INTO z_tag_T191 (tag_value, created) values(" & tag_T191 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T191 = Nothing
Dim tag_T197
tag_T197 = HMIRuntime.Tags("T197").Read
strSQL = "INSERT INTO z_tag_T197 (tag_value, created) values(" & tag_T197 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T197 = Nothing
Dim tag_T202
tag_T202 = HMIRuntime.Tags("T202").Read
strSQL = "INSERT INTO z_tag_T202 (tag_value, created) values(" & tag_T202 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T202 = Nothing
Dim tag_T205
tag_T205 = HMIRuntime.Tags("T205").Read
strSQL = "INSERT INTO z_tag_T205 (tag_value, created) values(" & tag_T205 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T205 = Nothing
Dim tag_T336
tag_T336 = HMIRuntime.Tags("T336").Read
strSQL = "INSERT INTO z_tag_T336 (tag_value, created) values(" & tag_T336 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T336 = Nothing
Dim tag_T600
tag_T600 = HMIRuntime.Tags("T600").Read
strSQL = "INSERT INTO z_tag_T600 (tag_value, created) values(" & tag_T600 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T600 = Nothing
Dim tag_T601
tag_T601 = HMIRuntime.Tags("T601").Read
strSQL = "INSERT INTO z_tag_T601 (tag_value, created) values(" & tag_T601 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T601 = Nothing
Dim tag_T602
tag_T602 = HMIRuntime.Tags("T602").Read
strSQL = "INSERT INTO z_tag_T602 (tag_value, created) values(" & tag_T602 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T602 = Nothing
Dim tag_T603
tag_T603 = HMIRuntime.Tags("T603").Read
strSQL = "INSERT INTO z_tag_T603 (tag_value, created) values(" & tag_T603 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T603 = Nothing
Dim tag_T604
tag_T604 = HMIRuntime.Tags("T604").Read
strSQL = "INSERT INTO z_tag_T604 (tag_value, created) values(" & tag_T604 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T604 = Nothing
Dim tag_T830
tag_T830 = HMIRuntime.Tags("T830").Read
strSQL = "INSERT INTO z_tag_T830 (tag_value, created) values(" & tag_T830 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T830 = Nothing
Dim tag_T831
tag_T831 = HMIRuntime.Tags("T831").Read
strSQL = "INSERT INTO z_tag_T831 (tag_value, created) values(" & tag_T831 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T831 = Nothing
Dim tag_T832
tag_T832 = HMIRuntime.Tags("T832").Read
strSQL = "INSERT INTO z_tag_T832 (tag_value, created) values(" & tag_T832 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T832 = Nothing
Dim tag_T1401
tag_T1401 = HMIRuntime.Tags("T1401").Read
strSQL = "INSERT INTO z_tag_T1401 (tag_value, created) values(" & tag_T1401 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1401 = Nothing
Dim tag_T1481
tag_T1481 = HMIRuntime.Tags("T1481").Read
strSQL = "INSERT INTO z_tag_T1481 (tag_value, created) values(" & tag_T1481 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1481 = Nothing
Dim tag_T1491
tag_T1491 = HMIRuntime.Tags("T1491").Read
strSQL = "INSERT INTO z_tag_T1491 (tag_value, created) values(" & tag_T1491 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1491 = Nothing
Dim tag_T1492
tag_T1492 = HMIRuntime.Tags("T1492").Read
strSQL = "INSERT INTO z_tag_T1492 (tag_value, created) values(" & tag_T1492 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1492 = Nothing
Dim tag_T1501
tag_T1501 = HMIRuntime.Tags("T1501").Read
strSQL = "INSERT INTO z_tag_T1501 (tag_value, created) values(" & tag_T1501 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1501 = Nothing
Dim tag_T1504
tag_T1504 = HMIRuntime.Tags("T1504").Read
strSQL = "INSERT INTO z_tag_T1504 (tag_value, created) values(" & tag_T1504 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1504 = Nothing
Dim tag_T1505
tag_T1505 = HMIRuntime.Tags("T1505").Read
strSQL = "INSERT INTO z_tag_T1505 (tag_value, created) values(" & tag_T1505 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1505 = Nothing
Dim tag_T1841
tag_T1841 = HMIRuntime.Tags("T1841").Read
strSQL = "INSERT INTO z_tag_T1841 (tag_value, created) values(" & tag_T1841 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1841 = Nothing
Dim tag_T1842
tag_T1842 = HMIRuntime.Tags("T1842").Read
strSQL = "INSERT INTO z_tag_T1842 (tag_value, created) values(" & tag_T1842 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1842 = Nothing
Dim tag_T1895
tag_T1895 = HMIRuntime.Tags("T1895").Read
strSQL = "INSERT INTO z_tag_T1895 (tag_value, created) values(" & tag_T1895 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_T1895 = Nothing
Dim tag_ProductCount_Fill
tag_ProductCount_Fill = HMIRuntime.Tags("ProductCount_Fill").Read
strSQL = "INSERT INTO z_tag_ProductCount_Fill (tag_value, created) values(" & tag_ProductCount_Fill & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ProductCount_Fill = Nothing
Dim tag_ProductCount_Label
tag_ProductCount_Label = HMIRuntime.Tags("ProductCount_Label").Read
strSQL = "INSERT INTO z_tag_ProductCount_Label (tag_value, created) values(" & tag_ProductCount_Label & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ProductCount_Label = Nothing
Dim tag_Test_lo
tag_Test_lo = HMIRuntime.Tags("Test_lo").Read
strSQL = "INSERT INTO z_tag_Test_lo (tag_value, created) values(" & tag_Test_lo & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Test_lo = Nothing
Dim tag_Bit_loai_chai
tag_Bit_loai_chai = HMIRuntime.Tags("Bit_loai_chai").Read
strSQL = "INSERT INTO z_tag_Bit_loai_chai (tag_value, created) values(" & tag_Bit_loai_chai & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Bit_loai_chai = Nothing
Dim tag_BITONOFF_Fill_1s
tag_BITONOFF_Fill_1s = HMIRuntime.Tags("BITONOFF_Fill_1s").Read
strSQL = "INSERT INTO z_tag_BITONOFF_Fill_1s (tag_value, created) values(" & tag_BITONOFF_Fill_1s & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BITONOFF_Fill_1s = Nothing
Dim tag_Status_loader
tag_Status_loader = HMIRuntime.Tags("Status_loader").Read
strSQL = "INSERT INTO z_tag_Status_loader (tag_value, created) values(" & tag_Status_loader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_loader = Nothing
Dim tag_Status_unloader
tag_Status_unloader = HMIRuntime.Tags("Status_unloader").Read
strSQL = "INSERT INTO z_tag_Status_unloader (tag_value, created) values(" & tag_Status_unloader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_unloader = Nothing
Dim tag_ActualTime_Unloader
tag_ActualTime_Unloader = HMIRuntime.Tags("ActualTime_Unloader").Read
strSQL = "INSERT INTO z_tag_ActualTime_Unloader (tag_value, created) values(" & tag_ActualTime_Unloader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Unloader = Nothing
Dim tag_StopTime_Unloader
tag_StopTime_Unloader = HMIRuntime.Tags("StopTime_Unloader").Read
strSQL = "INSERT INTO z_tag_StopTime_Unloader (tag_value, created) values(" & tag_StopTime_Unloader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Unloader = Nothing
Dim tag_Capacity_Unloader
tag_Capacity_Unloader = HMIRuntime.Tags("Capacity_Unloader").Read
strSQL = "INSERT INTO z_tag_Capacity_Unloader (tag_value, created) values(" & tag_Capacity_Unloader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Unloader = Nothing
Dim tag_ActualTime_Loader
tag_ActualTime_Loader = HMIRuntime.Tags("ActualTime_Loader").Read
strSQL = "INSERT INTO z_tag_ActualTime_Loader (tag_value, created) values(" & tag_ActualTime_Loader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ActualTime_Loader = Nothing
Dim tag_StopTime_Loader
tag_StopTime_Loader = HMIRuntime.Tags("StopTime_Loader").Read
strSQL = "INSERT INTO z_tag_StopTime_Loader (tag_value, created) values(" & tag_StopTime_Loader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_StopTime_Loader = Nothing
Dim tag_Capacity_Loader
tag_Capacity_Loader = HMIRuntime.Tags("Capacity_Loader").Read
strSQL = "INSERT INTO z_tag_Capacity_Loader (tag_value, created) values(" & tag_Capacity_Loader & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Loader = Nothing
Dim tag_Start_Time_set
tag_Start_Time_set = HMIRuntime.Tags("Start_Time_set").Read
strSQL = "INSERT INTO z_tag_Start_Time_set (tag_value, created) values(" & tag_Start_Time_set & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Start_Time_set = Nothing
Dim tag_Capacity_Fill_2
tag_Capacity_Fill_2 = HMIRuntime.Tags("Capacity_Fill_2").Read
strSQL = "INSERT INTO z_tag_Capacity_Fill_2 (tag_value, created) values(" & tag_Capacity_Fill_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Fill_2 = Nothing
Dim tag_TEXT_NOTE
tag_TEXT_NOTE = HMIRuntime.Tags("TEXT_NOTE").Read
strSQL = "INSERT INTO z_tag_TEXT_NOTE (tag_value, created) values(" & tag_TEXT_NOTE & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TEXT_NOTE = Nothing
Dim tag_ON_Filler_machine
tag_ON_Filler_machine = HMIRuntime.Tags("ON_Filler_machine").Read
strSQL = "INSERT INTO z_tag_ON_Filler_machine (tag_value, created) values(" & tag_ON_Filler_machine & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ON_Filler_machine = Nothing
Dim tag_ON_Labeler_machine
tag_ON_Labeler_machine = HMIRuntime.Tags("ON_Labeler_machine").Read
strSQL = "INSERT INTO z_tag_ON_Labeler_machine (tag_value, created) values(" & tag_ON_Labeler_machine & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ON_Labeler_machine = Nothing
Dim tag_Capacity_Label_2
tag_Capacity_Label_2 = HMIRuntime.Tags("Capacity_Label_2").Read
strSQL = "INSERT INTO z_tag_Capacity_Label_2 (tag_value, created) values(" & tag_Capacity_Label_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Capacity_Label_2 = Nothing
Dim tag_Input_conveyor_motor_1
tag_Input_conveyor_motor_1 = HMIRuntime.Tags("Input_conveyor_motor_1").Read
strSQL = "INSERT INTO z_tag_Input_conveyor_motor_1 (tag_value, created) values(" & tag_Input_conveyor_motor_1 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Input_conveyor_motor_1 = Nothing
Dim tag_Input_conveyor_motor_2
tag_Input_conveyor_motor_2 = HMIRuntime.Tags("Input_conveyor_motor_2").Read
strSQL = "INSERT INTO z_tag_Input_conveyor_motor_2 (tag_value, created) values(" & tag_Input_conveyor_motor_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Input_conveyor_motor_2 = Nothing
Dim tag_Output_conveyor_motor
tag_Output_conveyor_motor = HMIRuntime.Tags("Output_conveyor_motor").Read
strSQL = "INSERT INTO z_tag_Output_conveyor_motor (tag_value, created) values(" & tag_Output_conveyor_motor & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Output_conveyor_motor = Nothing
Dim tag_Lack_infeed_2
tag_Lack_infeed_2 = HMIRuntime.Tags("Lack_infeed_2").Read
strSQL = "INSERT INTO z_tag_Lack_infeed_2 (tag_value, created) values(" & tag_Lack_infeed_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Lack_infeed_2 = Nothing
Dim tag_Lack_infeed_1
tag_Lack_infeed_1 = HMIRuntime.Tags("Lack_infeed_1").Read
strSQL = "INSERT INTO z_tag_Lack_infeed_1 (tag_value, created) values(" & tag_Lack_infeed_1 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Lack_infeed_1 = Nothing
Dim tag_Back_up_discharge_1
tag_Back_up_discharge_1 = HMIRuntime.Tags("Back_up_discharge_1").Read
strSQL = "INSERT INTO z_tag_Back_up_discharge_1 (tag_value, created) values(" & tag_Back_up_discharge_1 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Back_up_discharge_1 = Nothing
Dim tag_Back_up_discharge_2
tag_Back_up_discharge_2 = HMIRuntime.Tags("Back_up_discharge_2").Read
strSQL = "INSERT INTO z_tag_Back_up_discharge_2 (tag_value, created) values(" & tag_Back_up_discharge_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Back_up_discharge_2 = Nothing
Dim tag_Al_Filler_1
tag_Al_Filler_1 = HMIRuntime.Tags("Al_Filler_1").Read
strSQL = "INSERT INTO z_tag_Al_Filler_1 (tag_value, created) values(" & tag_Al_Filler_1 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_1 = Nothing
Dim tag_Al_Filler_2
tag_Al_Filler_2 = HMIRuntime.Tags("Al_Filler_2").Read
strSQL = "INSERT INTO z_tag_Al_Filler_2 (tag_value, created) values(" & tag_Al_Filler_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_2 = Nothing
Dim tag_Al_Filler_3
tag_Al_Filler_3 = HMIRuntime.Tags("Al_Filler_3").Read
strSQL = "INSERT INTO z_tag_Al_Filler_3 (tag_value, created) values(" & tag_Al_Filler_3 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_3 = Nothing
Dim tag_Al_Filler_4
tag_Al_Filler_4 = HMIRuntime.Tags("Al_Filler_4").Read
strSQL = "INSERT INTO z_tag_Al_Filler_4 (tag_value, created) values(" & tag_Al_Filler_4 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_4 = Nothing
Dim tag_Al_Filler_5
tag_Al_Filler_5 = HMIRuntime.Tags("Al_Filler_5").Read
strSQL = "INSERT INTO z_tag_Al_Filler_5 (tag_value, created) values(" & tag_Al_Filler_5 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_5 = Nothing
Dim tag_Al_Filler_6
tag_Al_Filler_6 = HMIRuntime.Tags("Al_Filler_6").Read
strSQL = "INSERT INTO z_tag_Al_Filler_6 (tag_value, created) values(" & tag_Al_Filler_6 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_6 = Nothing
Dim tag_Al_Filler_7
tag_Al_Filler_7 = HMIRuntime.Tags("Al_Filler_7").Read
strSQL = "INSERT INTO z_tag_Al_Filler_7 (tag_value, created) values(" & tag_Al_Filler_7 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_7 = Nothing
Dim tag_Al_Filler_8
tag_Al_Filler_8 = HMIRuntime.Tags("Al_Filler_8").Read
strSQL = "INSERT INTO z_tag_Al_Filler_8 (tag_value, created) values(" & tag_Al_Filler_8 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_8 = Nothing
Dim tag_Al_Filler_9
tag_Al_Filler_9 = HMIRuntime.Tags("Al_Filler_9").Read
strSQL = "INSERT INTO z_tag_Al_Filler_9 (tag_value, created) values(" & tag_Al_Filler_9 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_9 = Nothing
Dim tag_Al_Filler_10
tag_Al_Filler_10 = HMIRuntime.Tags("Al_Filler_10").Read
strSQL = "INSERT INTO z_tag_Al_Filler_10 (tag_value, created) values(" & tag_Al_Filler_10 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_10 = Nothing
Dim tag_Al_Filler_11
tag_Al_Filler_11 = HMIRuntime.Tags("Al_Filler_11").Read
strSQL = "INSERT INTO z_tag_Al_Filler_11 (tag_value, created) values(" & tag_Al_Filler_11 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_11 = Nothing
Dim tag_Al_Filler_12
tag_Al_Filler_12 = HMIRuntime.Tags("Al_Filler_12").Read
strSQL = "INSERT INTO z_tag_Al_Filler_12 (tag_value, created) values(" & tag_Al_Filler_12 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_12 = Nothing
Dim tag_Al_Filler_13
tag_Al_Filler_13 = HMIRuntime.Tags("Al_Filler_13").Read
strSQL = "INSERT INTO z_tag_Al_Filler_13 (tag_value, created) values(" & tag_Al_Filler_13 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_13 = Nothing
Dim tag_Al_Filler_14
tag_Al_Filler_14 = HMIRuntime.Tags("Al_Filler_14").Read
strSQL = "INSERT INTO z_tag_Al_Filler_14 (tag_value, created) values(" & tag_Al_Filler_14 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_14 = Nothing
Dim tag_Al_Filler_15
tag_Al_Filler_15 = HMIRuntime.Tags("Al_Filler_15").Read
strSQL = "INSERT INTO z_tag_Al_Filler_15 (tag_value, created) values(" & tag_Al_Filler_15 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_15 = Nothing
Dim tag_Al_Filler_16
tag_Al_Filler_16 = HMIRuntime.Tags("Al_Filler_16").Read
strSQL = "INSERT INTO z_tag_Al_Filler_16 (tag_value, created) values(" & tag_Al_Filler_16 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_16 = Nothing
Dim tag_Al_Filler_17
tag_Al_Filler_17 = HMIRuntime.Tags("Al_Filler_17").Read
strSQL = "INSERT INTO z_tag_Al_Filler_17 (tag_value, created) values(" & tag_Al_Filler_17 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_17 = Nothing
Dim tag_Al_Filler_18
tag_Al_Filler_18 = HMIRuntime.Tags("Al_Filler_18").Read
strSQL = "INSERT INTO z_tag_Al_Filler_18 (tag_value, created) values(" & tag_Al_Filler_18 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_18 = Nothing
Dim tag_Al_Filler_19
tag_Al_Filler_19 = HMIRuntime.Tags("Al_Filler_19").Read
strSQL = "INSERT INTO z_tag_Al_Filler_19 (tag_value, created) values(" & tag_Al_Filler_19 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_19 = Nothing
Dim tag_Al_Filler_20
tag_Al_Filler_20 = HMIRuntime.Tags("Al_Filler_20").Read
strSQL = "INSERT INTO z_tag_Al_Filler_20 (tag_value, created) values(" & tag_Al_Filler_20 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_20 = Nothing
Dim tag_Al_Filler_21
tag_Al_Filler_21 = HMIRuntime.Tags("Al_Filler_21").Read
strSQL = "INSERT INTO z_tag_Al_Filler_21 (tag_value, created) values(" & tag_Al_Filler_21 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_21 = Nothing
Dim tag_Al_Filler_22
tag_Al_Filler_22 = HMIRuntime.Tags("Al_Filler_22").Read
strSQL = "INSERT INTO z_tag_Al_Filler_22 (tag_value, created) values(" & tag_Al_Filler_22 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_22 = Nothing
Dim tag_Al_Filler_23
tag_Al_Filler_23 = HMIRuntime.Tags("Al_Filler_23").Read
strSQL = "INSERT INTO z_tag_Al_Filler_23 (tag_value, created) values(" & tag_Al_Filler_23 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_23 = Nothing
Dim tag_Al_Filler_24
tag_Al_Filler_24 = HMIRuntime.Tags("Al_Filler_24").Read
strSQL = "INSERT INTO z_tag_Al_Filler_24 (tag_value, created) values(" & tag_Al_Filler_24 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_24 = Nothing
Dim tag_Al_Filler_25
tag_Al_Filler_25 = HMIRuntime.Tags("Al_Filler_25").Read
strSQL = "INSERT INTO z_tag_Al_Filler_25 (tag_value, created) values(" & tag_Al_Filler_25 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_25 = Nothing
Dim tag_Al_Filler_26
tag_Al_Filler_26 = HMIRuntime.Tags("Al_Filler_26").Read
strSQL = "INSERT INTO z_tag_Al_Filler_26 (tag_value, created) values(" & tag_Al_Filler_26 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_26 = Nothing
Dim tag_Al_Filler_27
tag_Al_Filler_27 = HMIRuntime.Tags("Al_Filler_27").Read
strSQL = "INSERT INTO z_tag_Al_Filler_27 (tag_value, created) values(" & tag_Al_Filler_27 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_27 = Nothing
Dim tag_Al_Filler_28
tag_Al_Filler_28 = HMIRuntime.Tags("Al_Filler_28").Read
strSQL = "INSERT INTO z_tag_Al_Filler_28 (tag_value, created) values(" & tag_Al_Filler_28 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_28 = Nothing
Dim tag_Al_Filler_29
tag_Al_Filler_29 = HMIRuntime.Tags("Al_Filler_29").Read
strSQL = "INSERT INTO z_tag_Al_Filler_29 (tag_value, created) values(" & tag_Al_Filler_29 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_29 = Nothing
Dim tag_Al_Filler_30
tag_Al_Filler_30 = HMIRuntime.Tags("Al_Filler_30").Read
strSQL = "INSERT INTO z_tag_Al_Filler_30 (tag_value, created) values(" & tag_Al_Filler_30 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_30 = Nothing
Dim tag_Al_Filler_31
tag_Al_Filler_31 = HMIRuntime.Tags("Al_Filler_31").Read
strSQL = "INSERT INTO z_tag_Al_Filler_31 (tag_value, created) values(" & tag_Al_Filler_31 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_31 = Nothing
Dim tag_Al_Filler_32
tag_Al_Filler_32 = HMIRuntime.Tags("Al_Filler_32").Read
strSQL = "INSERT INTO z_tag_Al_Filler_32 (tag_value, created) values(" & tag_Al_Filler_32 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_32 = Nothing
Dim tag_Al_Filler_33
tag_Al_Filler_33 = HMIRuntime.Tags("Al_Filler_33").Read
strSQL = "INSERT INTO z_tag_Al_Filler_33 (tag_value, created) values(" & tag_Al_Filler_33 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_33 = Nothing
Dim tag_Al_Filler_34
tag_Al_Filler_34 = HMIRuntime.Tags("Al_Filler_34").Read
strSQL = "INSERT INTO z_tag_Al_Filler_34 (tag_value, created) values(" & tag_Al_Filler_34 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_34 = Nothing
Dim tag_Al_Filler_35
tag_Al_Filler_35 = HMIRuntime.Tags("Al_Filler_35").Read
strSQL = "INSERT INTO z_tag_Al_Filler_35 (tag_value, created) values(" & tag_Al_Filler_35 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_35 = Nothing
Dim tag_Al_Filler_36
tag_Al_Filler_36 = HMIRuntime.Tags("Al_Filler_36").Read
strSQL = "INSERT INTO z_tag_Al_Filler_36 (tag_value, created) values(" & tag_Al_Filler_36 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_36 = Nothing
Dim tag_Al_Filler_37
tag_Al_Filler_37 = HMIRuntime.Tags("Al_Filler_37").Read
strSQL = "INSERT INTO z_tag_Al_Filler_37 (tag_value, created) values(" & tag_Al_Filler_37 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_37 = Nothing
Dim tag_Al_Filler_38
tag_Al_Filler_38 = HMIRuntime.Tags("Al_Filler_38").Read
strSQL = "INSERT INTO z_tag_Al_Filler_38 (tag_value, created) values(" & tag_Al_Filler_38 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_38 = Nothing
Dim tag_Al_Filler_39
tag_Al_Filler_39 = HMIRuntime.Tags("Al_Filler_39").Read
strSQL = "INSERT INTO z_tag_Al_Filler_39 (tag_value, created) values(" & tag_Al_Filler_39 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_39 = Nothing
Dim tag_Al_Filler_40
tag_Al_Filler_40 = HMIRuntime.Tags("Al_Filler_40").Read
strSQL = "INSERT INTO z_tag_Al_Filler_40 (tag_value, created) values(" & tag_Al_Filler_40 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_40 = Nothing
Dim tag_Al_Filler_41
tag_Al_Filler_41 = HMIRuntime.Tags("Al_Filler_41").Read
strSQL = "INSERT INTO z_tag_Al_Filler_41 (tag_value, created) values(" & tag_Al_Filler_41 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_41 = Nothing
Dim tag_Al_Filler_42
tag_Al_Filler_42 = HMIRuntime.Tags("Al_Filler_42").Read
strSQL = "INSERT INTO z_tag_Al_Filler_42 (tag_value, created) values(" & tag_Al_Filler_42 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_42 = Nothing
Dim tag_Al_Filler_43
tag_Al_Filler_43 = HMIRuntime.Tags("Al_Filler_43").Read
strSQL = "INSERT INTO z_tag_Al_Filler_43 (tag_value, created) values(" & tag_Al_Filler_43 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_43 = Nothing
Dim tag_Al_Filler_44
tag_Al_Filler_44 = HMIRuntime.Tags("Al_Filler_44").Read
strSQL = "INSERT INTO z_tag_Al_Filler_44 (tag_value, created) values(" & tag_Al_Filler_44 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_44 = Nothing
Dim tag_Al_Filler_45
tag_Al_Filler_45 = HMIRuntime.Tags("Al_Filler_45").Read
strSQL = "INSERT INTO z_tag_Al_Filler_45 (tag_value, created) values(" & tag_Al_Filler_45 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_45 = Nothing
Dim tag_Al_Filler_46
tag_Al_Filler_46 = HMIRuntime.Tags("Al_Filler_46").Read
strSQL = "INSERT INTO z_tag_Al_Filler_46 (tag_value, created) values(" & tag_Al_Filler_46 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_46 = Nothing
Dim tag_Al_Filler_47
tag_Al_Filler_47 = HMIRuntime.Tags("Al_Filler_47").Read
strSQL = "INSERT INTO z_tag_Al_Filler_47 (tag_value, created) values(" & tag_Al_Filler_47 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_47 = Nothing
Dim tag_Al_Filler_48
tag_Al_Filler_48 = HMIRuntime.Tags("Al_Filler_48").Read
strSQL = "INSERT INTO z_tag_Al_Filler_48 (tag_value, created) values(" & tag_Al_Filler_48 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_48 = Nothing
Dim tag_Al_Filler_49
tag_Al_Filler_49 = HMIRuntime.Tags("Al_Filler_49").Read
strSQL = "INSERT INTO z_tag_Al_Filler_49 (tag_value, created) values(" & tag_Al_Filler_49 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_49 = Nothing
Dim tag_Al_Filler_50
tag_Al_Filler_50 = HMIRuntime.Tags("Al_Filler_50").Read
strSQL = "INSERT INTO z_tag_Al_Filler_50 (tag_value, created) values(" & tag_Al_Filler_50 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_50 = Nothing
Dim tag_Al_Filler_51
tag_Al_Filler_51 = HMIRuntime.Tags("Al_Filler_51").Read
strSQL = "INSERT INTO z_tag_Al_Filler_51 (tag_value, created) values(" & tag_Al_Filler_51 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_51 = Nothing
Dim tag_Al_Filler_52
tag_Al_Filler_52 = HMIRuntime.Tags("Al_Filler_52").Read
strSQL = "INSERT INTO z_tag_Al_Filler_52 (tag_value, created) values(" & tag_Al_Filler_52 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_52 = Nothing
Dim tag_Al_Filler_53
tag_Al_Filler_53 = HMIRuntime.Tags("Al_Filler_53").Read
strSQL = "INSERT INTO z_tag_Al_Filler_53 (tag_value, created) values(" & tag_Al_Filler_53 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_53 = Nothing
Dim tag_Al_Filler_54
tag_Al_Filler_54 = HMIRuntime.Tags("Al_Filler_54").Read
strSQL = "INSERT INTO z_tag_Al_Filler_54 (tag_value, created) values(" & tag_Al_Filler_54 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_54 = Nothing
Dim tag_Al_Filler_55
tag_Al_Filler_55 = HMIRuntime.Tags("Al_Filler_55").Read
strSQL = "INSERT INTO z_tag_Al_Filler_55 (tag_value, created) values(" & tag_Al_Filler_55 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_55 = Nothing
Dim tag_Al_Filler_56
tag_Al_Filler_56 = HMIRuntime.Tags("Al_Filler_56").Read
strSQL = "INSERT INTO z_tag_Al_Filler_56 (tag_value, created) values(" & tag_Al_Filler_56 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_56 = Nothing
Dim tag_Al_Filler_57
tag_Al_Filler_57 = HMIRuntime.Tags("Al_Filler_57").Read
strSQL = "INSERT INTO z_tag_Al_Filler_57 (tag_value, created) values(" & tag_Al_Filler_57 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_57 = Nothing
Dim tag_Al_Filler_58
tag_Al_Filler_58 = HMIRuntime.Tags("Al_Filler_58").Read
strSQL = "INSERT INTO z_tag_Al_Filler_58 (tag_value, created) values(" & tag_Al_Filler_58 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_58 = Nothing
Dim tag_Al_Filler_59
tag_Al_Filler_59 = HMIRuntime.Tags("Al_Filler_59").Read
strSQL = "INSERT INTO z_tag_Al_Filler_59 (tag_value, created) values(" & tag_Al_Filler_59 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_59 = Nothing
Dim tag_Al_Filler_60
tag_Al_Filler_60 = HMIRuntime.Tags("Al_Filler_60").Read
strSQL = "INSERT INTO z_tag_Al_Filler_60 (tag_value, created) values(" & tag_Al_Filler_60 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_60 = Nothing
Dim tag_Al_Filler_61
tag_Al_Filler_61 = HMIRuntime.Tags("Al_Filler_61").Read
strSQL = "INSERT INTO z_tag_Al_Filler_61 (tag_value, created) values(" & tag_Al_Filler_61 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_61 = Nothing
Dim tag_Al_Filler_62
tag_Al_Filler_62 = HMIRuntime.Tags("Al_Filler_62").Read
strSQL = "INSERT INTO z_tag_Al_Filler_62 (tag_value, created) values(" & tag_Al_Filler_62 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Filler_62 = Nothing
Dim tag_Al_Labeller_1
tag_Al_Labeller_1 = HMIRuntime.Tags("Al_Labeller_1").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_1 (tag_value, created) values(" & tag_Al_Labeller_1 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_1 = Nothing
Dim tag_Al_Labeller_2
tag_Al_Labeller_2 = HMIRuntime.Tags("Al_Labeller_2").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_2 (tag_value, created) values(" & tag_Al_Labeller_2 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_2 = Nothing
Dim tag_Al_Labeller_3
tag_Al_Labeller_3 = HMIRuntime.Tags("Al_Labeller_3").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_3 (tag_value, created) values(" & tag_Al_Labeller_3 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_3 = Nothing
Dim tag_Al_Labeller_4
tag_Al_Labeller_4 = HMIRuntime.Tags("Al_Labeller_4").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_4 (tag_value, created) values(" & tag_Al_Labeller_4 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_4 = Nothing
Dim tag_Al_Labeller_5
tag_Al_Labeller_5 = HMIRuntime.Tags("Al_Labeller_5").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_5 (tag_value, created) values(" & tag_Al_Labeller_5 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_5 = Nothing
Dim tag_Al_Labeller_6
tag_Al_Labeller_6 = HMIRuntime.Tags("Al_Labeller_6").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_6 (tag_value, created) values(" & tag_Al_Labeller_6 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_6 = Nothing
Dim tag_Al_Labeller_7
tag_Al_Labeller_7 = HMIRuntime.Tags("Al_Labeller_7").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_7 (tag_value, created) values(" & tag_Al_Labeller_7 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_7 = Nothing
Dim tag_Al_Labeller_8
tag_Al_Labeller_8 = HMIRuntime.Tags("Al_Labeller_8").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_8 (tag_value, created) values(" & tag_Al_Labeller_8 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_8 = Nothing
Dim tag_Al_Labeller_9
tag_Al_Labeller_9 = HMIRuntime.Tags("Al_Labeller_9").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_9 (tag_value, created) values(" & tag_Al_Labeller_9 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_9 = Nothing
Dim tag_Al_Labeller_10
tag_Al_Labeller_10 = HMIRuntime.Tags("Al_Labeller_10").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_10 (tag_value, created) values(" & tag_Al_Labeller_10 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_10 = Nothing
Dim tag_Al_Labeller_11
tag_Al_Labeller_11 = HMIRuntime.Tags("Al_Labeller_11").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_11 (tag_value, created) values(" & tag_Al_Labeller_11 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_11 = Nothing
Dim tag_Al_Labeller_12
tag_Al_Labeller_12 = HMIRuntime.Tags("Al_Labeller_12").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_12 (tag_value, created) values(" & tag_Al_Labeller_12 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_12 = Nothing
Dim tag_Al_Labeller_13
tag_Al_Labeller_13 = HMIRuntime.Tags("Al_Labeller_13").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_13 (tag_value, created) values(" & tag_Al_Labeller_13 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_13 = Nothing
Dim tag_Al_Labeller_14
tag_Al_Labeller_14 = HMIRuntime.Tags("Al_Labeller_14").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_14 (tag_value, created) values(" & tag_Al_Labeller_14 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_14 = Nothing
Dim tag_Al_Labeller_15
tag_Al_Labeller_15 = HMIRuntime.Tags("Al_Labeller_15").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_15 (tag_value, created) values(" & tag_Al_Labeller_15 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_15 = Nothing
Dim tag_Al_Labeller_16
tag_Al_Labeller_16 = HMIRuntime.Tags("Al_Labeller_16").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_16 (tag_value, created) values(" & tag_Al_Labeller_16 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_16 = Nothing
Dim tag_Al_Labeller_17
tag_Al_Labeller_17 = HMIRuntime.Tags("Al_Labeller_17").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_17 (tag_value, created) values(" & tag_Al_Labeller_17 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_17 = Nothing
Dim tag_Al_Labeller_18
tag_Al_Labeller_18 = HMIRuntime.Tags("Al_Labeller_18").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_18 (tag_value, created) values(" & tag_Al_Labeller_18 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_18 = Nothing
Dim tag_Al_Labeller_19
tag_Al_Labeller_19 = HMIRuntime.Tags("Al_Labeller_19").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_19 (tag_value, created) values(" & tag_Al_Labeller_19 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_19 = Nothing
Dim tag_Al_Labeller_20
tag_Al_Labeller_20 = HMIRuntime.Tags("Al_Labeller_20").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_20 (tag_value, created) values(" & tag_Al_Labeller_20 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_20 = Nothing
Dim tag_Al_Labeller_21
tag_Al_Labeller_21 = HMIRuntime.Tags("Al_Labeller_21").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_21 (tag_value, created) values(" & tag_Al_Labeller_21 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_21 = Nothing
Dim tag_Al_Labeller_22
tag_Al_Labeller_22 = HMIRuntime.Tags("Al_Labeller_22").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_22 (tag_value, created) values(" & tag_Al_Labeller_22 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_22 = Nothing
Dim tag_Al_Labeller_23
tag_Al_Labeller_23 = HMIRuntime.Tags("Al_Labeller_23").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_23 (tag_value, created) values(" & tag_Al_Labeller_23 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_23 = Nothing
Dim tag_Al_Labeller_24
tag_Al_Labeller_24 = HMIRuntime.Tags("Al_Labeller_24").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_24 (tag_value, created) values(" & tag_Al_Labeller_24 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_24 = Nothing
Dim tag_Al_Labeller_25
tag_Al_Labeller_25 = HMIRuntime.Tags("Al_Labeller_25").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_25 (tag_value, created) values(" & tag_Al_Labeller_25 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_25 = Nothing
Dim tag_Al_Labeller_26
tag_Al_Labeller_26 = HMIRuntime.Tags("Al_Labeller_26").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_26 (tag_value, created) values(" & tag_Al_Labeller_26 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_26 = Nothing
Dim tag_Al_Labeller_27
tag_Al_Labeller_27 = HMIRuntime.Tags("Al_Labeller_27").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_27 (tag_value, created) values(" & tag_Al_Labeller_27 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_27 = Nothing
Dim tag_Al_Labeller_28
tag_Al_Labeller_28 = HMIRuntime.Tags("Al_Labeller_28").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_28 (tag_value, created) values(" & tag_Al_Labeller_28 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_28 = Nothing
Dim tag_Al_Labeller_29
tag_Al_Labeller_29 = HMIRuntime.Tags("Al_Labeller_29").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_29 (tag_value, created) values(" & tag_Al_Labeller_29 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_29 = Nothing
Dim tag_Al_Labeller_30
tag_Al_Labeller_30 = HMIRuntime.Tags("Al_Labeller_30").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_30 (tag_value, created) values(" & tag_Al_Labeller_30 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_30 = Nothing
Dim tag_Al_Labeller_31
tag_Al_Labeller_31 = HMIRuntime.Tags("Al_Labeller_31").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_31 (tag_value, created) values(" & tag_Al_Labeller_31 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_31 = Nothing
Dim tag_Al_Labeller_32
tag_Al_Labeller_32 = HMIRuntime.Tags("Al_Labeller_32").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_32 (tag_value, created) values(" & tag_Al_Labeller_32 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_32 = Nothing
Dim tag_Al_Labeller_33
tag_Al_Labeller_33 = HMIRuntime.Tags("Al_Labeller_33").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_33 (tag_value, created) values(" & tag_Al_Labeller_33 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_33 = Nothing
Dim tag_Al_Labeller_34
tag_Al_Labeller_34 = HMIRuntime.Tags("Al_Labeller_34").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_34 (tag_value, created) values(" & tag_Al_Labeller_34 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_34 = Nothing
Dim tag_Al_Labeller_35
tag_Al_Labeller_35 = HMIRuntime.Tags("Al_Labeller_35").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_35 (tag_value, created) values(" & tag_Al_Labeller_35 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_35 = Nothing
Dim tag_Al_Labeller_36
tag_Al_Labeller_36 = HMIRuntime.Tags("Al_Labeller_36").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_36 (tag_value, created) values(" & tag_Al_Labeller_36 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_36 = Nothing
Dim tag_Al_Labeller_37
tag_Al_Labeller_37 = HMIRuntime.Tags("Al_Labeller_37").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_37 (tag_value, created) values(" & tag_Al_Labeller_37 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_37 = Nothing
Dim tag_Al_Labeller_38
tag_Al_Labeller_38 = HMIRuntime.Tags("Al_Labeller_38").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_38 (tag_value, created) values(" & tag_Al_Labeller_38 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_38 = Nothing
Dim tag_Al_Labeller_39
tag_Al_Labeller_39 = HMIRuntime.Tags("Al_Labeller_39").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_39 (tag_value, created) values(" & tag_Al_Labeller_39 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_39 = Nothing
Dim tag_Al_Labeller_40
tag_Al_Labeller_40 = HMIRuntime.Tags("Al_Labeller_40").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_40 (tag_value, created) values(" & tag_Al_Labeller_40 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_40 = Nothing
Dim tag_Al_Labeller_41
tag_Al_Labeller_41 = HMIRuntime.Tags("Al_Labeller_41").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_41 (tag_value, created) values(" & tag_Al_Labeller_41 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_41 = Nothing
Dim tag_Al_Labeller_42
tag_Al_Labeller_42 = HMIRuntime.Tags("Al_Labeller_42").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_42 (tag_value, created) values(" & tag_Al_Labeller_42 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_42 = Nothing
Dim tag_Al_Labeller_43
tag_Al_Labeller_43 = HMIRuntime.Tags("Al_Labeller_43").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_43 (tag_value, created) values(" & tag_Al_Labeller_43 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_43 = Nothing
Dim tag_Al_Labeller_44
tag_Al_Labeller_44 = HMIRuntime.Tags("Al_Labeller_44").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_44 (tag_value, created) values(" & tag_Al_Labeller_44 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_44 = Nothing
Dim tag_Al_Labeller_45
tag_Al_Labeller_45 = HMIRuntime.Tags("Al_Labeller_45").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_45 (tag_value, created) values(" & tag_Al_Labeller_45 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_45 = Nothing
Dim tag_Al_Labeller_46
tag_Al_Labeller_46 = HMIRuntime.Tags("Al_Labeller_46").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_46 (tag_value, created) values(" & tag_Al_Labeller_46 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_46 = Nothing
Dim tag_Al_Labeller_47
tag_Al_Labeller_47 = HMIRuntime.Tags("Al_Labeller_47").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_47 (tag_value, created) values(" & tag_Al_Labeller_47 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_47 = Nothing
Dim tag_Al_Labeller_48
tag_Al_Labeller_48 = HMIRuntime.Tags("Al_Labeller_48").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_48 (tag_value, created) values(" & tag_Al_Labeller_48 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_48 = Nothing
Dim tag_Al_Labeller_49
tag_Al_Labeller_49 = HMIRuntime.Tags("Al_Labeller_49").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_49 (tag_value, created) values(" & tag_Al_Labeller_49 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_49 = Nothing
Dim tag_Al_Labeller_50
tag_Al_Labeller_50 = HMIRuntime.Tags("Al_Labeller_50").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_50 (tag_value, created) values(" & tag_Al_Labeller_50 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_50 = Nothing
Dim tag_Al_Labeller_51
tag_Al_Labeller_51 = HMIRuntime.Tags("Al_Labeller_51").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_51 (tag_value, created) values(" & tag_Al_Labeller_51 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_51 = Nothing
Dim tag_Al_Labeller_52
tag_Al_Labeller_52 = HMIRuntime.Tags("Al_Labeller_52").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_52 (tag_value, created) values(" & tag_Al_Labeller_52 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_52 = Nothing
Dim tag_Al_Labeller_53
tag_Al_Labeller_53 = HMIRuntime.Tags("Al_Labeller_53").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_53 (tag_value, created) values(" & tag_Al_Labeller_53 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_53 = Nothing
Dim tag_Al_Labeller_54
tag_Al_Labeller_54 = HMIRuntime.Tags("Al_Labeller_54").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_54 (tag_value, created) values(" & tag_Al_Labeller_54 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_54 = Nothing
Dim tag_Al_Labeller_55
tag_Al_Labeller_55 = HMIRuntime.Tags("Al_Labeller_55").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_55 (tag_value, created) values(" & tag_Al_Labeller_55 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_55 = Nothing
Dim tag_Al_Labeller_56
tag_Al_Labeller_56 = HMIRuntime.Tags("Al_Labeller_56").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_56 (tag_value, created) values(" & tag_Al_Labeller_56 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_56 = Nothing
Dim tag_Al_Labeller_57
tag_Al_Labeller_57 = HMIRuntime.Tags("Al_Labeller_57").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_57 (tag_value, created) values(" & tag_Al_Labeller_57 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_57 = Nothing
Dim tag_Al_Labeller_58
tag_Al_Labeller_58 = HMIRuntime.Tags("Al_Labeller_58").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_58 (tag_value, created) values(" & tag_Al_Labeller_58 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_58 = Nothing
Dim tag_Al_Labeller_59
tag_Al_Labeller_59 = HMIRuntime.Tags("Al_Labeller_59").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_59 (tag_value, created) values(" & tag_Al_Labeller_59 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_59 = Nothing
Dim tag_Al_Labeller_60
tag_Al_Labeller_60 = HMIRuntime.Tags("Al_Labeller_60").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_60 (tag_value, created) values(" & tag_Al_Labeller_60 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_60 = Nothing
Dim tag_Al_Labeller_61
tag_Al_Labeller_61 = HMIRuntime.Tags("Al_Labeller_61").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_61 (tag_value, created) values(" & tag_Al_Labeller_61 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_61 = Nothing
Dim tag_Al_Labeller_62
tag_Al_Labeller_62 = HMIRuntime.Tags("Al_Labeller_62").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_62 (tag_value, created) values(" & tag_Al_Labeller_62 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_62 = Nothing
Dim tag_Al_Labeller_63
tag_Al_Labeller_63 = HMIRuntime.Tags("Al_Labeller_63").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_63 (tag_value, created) values(" & tag_Al_Labeller_63 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_63 = Nothing
Dim tag_Al_Labeller_64
tag_Al_Labeller_64 = HMIRuntime.Tags("Al_Labeller_64").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_64 (tag_value, created) values(" & tag_Al_Labeller_64 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_64 = Nothing
Dim tag_Al_Labeller_65
tag_Al_Labeller_65 = HMIRuntime.Tags("Al_Labeller_65").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_65 (tag_value, created) values(" & tag_Al_Labeller_65 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_65 = Nothing
Dim tag_Al_Labeller_66
tag_Al_Labeller_66 = HMIRuntime.Tags("Al_Labeller_66").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_66 (tag_value, created) values(" & tag_Al_Labeller_66 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_66 = Nothing
Dim tag_Al_Labeller_67
tag_Al_Labeller_67 = HMIRuntime.Tags("Al_Labeller_67").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_67 (tag_value, created) values(" & tag_Al_Labeller_67 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_67 = Nothing
Dim tag_Al_Labeller_68
tag_Al_Labeller_68 = HMIRuntime.Tags("Al_Labeller_68").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_68 (tag_value, created) values(" & tag_Al_Labeller_68 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_68 = Nothing
Dim tag_Al_Labeller_69
tag_Al_Labeller_69 = HMIRuntime.Tags("Al_Labeller_69").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_69 (tag_value, created) values(" & tag_Al_Labeller_69 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_69 = Nothing
Dim tag_Al_Labeller_70
tag_Al_Labeller_70 = HMIRuntime.Tags("Al_Labeller_70").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_70 (tag_value, created) values(" & tag_Al_Labeller_70 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_70 = Nothing
Dim tag_Al_Labeller_71
tag_Al_Labeller_71 = HMIRuntime.Tags("Al_Labeller_71").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_71 (tag_value, created) values(" & tag_Al_Labeller_71 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_71 = Nothing
Dim tag_Al_Labeller_72
tag_Al_Labeller_72 = HMIRuntime.Tags("Al_Labeller_72").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_72 (tag_value, created) values(" & tag_Al_Labeller_72 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_72 = Nothing
Dim tag_Al_Labeller_73
tag_Al_Labeller_73 = HMIRuntime.Tags("Al_Labeller_73").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_73 (tag_value, created) values(" & tag_Al_Labeller_73 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_73 = Nothing
Dim tag_Al_Labeller_74
tag_Al_Labeller_74 = HMIRuntime.Tags("Al_Labeller_74").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_74 (tag_value, created) values(" & tag_Al_Labeller_74 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_74 = Nothing
Dim tag_Al_Labeller_75
tag_Al_Labeller_75 = HMIRuntime.Tags("Al_Labeller_75").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_75 (tag_value, created) values(" & tag_Al_Labeller_75 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_75 = Nothing
Dim tag_Al_Labeller_76
tag_Al_Labeller_76 = HMIRuntime.Tags("Al_Labeller_76").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_76 (tag_value, created) values(" & tag_Al_Labeller_76 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_76 = Nothing
Dim tag_Al_Labeller_77
tag_Al_Labeller_77 = HMIRuntime.Tags("Al_Labeller_77").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_77 (tag_value, created) values(" & tag_Al_Labeller_77 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_77 = Nothing
Dim tag_Al_Labeller_78
tag_Al_Labeller_78 = HMIRuntime.Tags("Al_Labeller_78").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_78 (tag_value, created) values(" & tag_Al_Labeller_78 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_78 = Nothing
Dim tag_Al_Labeller_79
tag_Al_Labeller_79 = HMIRuntime.Tags("Al_Labeller_79").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_79 (tag_value, created) values(" & tag_Al_Labeller_79 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_79 = Nothing
Dim tag_Al_Labeller_80
tag_Al_Labeller_80 = HMIRuntime.Tags("Al_Labeller_80").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_80 (tag_value, created) values(" & tag_Al_Labeller_80 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_80 = Nothing
Dim tag_Al_Labeller_81
tag_Al_Labeller_81 = HMIRuntime.Tags("Al_Labeller_81").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_81 (tag_value, created) values(" & tag_Al_Labeller_81 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_81 = Nothing
Dim tag_Al_Labeller_82
tag_Al_Labeller_82 = HMIRuntime.Tags("Al_Labeller_82").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_82 (tag_value, created) values(" & tag_Al_Labeller_82 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_82 = Nothing
Dim tag_Al_Labeller_83
tag_Al_Labeller_83 = HMIRuntime.Tags("Al_Labeller_83").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_83 (tag_value, created) values(" & tag_Al_Labeller_83 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_83 = Nothing
Dim tag_Al_Labeller_84
tag_Al_Labeller_84 = HMIRuntime.Tags("Al_Labeller_84").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_84 (tag_value, created) values(" & tag_Al_Labeller_84 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_84 = Nothing
Dim tag_Al_Labeller_85
tag_Al_Labeller_85 = HMIRuntime.Tags("Al_Labeller_85").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_85 (tag_value, created) values(" & tag_Al_Labeller_85 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_85 = Nothing
Dim tag_Al_Labeller_86
tag_Al_Labeller_86 = HMIRuntime.Tags("Al_Labeller_86").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_86 (tag_value, created) values(" & tag_Al_Labeller_86 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_86 = Nothing
Dim tag_Al_Labeller_87
tag_Al_Labeller_87 = HMIRuntime.Tags("Al_Labeller_87").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_87 (tag_value, created) values(" & tag_Al_Labeller_87 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_87 = Nothing
Dim tag_Al_Labeller_88
tag_Al_Labeller_88 = HMIRuntime.Tags("Al_Labeller_88").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_88 (tag_value, created) values(" & tag_Al_Labeller_88 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_88 = Nothing
Dim tag_Al_Labeller_89
tag_Al_Labeller_89 = HMIRuntime.Tags("Al_Labeller_89").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_89 (tag_value, created) values(" & tag_Al_Labeller_89 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_89 = Nothing
Dim tag_Al_Labeller_90
tag_Al_Labeller_90 = HMIRuntime.Tags("Al_Labeller_90").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_90 (tag_value, created) values(" & tag_Al_Labeller_90 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_90 = Nothing
Dim tag_Al_Labeller_91
tag_Al_Labeller_91 = HMIRuntime.Tags("Al_Labeller_91").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_91 (tag_value, created) values(" & tag_Al_Labeller_91 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_91 = Nothing
Dim tag_Al_Labeller_92
tag_Al_Labeller_92 = HMIRuntime.Tags("Al_Labeller_92").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_92 (tag_value, created) values(" & tag_Al_Labeller_92 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_92 = Nothing
Dim tag_Al_Labeller_93
tag_Al_Labeller_93 = HMIRuntime.Tags("Al_Labeller_93").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_93 (tag_value, created) values(" & tag_Al_Labeller_93 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_93 = Nothing
Dim tag_Al_Labeller_94
tag_Al_Labeller_94 = HMIRuntime.Tags("Al_Labeller_94").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_94 (tag_value, created) values(" & tag_Al_Labeller_94 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_94 = Nothing
Dim tag_Al_Labeller_95
tag_Al_Labeller_95 = HMIRuntime.Tags("Al_Labeller_95").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_95 (tag_value, created) values(" & tag_Al_Labeller_95 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_95 = Nothing
Dim tag_Al_Labeller_96
tag_Al_Labeller_96 = HMIRuntime.Tags("Al_Labeller_96").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_96 (tag_value, created) values(" & tag_Al_Labeller_96 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_96 = Nothing
Dim tag_Al_Labeller_97
tag_Al_Labeller_97 = HMIRuntime.Tags("Al_Labeller_97").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_97 (tag_value, created) values(" & tag_Al_Labeller_97 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_97 = Nothing
Dim tag_Al_Labeller_98
tag_Al_Labeller_98 = HMIRuntime.Tags("Al_Labeller_98").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_98 (tag_value, created) values(" & tag_Al_Labeller_98 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_98 = Nothing
Dim tag_Al_Labeller_99
tag_Al_Labeller_99 = HMIRuntime.Tags("Al_Labeller_99").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_99 (tag_value, created) values(" & tag_Al_Labeller_99 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_99 = Nothing
Dim tag_Al_Labeller_100
tag_Al_Labeller_100 = HMIRuntime.Tags("Al_Labeller_100").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_100 (tag_value, created) values(" & tag_Al_Labeller_100 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_100 = Nothing
Dim tag_Al_Labeller_101
tag_Al_Labeller_101 = HMIRuntime.Tags("Al_Labeller_101").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_101 (tag_value, created) values(" & tag_Al_Labeller_101 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_101 = Nothing
Dim tag_Al_Labeller_102
tag_Al_Labeller_102 = HMIRuntime.Tags("Al_Labeller_102").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_102 (tag_value, created) values(" & tag_Al_Labeller_102 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_102 = Nothing
Dim tag_Al_Labeller_103
tag_Al_Labeller_103 = HMIRuntime.Tags("Al_Labeller_103").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_103 (tag_value, created) values(" & tag_Al_Labeller_103 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_103 = Nothing
Dim tag_Al_Labeller_104
tag_Al_Labeller_104 = HMIRuntime.Tags("Al_Labeller_104").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_104 (tag_value, created) values(" & tag_Al_Labeller_104 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_104 = Nothing
Dim tag_Al_Labeller_105
tag_Al_Labeller_105 = HMIRuntime.Tags("Al_Labeller_105").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_105 (tag_value, created) values(" & tag_Al_Labeller_105 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_105 = Nothing
Dim tag_Al_Labeller_106
tag_Al_Labeller_106 = HMIRuntime.Tags("Al_Labeller_106").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_106 (tag_value, created) values(" & tag_Al_Labeller_106 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_106 = Nothing
Dim tag_Al_Labeller_107
tag_Al_Labeller_107 = HMIRuntime.Tags("Al_Labeller_107").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_107 (tag_value, created) values(" & tag_Al_Labeller_107 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_107 = Nothing
Dim tag_Al_Labeller_108
tag_Al_Labeller_108 = HMIRuntime.Tags("Al_Labeller_108").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_108 (tag_value, created) values(" & tag_Al_Labeller_108 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_108 = Nothing
Dim tag_Al_Labeller_109
tag_Al_Labeller_109 = HMIRuntime.Tags("Al_Labeller_109").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_109 (tag_value, created) values(" & tag_Al_Labeller_109 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_109 = Nothing
Dim tag_Al_Labeller_110
tag_Al_Labeller_110 = HMIRuntime.Tags("Al_Labeller_110").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_110 (tag_value, created) values(" & tag_Al_Labeller_110 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_110 = Nothing
Dim tag_Al_Labeller_111
tag_Al_Labeller_111 = HMIRuntime.Tags("Al_Labeller_111").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_111 (tag_value, created) values(" & tag_Al_Labeller_111 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_111 = Nothing
Dim tag_Al_Labeller_112
tag_Al_Labeller_112 = HMIRuntime.Tags("Al_Labeller_112").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_112 (tag_value, created) values(" & tag_Al_Labeller_112 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_112 = Nothing
Dim tag_Al_Labeller_113
tag_Al_Labeller_113 = HMIRuntime.Tags("Al_Labeller_113").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_113 (tag_value, created) values(" & tag_Al_Labeller_113 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_113 = Nothing
Dim tag_Al_Labeller_114
tag_Al_Labeller_114 = HMIRuntime.Tags("Al_Labeller_114").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_114 (tag_value, created) values(" & tag_Al_Labeller_114 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_114 = Nothing
Dim tag_Al_Labeller_115
tag_Al_Labeller_115 = HMIRuntime.Tags("Al_Labeller_115").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_115 (tag_value, created) values(" & tag_Al_Labeller_115 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_115 = Nothing
Dim tag_Al_Labeller_116
tag_Al_Labeller_116 = HMIRuntime.Tags("Al_Labeller_116").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_116 (tag_value, created) values(" & tag_Al_Labeller_116 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_116 = Nothing
Dim tag_Al_Labeller_117
tag_Al_Labeller_117 = HMIRuntime.Tags("Al_Labeller_117").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_117 (tag_value, created) values(" & tag_Al_Labeller_117 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_117 = Nothing
Dim tag_Al_Labeller_118
tag_Al_Labeller_118 = HMIRuntime.Tags("Al_Labeller_118").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_118 (tag_value, created) values(" & tag_Al_Labeller_118 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_118 = Nothing
Dim tag_Al_Labeller_119
tag_Al_Labeller_119 = HMIRuntime.Tags("Al_Labeller_119").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_119 (tag_value, created) values(" & tag_Al_Labeller_119 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_119 = Nothing
Dim tag_Al_Labeller_120
tag_Al_Labeller_120 = HMIRuntime.Tags("Al_Labeller_120").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_120 (tag_value, created) values(" & tag_Al_Labeller_120 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_120 = Nothing
Dim tag_Al_Labeller_121
tag_Al_Labeller_121 = HMIRuntime.Tags("Al_Labeller_121").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_121 (tag_value, created) values(" & tag_Al_Labeller_121 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_121 = Nothing
Dim tag_Al_Labeller_122
tag_Al_Labeller_122 = HMIRuntime.Tags("Al_Labeller_122").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_122 (tag_value, created) values(" & tag_Al_Labeller_122 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_122 = Nothing
Dim tag_Al_Labeller_123
tag_Al_Labeller_123 = HMIRuntime.Tags("Al_Labeller_123").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_123 (tag_value, created) values(" & tag_Al_Labeller_123 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_123 = Nothing
Dim tag_Al_Labeller_124
tag_Al_Labeller_124 = HMIRuntime.Tags("Al_Labeller_124").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_124 (tag_value, created) values(" & tag_Al_Labeller_124 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_124 = Nothing
Dim tag_Al_Labeller_125
tag_Al_Labeller_125 = HMIRuntime.Tags("Al_Labeller_125").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_125 (tag_value, created) values(" & tag_Al_Labeller_125 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_125 = Nothing
Dim tag_Al_Labeller_126
tag_Al_Labeller_126 = HMIRuntime.Tags("Al_Labeller_126").Read
strSQL = "INSERT INTO z_tag_Al_Labeller_126 (tag_value, created) values(" & tag_Al_Labeller_126 & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Al_Labeller_126 = Nothing
Dim tag_Status_chiet
tag_Status_chiet = HMIRuntime.Tags("Status_chiet").Read
strSQL = "INSERT INTO z_tag_Status_chiet (tag_value, created) values(" & tag_Status_chiet & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_chiet = Nothing
Dim tag_Status_dan
tag_Status_dan = HMIRuntime.Tags("Status_dan").Read
strSQL = "INSERT INTO z_tag_Status_dan (tag_value, created) values(" & tag_Status_dan & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_dan = Nothing
Dim tag_Status_EBI
tag_Status_EBI = HMIRuntime.Tags("Status_EBI").Read
strSQL = "INSERT INTO z_tag_Status_EBI (tag_value, created) values(" & tag_Status_EBI & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_EBI = Nothing
Dim tag_Status_rua
tag_Status_rua = HMIRuntime.Tags("Status_rua").Read
strSQL = "INSERT INTO z_tag_Status_rua (tag_value, created) values(" & tag_Status_rua & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_rua = Nothing
Dim tag_Status_unpacker
tag_Status_unpacker = HMIRuntime.Tags("Status_unpacker").Read
strSQL = "INSERT INTO z_tag_Status_unpacker (tag_value, created) values(" & tag_Status_unpacker & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_unpacker = Nothing
Dim tag_Status_packer
tag_Status_packer = HMIRuntime.Tags("Status_packer").Read
strSQL = "INSERT INTO z_tag_Status_packer (tag_value, created) values(" & tag_Status_packer & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_packer = Nothing
Dim tag_Status_pastue
tag_Status_pastue = HMIRuntime.Tags("Status_pastue").Read
strSQL = "INSERT INTO z_tag_Status_pastue (tag_value, created) values(" & tag_Status_pastue & ", CURRENT_TIMESTAMP)"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_pastue = Nothing

objConnection.Close
Set objConnection = Nothing
End Function