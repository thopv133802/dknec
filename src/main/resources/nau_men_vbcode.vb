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

Dim tag_0101LSL20
tag_0101LSL20 = HMIRuntime.Tags("0101LSL20").Read
strSQL = "INSERT INTO z_tag_0101LSL20 (tag_value, created) values(" & tag_0101LSL20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL20 = Nothing
Dim tag_0101LSL21
tag_0101LSL21 = HMIRuntime.Tags("0101LSL21").Read
strSQL = "INSERT INTO z_tag_0101LSL21 (tag_value, created) values(" & tag_0101LSL21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL21 = Nothing
Dim tag_0101LSL22
tag_0101LSL22 = HMIRuntime.Tags("0101LSL22").Read
strSQL = "INSERT INTO z_tag_0101LSL22 (tag_value, created) values(" & tag_0101LSL22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL22 = Nothing
Dim tag_0101LSL23
tag_0101LSL23 = HMIRuntime.Tags("0101LSL23").Read
strSQL = "INSERT INTO z_tag_0101LSL23 (tag_value, created) values(" & tag_0101LSL23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL23 = Nothing
Dim tag_0101LSL24
tag_0101LSL24 = HMIRuntime.Tags("0101LSL24").Read
strSQL = "INSERT INTO z_tag_0101LSL24 (tag_value, created) values(" & tag_0101LSL24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL24 = Nothing
Dim tag_0101LSL25
tag_0101LSL25 = HMIRuntime.Tags("0101LSL25").Read
strSQL = "INSERT INTO z_tag_0101LSL25 (tag_value, created) values(" & tag_0101LSL25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL25 = Nothing
Dim tag_0101LSL26
tag_0101LSL26 = HMIRuntime.Tags("0101LSL26").Read
strSQL = "INSERT INTO z_tag_0101LSL26 (tag_value, created) values(" & tag_0101LSL26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL26 = Nothing
Dim tag_0101LSL27
tag_0101LSL27 = HMIRuntime.Tags("0101LSL27").Read
strSQL = "INSERT INTO z_tag_0101LSL27 (tag_value, created) values(" & tag_0101LSL27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSL27 = Nothing
Dim tag_0101LSM20
tag_0101LSM20 = HMIRuntime.Tags("0101LSM20").Read
strSQL = "INSERT INTO z_tag_0101LSM20 (tag_value, created) values(" & tag_0101LSM20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101LSM20 = Nothing
Dim tag_0101M01
tag_0101M01 = HMIRuntime.Tags("0101M01").Read
strSQL = "INSERT INTO z_tag_0101M01 (tag_value, created) values(" & tag_0101M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M01 = Nothing
Dim tag_0101M02
tag_0101M02 = HMIRuntime.Tags("0101M02").Read
strSQL = "INSERT INTO z_tag_0101M02 (tag_value, created) values(" & tag_0101M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M02 = Nothing
Dim tag_0101M03
tag_0101M03 = HMIRuntime.Tags("0101M03").Read
strSQL = "INSERT INTO z_tag_0101M03 (tag_value, created) values(" & tag_0101M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M03 = Nothing
Dim tag_0101M04
tag_0101M04 = HMIRuntime.Tags("0101M04").Read
strSQL = "INSERT INTO z_tag_0101M04 (tag_value, created) values(" & tag_0101M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M04 = Nothing
Dim tag_0101M05
tag_0101M05 = HMIRuntime.Tags("0101M05").Read
strSQL = "INSERT INTO z_tag_0101M05 (tag_value, created) values(" & tag_0101M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M05 = Nothing
Dim tag_0101M06
tag_0101M06 = HMIRuntime.Tags("0101M06").Read
strSQL = "INSERT INTO z_tag_0101M06 (tag_value, created) values(" & tag_0101M06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M06 = Nothing
Dim tag_0101M07
tag_0101M07 = HMIRuntime.Tags("0101M07").Read
strSQL = "INSERT INTO z_tag_0101M07 (tag_value, created) values(" & tag_0101M07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M07 = Nothing
Dim tag_0101M08
tag_0101M08 = HMIRuntime.Tags("0101M08").Read
strSQL = "INSERT INTO z_tag_0101M08 (tag_value, created) values(" & tag_0101M08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M08 = Nothing
Dim tag_0101M09
tag_0101M09 = HMIRuntime.Tags("0101M09").Read
strSQL = "INSERT INTO z_tag_0101M09 (tag_value, created) values(" & tag_0101M09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M09 = Nothing
Dim tag_0101M10
tag_0101M10 = HMIRuntime.Tags("0101M10").Read
strSQL = "INSERT INTO z_tag_0101M10 (tag_value, created) values(" & tag_0101M10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M10 = Nothing
Dim tag_0101M11
tag_0101M11 = HMIRuntime.Tags("0101M11").Read
strSQL = "INSERT INTO z_tag_0101M11 (tag_value, created) values(" & tag_0101M11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M11 = Nothing
Dim tag_0101M12
tag_0101M12 = HMIRuntime.Tags("0101M12").Read
strSQL = "INSERT INTO z_tag_0101M12 (tag_value, created) values(" & tag_0101M12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M12 = Nothing
Dim tag_0101M13
tag_0101M13 = HMIRuntime.Tags("0101M13").Read
strSQL = "INSERT INTO z_tag_0101M13 (tag_value, created) values(" & tag_0101M13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M13 = Nothing
Dim tag_0101M15
tag_0101M15 = HMIRuntime.Tags("0101M15").Read
strSQL = "INSERT INTO z_tag_0101M15 (tag_value, created) values(" & tag_0101M15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M15 = Nothing
Dim tag_0101M17
tag_0101M17 = HMIRuntime.Tags("0101M17").Read
strSQL = "INSERT INTO z_tag_0101M17 (tag_value, created) values(" & tag_0101M17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M17 = Nothing
Dim tag_0101M19
tag_0101M19 = HMIRuntime.Tags("0101M19").Read
strSQL = "INSERT INTO z_tag_0101M19 (tag_value, created) values(" & tag_0101M19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M19 = Nothing
Dim tag_0101M20
tag_0101M20 = HMIRuntime.Tags("0101M20").Read
strSQL = "INSERT INTO z_tag_0101M20 (tag_value, created) values(" & tag_0101M20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M20 = Nothing
Dim tag_0101M21
tag_0101M21 = HMIRuntime.Tags("0101M21").Read
strSQL = "INSERT INTO z_tag_0101M21 (tag_value, created) values(" & tag_0101M21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M21 = Nothing
Dim tag_0101M22
tag_0101M22 = HMIRuntime.Tags("0101M22").Read
strSQL = "INSERT INTO z_tag_0101M22 (tag_value, created) values(" & tag_0101M22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M22 = Nothing
Dim tag_0101M23
tag_0101M23 = HMIRuntime.Tags("0101M23").Read
strSQL = "INSERT INTO z_tag_0101M23 (tag_value, created) values(" & tag_0101M23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M23 = Nothing
Dim tag_0101M24
tag_0101M24 = HMIRuntime.Tags("0101M24").Read
strSQL = "INSERT INTO z_tag_0101M24 (tag_value, created) values(" & tag_0101M24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M24 = Nothing
Dim tag_0101M25
tag_0101M25 = HMIRuntime.Tags("0101M25").Read
strSQL = "INSERT INTO z_tag_0101M25 (tag_value, created) values(" & tag_0101M25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M25 = Nothing
Dim tag_0101M26
tag_0101M26 = HMIRuntime.Tags("0101M26").Read
strSQL = "INSERT INTO z_tag_0101M26 (tag_value, created) values(" & tag_0101M26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M26 = Nothing
Dim tag_0101M27
tag_0101M27 = HMIRuntime.Tags("0101M27").Read
strSQL = "INSERT INTO z_tag_0101M27 (tag_value, created) values(" & tag_0101M27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M27 = Nothing
Dim tag_0101M28
tag_0101M28 = HMIRuntime.Tags("0101M28").Read
strSQL = "INSERT INTO z_tag_0101M28 (tag_value, created) values(" & tag_0101M28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M28 = Nothing
Dim tag_0101M29
tag_0101M29 = HMIRuntime.Tags("0101M29").Read
strSQL = "INSERT INTO z_tag_0101M29 (tag_value, created) values(" & tag_0101M29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M29 = Nothing
Dim tag_0101M30
tag_0101M30 = HMIRuntime.Tags("0101M30").Read
strSQL = "INSERT INTO z_tag_0101M30 (tag_value, created) values(" & tag_0101M30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M30 = Nothing
Dim tag_0101M31
tag_0101M31 = HMIRuntime.Tags("0101M31").Read
strSQL = "INSERT INTO z_tag_0101M31 (tag_value, created) values(" & tag_0101M31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M31 = Nothing
Dim tag_0101M32
tag_0101M32 = HMIRuntime.Tags("0101M32").Read
strSQL = "INSERT INTO z_tag_0101M32 (tag_value, created) values(" & tag_0101M32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M32 = Nothing
Dim tag_0101M33
tag_0101M33 = HMIRuntime.Tags("0101M33").Read
strSQL = "INSERT INTO z_tag_0101M33 (tag_value, created) values(" & tag_0101M33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M33 = Nothing
Dim tag_0101M34
tag_0101M34 = HMIRuntime.Tags("0101M34").Read
strSQL = "INSERT INTO z_tag_0101M34 (tag_value, created) values(" & tag_0101M34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M34 = Nothing
Dim tag_0101M35
tag_0101M35 = HMIRuntime.Tags("0101M35").Read
strSQL = "INSERT INTO z_tag_0101M35 (tag_value, created) values(" & tag_0101M35 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M35 = Nothing
Dim tag_0101M36
tag_0101M36 = HMIRuntime.Tags("0101M36").Read
strSQL = "INSERT INTO z_tag_0101M36 (tag_value, created) values(" & tag_0101M36 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M36 = Nothing
Dim tag_0101M37
tag_0101M37 = HMIRuntime.Tags("0101M37").Read
strSQL = "INSERT INTO z_tag_0101M37 (tag_value, created) values(" & tag_0101M37 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M37 = Nothing
Dim tag_0101M38
tag_0101M38 = HMIRuntime.Tags("0101M38").Read
strSQL = "INSERT INTO z_tag_0101M38 (tag_value, created) values(" & tag_0101M38 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M38 = Nothing
Dim tag_0101M39
tag_0101M39 = HMIRuntime.Tags("0101M39").Read
strSQL = "INSERT INTO z_tag_0101M39 (tag_value, created) values(" & tag_0101M39 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M39 = Nothing
Dim tag_0101M40
tag_0101M40 = HMIRuntime.Tags("0101M40").Read
strSQL = "INSERT INTO z_tag_0101M40 (tag_value, created) values(" & tag_0101M40 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M40 = Nothing
Dim tag_0101M41
tag_0101M41 = HMIRuntime.Tags("0101M41").Read
strSQL = "INSERT INTO z_tag_0101M41 (tag_value, created) values(" & tag_0101M41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M41 = Nothing
Dim tag_0101M42
tag_0101M42 = HMIRuntime.Tags("0101M42").Read
strSQL = "INSERT INTO z_tag_0101M42 (tag_value, created) values(" & tag_0101M42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M42 = Nothing
Dim tag_0101M43
tag_0101M43 = HMIRuntime.Tags("0101M43").Read
strSQL = "INSERT INTO z_tag_0101M43 (tag_value, created) values(" & tag_0101M43 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M43 = Nothing
Dim tag_0101M44
tag_0101M44 = HMIRuntime.Tags("0101M44").Read
strSQL = "INSERT INTO z_tag_0101M44 (tag_value, created) values(" & tag_0101M44 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M44 = Nothing
Dim tag_0101M45
tag_0101M45 = HMIRuntime.Tags("0101M45").Read
strSQL = "INSERT INTO z_tag_0101M45 (tag_value, created) values(" & tag_0101M45 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M45 = Nothing
Dim tag_0101M46
tag_0101M46 = HMIRuntime.Tags("0101M46").Read
strSQL = "INSERT INTO z_tag_0101M46 (tag_value, created) values(" & tag_0101M46 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M46 = Nothing
Dim tag_0101M47
tag_0101M47 = HMIRuntime.Tags("0101M47").Read
strSQL = "INSERT INTO z_tag_0101M47 (tag_value, created) values(" & tag_0101M47 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M47 = Nothing
Dim tag_0101M48
tag_0101M48 = HMIRuntime.Tags("0101M48").Read
strSQL = "INSERT INTO z_tag_0101M48 (tag_value, created) values(" & tag_0101M48 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101M48 = Nothing
Dim tag_0101PV01
tag_0101PV01 = HMIRuntime.Tags("0101PV01").Read
strSQL = "INSERT INTO z_tag_0101PV01 (tag_value, created) values(" & tag_0101PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV01 = Nothing
Dim tag_0101PV03
tag_0101PV03 = HMIRuntime.Tags("0101PV03").Read
strSQL = "INSERT INTO z_tag_0101PV03 (tag_value, created) values(" & tag_0101PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV03 = Nothing
Dim tag_0101PV04
tag_0101PV04 = HMIRuntime.Tags("0101PV04").Read
strSQL = "INSERT INTO z_tag_0101PV04 (tag_value, created) values(" & tag_0101PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV04 = Nothing
Dim tag_0101PV05
tag_0101PV05 = HMIRuntime.Tags("0101PV05").Read
strSQL = "INSERT INTO z_tag_0101PV05 (tag_value, created) values(" & tag_0101PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV05 = Nothing
Dim tag_0101PV06
tag_0101PV06 = HMIRuntime.Tags("0101PV06").Read
strSQL = "INSERT INTO z_tag_0101PV06 (tag_value, created) values(" & tag_0101PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV06 = Nothing
Dim tag_0101PV07
tag_0101PV07 = HMIRuntime.Tags("0101PV07").Read
strSQL = "INSERT INTO z_tag_0101PV07 (tag_value, created) values(" & tag_0101PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV07 = Nothing
Dim tag_0101PV08
tag_0101PV08 = HMIRuntime.Tags("0101PV08").Read
strSQL = "INSERT INTO z_tag_0101PV08 (tag_value, created) values(" & tag_0101PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV08 = Nothing
Dim tag_0101PV09
tag_0101PV09 = HMIRuntime.Tags("0101PV09").Read
strSQL = "INSERT INTO z_tag_0101PV09 (tag_value, created) values(" & tag_0101PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV09 = Nothing
Dim tag_0101PV10
tag_0101PV10 = HMIRuntime.Tags("0101PV10").Read
strSQL = "INSERT INTO z_tag_0101PV10 (tag_value, created) values(" & tag_0101PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV10 = Nothing
Dim tag_0101PV12
tag_0101PV12 = HMIRuntime.Tags("0101PV12").Read
strSQL = "INSERT INTO z_tag_0101PV12 (tag_value, created) values(" & tag_0101PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV12 = Nothing
Dim tag_0101PV14
tag_0101PV14 = HMIRuntime.Tags("0101PV14").Read
strSQL = "INSERT INTO z_tag_0101PV14 (tag_value, created) values(" & tag_0101PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV14 = Nothing
Dim tag_0101PV16
tag_0101PV16 = HMIRuntime.Tags("0101PV16").Read
strSQL = "INSERT INTO z_tag_0101PV16 (tag_value, created) values(" & tag_0101PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV16 = Nothing
Dim tag_0101PV16_1
tag_0101PV16_1 = HMIRuntime.Tags("0101PV16_1").Read
strSQL = "INSERT INTO z_tag_0101PV16_1 (tag_value, created) values(" & tag_0101PV16_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV16_1 = Nothing
Dim tag_0101PV17
tag_0101PV17 = HMIRuntime.Tags("0101PV17").Read
strSQL = "INSERT INTO z_tag_0101PV17 (tag_value, created) values(" & tag_0101PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV17 = Nothing
Dim tag_0101PV18
tag_0101PV18 = HMIRuntime.Tags("0101PV18").Read
strSQL = "INSERT INTO z_tag_0101PV18 (tag_value, created) values(" & tag_0101PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101PV18 = Nothing
Dim tag_0101WET01
tag_0101WET01 = HMIRuntime.Tags("0101WET01").Read
strSQL = "INSERT INTO z_tag_0101WET01 (tag_value, created) values(" & tag_0101WET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101WET01 = Nothing
Dim tag_0101WET02
tag_0101WET02 = HMIRuntime.Tags("0101WET02").Read
strSQL = "INSERT INTO z_tag_0101WET02 (tag_value, created) values(" & tag_0101WET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0101WET02 = Nothing
Dim tag_0201LAMP01
tag_0201LAMP01 = HMIRuntime.Tags("0201LAMP01").Read
strSQL = "INSERT INTO z_tag_0201LAMP01 (tag_value, created) values(" & tag_0201LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LAMP01 = Nothing
Dim tag_0201LAMP02
tag_0201LAMP02 = HMIRuntime.Tags("0201LAMP02").Read
strSQL = "INSERT INTO z_tag_0201LAMP02 (tag_value, created) values(" & tag_0201LAMP02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LAMP02 = Nothing
Dim tag_0201LET01
tag_0201LET01 = HMIRuntime.Tags("0201LET01").Read
strSQL = "INSERT INTO z_tag_0201LET01 (tag_value, created) values(" & tag_0201LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LET01 = Nothing
Dim tag_0201LET02
tag_0201LET02 = HMIRuntime.Tags("0201LET02").Read
strSQL = "INSERT INTO z_tag_0201LET02 (tag_value, created) values(" & tag_0201LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LET02 = Nothing
Dim tag_0201LSL01
tag_0201LSL01 = HMIRuntime.Tags("0201LSL01").Read
strSQL = "INSERT INTO z_tag_0201LSL01 (tag_value, created) values(" & tag_0201LSL01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LSL01 = Nothing
Dim tag_0201LSL02
tag_0201LSL02 = HMIRuntime.Tags("0201LSL02").Read
strSQL = "INSERT INTO z_tag_0201LSL02 (tag_value, created) values(" & tag_0201LSL02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201LSL02 = Nothing
Dim tag_0201M01
tag_0201M01 = HMIRuntime.Tags("0201M01").Read
strSQL = "INSERT INTO z_tag_0201M01 (tag_value, created) values(" & tag_0201M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201M01 = Nothing
Dim tag_0201M02
tag_0201M02 = HMIRuntime.Tags("0201M02").Read
strSQL = "INSERT INTO z_tag_0201M02 (tag_value, created) values(" & tag_0201M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201M02 = Nothing
Dim tag_0201M03
tag_0201M03 = HMIRuntime.Tags("0201M03").Read
strSQL = "INSERT INTO z_tag_0201M03 (tag_value, created) values(" & tag_0201M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201M03 = Nothing
Dim tag_0201M04
tag_0201M04 = HMIRuntime.Tags("0201M04").Read
strSQL = "INSERT INTO z_tag_0201M04 (tag_value, created) values(" & tag_0201M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201M04 = Nothing
Dim tag_0201M05
tag_0201M05 = HMIRuntime.Tags("0201M05").Read
strSQL = "INSERT INTO z_tag_0201M05 (tag_value, created) values(" & tag_0201M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201M05 = Nothing
Dim tag_0201PET01
tag_0201PET01 = HMIRuntime.Tags("0201PET01").Read
strSQL = "INSERT INTO z_tag_0201PET01 (tag_value, created) values(" & tag_0201PET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PET01 = Nothing
Dim tag_0201PET02
tag_0201PET02 = HMIRuntime.Tags("0201PET02").Read
strSQL = "INSERT INTO z_tag_0201PET02 (tag_value, created) values(" & tag_0201PET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PET02 = Nothing
Dim tag_0201PV01
tag_0201PV01 = HMIRuntime.Tags("0201PV01").Read
strSQL = "INSERT INTO z_tag_0201PV01 (tag_value, created) values(" & tag_0201PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV01 = Nothing
Dim tag_0201PV02
tag_0201PV02 = HMIRuntime.Tags("0201PV02").Read
strSQL = "INSERT INTO z_tag_0201PV02 (tag_value, created) values(" & tag_0201PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV02 = Nothing
Dim tag_0201PV03
tag_0201PV03 = HMIRuntime.Tags("0201PV03").Read
strSQL = "INSERT INTO z_tag_0201PV03 (tag_value, created) values(" & tag_0201PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV03 = Nothing
Dim tag_0201PV04
tag_0201PV04 = HMIRuntime.Tags("0201PV04").Read
strSQL = "INSERT INTO z_tag_0201PV04 (tag_value, created) values(" & tag_0201PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV04 = Nothing
Dim tag_0201PV05
tag_0201PV05 = HMIRuntime.Tags("0201PV05").Read
strSQL = "INSERT INTO z_tag_0201PV05 (tag_value, created) values(" & tag_0201PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV05 = Nothing
Dim tag_0201PV06
tag_0201PV06 = HMIRuntime.Tags("0201PV06").Read
strSQL = "INSERT INTO z_tag_0201PV06 (tag_value, created) values(" & tag_0201PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV06 = Nothing
Dim tag_0201PV07
tag_0201PV07 = HMIRuntime.Tags("0201PV07").Read
strSQL = "INSERT INTO z_tag_0201PV07 (tag_value, created) values(" & tag_0201PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV07 = Nothing
Dim tag_0201PV08
tag_0201PV08 = HMIRuntime.Tags("0201PV08").Read
strSQL = "INSERT INTO z_tag_0201PV08 (tag_value, created) values(" & tag_0201PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV08 = Nothing
Dim tag_0201PV09
tag_0201PV09 = HMIRuntime.Tags("0201PV09").Read
strSQL = "INSERT INTO z_tag_0201PV09 (tag_value, created) values(" & tag_0201PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV09 = Nothing
Dim tag_0201PV10
tag_0201PV10 = HMIRuntime.Tags("0201PV10").Read
strSQL = "INSERT INTO z_tag_0201PV10 (tag_value, created) values(" & tag_0201PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV10 = Nothing
Dim tag_0201PV11
tag_0201PV11 = HMIRuntime.Tags("0201PV11").Read
strSQL = "INSERT INTO z_tag_0201PV11 (tag_value, created) values(" & tag_0201PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV11 = Nothing
Dim tag_0201PV12
tag_0201PV12 = HMIRuntime.Tags("0201PV12").Read
strSQL = "INSERT INTO z_tag_0201PV12 (tag_value, created) values(" & tag_0201PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV12 = Nothing
Dim tag_0201PV13
tag_0201PV13 = HMIRuntime.Tags("0201PV13").Read
strSQL = "INSERT INTO z_tag_0201PV13 (tag_value, created) values(" & tag_0201PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV13 = Nothing
Dim tag_0201PV14
tag_0201PV14 = HMIRuntime.Tags("0201PV14").Read
strSQL = "INSERT INTO z_tag_0201PV14 (tag_value, created) values(" & tag_0201PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV14 = Nothing
Dim tag_0201PV15
tag_0201PV15 = HMIRuntime.Tags("0201PV15").Read
strSQL = "INSERT INTO z_tag_0201PV15 (tag_value, created) values(" & tag_0201PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV15 = Nothing
Dim tag_0201PV16
tag_0201PV16 = HMIRuntime.Tags("0201PV16").Read
strSQL = "INSERT INTO z_tag_0201PV16 (tag_value, created) values(" & tag_0201PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV16 = Nothing
Dim tag_0201PV17
tag_0201PV17 = HMIRuntime.Tags("0201PV17").Read
strSQL = "INSERT INTO z_tag_0201PV17 (tag_value, created) values(" & tag_0201PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV17 = Nothing
Dim tag_0201PV18
tag_0201PV18 = HMIRuntime.Tags("0201PV18").Read
strSQL = "INSERT INTO z_tag_0201PV18 (tag_value, created) values(" & tag_0201PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV18 = Nothing
Dim tag_0201PV19
tag_0201PV19 = HMIRuntime.Tags("0201PV19").Read
strSQL = "INSERT INTO z_tag_0201PV19 (tag_value, created) values(" & tag_0201PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV19 = Nothing
Dim tag_0201PV20
tag_0201PV20 = HMIRuntime.Tags("0201PV20").Read
strSQL = "INSERT INTO z_tag_0201PV20 (tag_value, created) values(" & tag_0201PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV20 = Nothing
Dim tag_0201PV21
tag_0201PV21 = HMIRuntime.Tags("0201PV21").Read
strSQL = "INSERT INTO z_tag_0201PV21 (tag_value, created) values(" & tag_0201PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV21 = Nothing
Dim tag_0201PV22
tag_0201PV22 = HMIRuntime.Tags("0201PV22").Read
strSQL = "INSERT INTO z_tag_0201PV22 (tag_value, created) values(" & tag_0201PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV22 = Nothing
Dim tag_0201PV23
tag_0201PV23 = HMIRuntime.Tags("0201PV23").Read
strSQL = "INSERT INTO z_tag_0201PV23 (tag_value, created) values(" & tag_0201PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV23 = Nothing
Dim tag_0201PV24
tag_0201PV24 = HMIRuntime.Tags("0201PV24").Read
strSQL = "INSERT INTO z_tag_0201PV24 (tag_value, created) values(" & tag_0201PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV24 = Nothing
Dim tag_0201PV25
tag_0201PV25 = HMIRuntime.Tags("0201PV25").Read
strSQL = "INSERT INTO z_tag_0201PV25 (tag_value, created) values(" & tag_0201PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV25 = Nothing
Dim tag_0201PV26
tag_0201PV26 = HMIRuntime.Tags("0201PV26").Read
strSQL = "INSERT INTO z_tag_0201PV26 (tag_value, created) values(" & tag_0201PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201PV26 = Nothing
Dim tag_0201TET01
tag_0201TET01 = HMIRuntime.Tags("0201TET01").Read
strSQL = "INSERT INTO z_tag_0201TET01 (tag_value, created) values(" & tag_0201TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201TET01 = Nothing
Dim tag_0201TET02
tag_0201TET02 = HMIRuntime.Tags("0201TET02").Read
strSQL = "INSERT INTO z_tag_0201TET02 (tag_value, created) values(" & tag_0201TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0201TET02 = Nothing
Dim tag_0202CV01
tag_0202CV01 = HMIRuntime.Tags("0202CV01").Read
strSQL = "INSERT INTO z_tag_0202CV01 (tag_value, created) values(" & tag_0202CV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202CV01 = Nothing
Dim tag_0202CV02
tag_0202CV02 = HMIRuntime.Tags("0202CV02").Read
strSQL = "INSERT INTO z_tag_0202CV02 (tag_value, created) values(" & tag_0202CV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202CV02 = Nothing
Dim tag_0202CV04_Man
tag_0202CV04_Man = HMIRuntime.Tags("0202CV04_Man").Read
strSQL = "INSERT INTO z_tag_0202CV04_Man (tag_value, created) values(" & tag_0202CV04_Man & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202CV04_Man = Nothing
Dim tag_0202FQET01
tag_0202FQET01 = HMIRuntime.Tags("0202FQET01").Read
strSQL = "INSERT INTO z_tag_0202FQET01 (tag_value, created) values(" & tag_0202FQET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202FQET01 = Nothing
Dim tag_0202FQET02
tag_0202FQET02 = HMIRuntime.Tags("0202FQET02").Read
strSQL = "INSERT INTO z_tag_0202FQET02 (tag_value, created) values(" & tag_0202FQET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202FQET02 = Nothing
Dim tag_0202GET01
tag_0202GET01 = HMIRuntime.Tags("0202GET01").Read
strSQL = "INSERT INTO z_tag_0202GET01 (tag_value, created) values(" & tag_0202GET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202GET01 = Nothing
Dim tag_0202GS03
tag_0202GS03 = HMIRuntime.Tags("0202GS03").Read
strSQL = "INSERT INTO z_tag_0202GS03 (tag_value, created) values(" & tag_0202GS03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202GS03 = Nothing
Dim tag_0202GS04
tag_0202GS04 = HMIRuntime.Tags("0202GS04").Read
strSQL = "INSERT INTO z_tag_0202GS04 (tag_value, created) values(" & tag_0202GS04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202GS04 = Nothing
Dim tag_0202GS06
tag_0202GS06 = HMIRuntime.Tags("0202GS06").Read
strSQL = "INSERT INTO z_tag_0202GS06 (tag_value, created) values(" & tag_0202GS06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202GS06 = Nothing
Dim tag_0202GS07
tag_0202GS07 = HMIRuntime.Tags("0202GS07").Read
strSQL = "INSERT INTO z_tag_0202GS07 (tag_value, created) values(" & tag_0202GS07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202GS07 = Nothing
Dim tag_0202LAMP01
tag_0202LAMP01 = HMIRuntime.Tags("0202LAMP01").Read
strSQL = "INSERT INTO z_tag_0202LAMP01 (tag_value, created) values(" & tag_0202LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202LAMP01 = Nothing
Dim tag_0202LET01
tag_0202LET01 = HMIRuntime.Tags("0202LET01").Read
strSQL = "INSERT INTO z_tag_0202LET01 (tag_value, created) values(" & tag_0202LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202LET01 = Nothing
Dim tag_0202LSL02
tag_0202LSL02 = HMIRuntime.Tags("0202LSL02").Read
strSQL = "INSERT INTO z_tag_0202LSL02 (tag_value, created) values(" & tag_0202LSL02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202LSL02 = Nothing
Dim tag_0202M01_N
tag_0202M01_N = HMIRuntime.Tags("0202M01_N").Read
strSQL = "INSERT INTO z_tag_0202M01_N (tag_value, created) values(" & tag_0202M01_N & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M01_N = Nothing
Dim tag_0202M01_T
tag_0202M01_T = HMIRuntime.Tags("0202M01_T").Read
strSQL = "INSERT INTO z_tag_0202M01_T (tag_value, created) values(" & tag_0202M01_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M01_T = Nothing
Dim tag_0202M02_H
tag_0202M02_H = HMIRuntime.Tags("0202M02_H").Read
strSQL = "INSERT INTO z_tag_0202M02_H (tag_value, created) values(" & tag_0202M02_H & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M02_H = Nothing
Dim tag_0202M02_N
tag_0202M02_N = HMIRuntime.Tags("0202M02_N").Read
strSQL = "INSERT INTO z_tag_0202M02_N (tag_value, created) values(" & tag_0202M02_N & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M02_N = Nothing
Dim tag_0202M03
tag_0202M03 = HMIRuntime.Tags("0202M03").Read
strSQL = "INSERT INTO z_tag_0202M03 (tag_value, created) values(" & tag_0202M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M03 = Nothing
Dim tag_0202M03_N
tag_0202M03_N = HMIRuntime.Tags("0202M03_N").Read
strSQL = "INSERT INTO z_tag_0202M03_N (tag_value, created) values(" & tag_0202M03_N & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M03_N = Nothing
Dim tag_0202M03_T
tag_0202M03_T = HMIRuntime.Tags("0202M03_T").Read
strSQL = "INSERT INTO z_tag_0202M03_T (tag_value, created) values(" & tag_0202M03_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M03_T = Nothing
Dim tag_0202M04
tag_0202M04 = HMIRuntime.Tags("0202M04").Read
strSQL = "INSERT INTO z_tag_0202M04 (tag_value, created) values(" & tag_0202M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M04 = Nothing
Dim tag_0202M05
tag_0202M05 = HMIRuntime.Tags("0202M05").Read
strSQL = "INSERT INTO z_tag_0202M05 (tag_value, created) values(" & tag_0202M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M05 = Nothing
Dim tag_0202M05_Man
tag_0202M05_Man = HMIRuntime.Tags("0202M05_Man").Read
strSQL = "INSERT INTO z_tag_0202M05_Man (tag_value, created) values(" & tag_0202M05_Man & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202M05_Man = Nothing
Dim tag_0202PET01
tag_0202PET01 = HMIRuntime.Tags("0202PET01").Read
strSQL = "INSERT INTO z_tag_0202PET01 (tag_value, created) values(" & tag_0202PET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PET01 = Nothing
Dim tag_0202PV01
tag_0202PV01 = HMIRuntime.Tags("0202PV01").Read
strSQL = "INSERT INTO z_tag_0202PV01 (tag_value, created) values(" & tag_0202PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV01 = Nothing
Dim tag_0202PV01_status
tag_0202PV01_status = HMIRuntime.Tags("0202PV01_status").Read
strSQL = "INSERT INTO z_tag_0202PV01_status (tag_value, created) values(" & tag_0202PV01_status & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV01_status = Nothing
Dim tag_0202PV02
tag_0202PV02 = HMIRuntime.Tags("0202PV02").Read
strSQL = "INSERT INTO z_tag_0202PV02 (tag_value, created) values(" & tag_0202PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV02 = Nothing
Dim tag_0202PV02_status
tag_0202PV02_status = HMIRuntime.Tags("0202PV02_status").Read
strSQL = "INSERT INTO z_tag_0202PV02_status (tag_value, created) values(" & tag_0202PV02_status & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV02_status = Nothing
Dim tag_0202PV03
tag_0202PV03 = HMIRuntime.Tags("0202PV03").Read
strSQL = "INSERT INTO z_tag_0202PV03 (tag_value, created) values(" & tag_0202PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV03 = Nothing
Dim tag_0202PV04
tag_0202PV04 = HMIRuntime.Tags("0202PV04").Read
strSQL = "INSERT INTO z_tag_0202PV04 (tag_value, created) values(" & tag_0202PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV04 = Nothing
Dim tag_0202PV05
tag_0202PV05 = HMIRuntime.Tags("0202PV05").Read
strSQL = "INSERT INTO z_tag_0202PV05 (tag_value, created) values(" & tag_0202PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV05 = Nothing
Dim tag_0202PV06
tag_0202PV06 = HMIRuntime.Tags("0202PV06").Read
strSQL = "INSERT INTO z_tag_0202PV06 (tag_value, created) values(" & tag_0202PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV06 = Nothing
Dim tag_0202PV07
tag_0202PV07 = HMIRuntime.Tags("0202PV07").Read
strSQL = "INSERT INTO z_tag_0202PV07 (tag_value, created) values(" & tag_0202PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV07 = Nothing
Dim tag_0202PV08
tag_0202PV08 = HMIRuntime.Tags("0202PV08").Read
strSQL = "INSERT INTO z_tag_0202PV08 (tag_value, created) values(" & tag_0202PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV08 = Nothing
Dim tag_0202PV09
tag_0202PV09 = HMIRuntime.Tags("0202PV09").Read
strSQL = "INSERT INTO z_tag_0202PV09 (tag_value, created) values(" & tag_0202PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV09 = Nothing
Dim tag_0202PV10
tag_0202PV10 = HMIRuntime.Tags("0202PV10").Read
strSQL = "INSERT INTO z_tag_0202PV10 (tag_value, created) values(" & tag_0202PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV10 = Nothing
Dim tag_0202PV11
tag_0202PV11 = HMIRuntime.Tags("0202PV11").Read
strSQL = "INSERT INTO z_tag_0202PV11 (tag_value, created) values(" & tag_0202PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV11 = Nothing
Dim tag_0202PV12
tag_0202PV12 = HMIRuntime.Tags("0202PV12").Read
strSQL = "INSERT INTO z_tag_0202PV12 (tag_value, created) values(" & tag_0202PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV12 = Nothing
Dim tag_0202PV14
tag_0202PV14 = HMIRuntime.Tags("0202PV14").Read
strSQL = "INSERT INTO z_tag_0202PV14 (tag_value, created) values(" & tag_0202PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV14 = Nothing
Dim tag_0202PV15
tag_0202PV15 = HMIRuntime.Tags("0202PV15").Read
strSQL = "INSERT INTO z_tag_0202PV15 (tag_value, created) values(" & tag_0202PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV15 = Nothing
Dim tag_0202PV16
tag_0202PV16 = HMIRuntime.Tags("0202PV16").Read
strSQL = "INSERT INTO z_tag_0202PV16 (tag_value, created) values(" & tag_0202PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV16 = Nothing
Dim tag_0202PV17
tag_0202PV17 = HMIRuntime.Tags("0202PV17").Read
strSQL = "INSERT INTO z_tag_0202PV17 (tag_value, created) values(" & tag_0202PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV17 = Nothing
Dim tag_0202PV18
tag_0202PV18 = HMIRuntime.Tags("0202PV18").Read
strSQL = "INSERT INTO z_tag_0202PV18 (tag_value, created) values(" & tag_0202PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV18 = Nothing
Dim tag_0202PV19
tag_0202PV19 = HMIRuntime.Tags("0202PV19").Read
strSQL = "INSERT INTO z_tag_0202PV19 (tag_value, created) values(" & tag_0202PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV19 = Nothing
Dim tag_0202PV20
tag_0202PV20 = HMIRuntime.Tags("0202PV20").Read
strSQL = "INSERT INTO z_tag_0202PV20 (tag_value, created) values(" & tag_0202PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202PV20 = Nothing
Dim tag_0202QET01
tag_0202QET01 = HMIRuntime.Tags("0202QET01").Read
strSQL = "INSERT INTO z_tag_0202QET01 (tag_value, created) values(" & tag_0202QET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202QET01 = Nothing
Dim tag_0202TET01
tag_0202TET01 = HMIRuntime.Tags("0202TET01").Read
strSQL = "INSERT INTO z_tag_0202TET01 (tag_value, created) values(" & tag_0202TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202TET01 = Nothing
Dim tag_0202TET02
tag_0202TET02 = HMIRuntime.Tags("0202TET02").Read
strSQL = "INSERT INTO z_tag_0202TET02 (tag_value, created) values(" & tag_0202TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0202TET02 = Nothing
Dim tag_0203LAMP01
tag_0203LAMP01 = HMIRuntime.Tags("0203LAMP01").Read
strSQL = "INSERT INTO z_tag_0203LAMP01 (tag_value, created) values(" & tag_0203LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203LAMP01 = Nothing
Dim tag_0203LET01
tag_0203LET01 = HMIRuntime.Tags("0203LET01").Read
strSQL = "INSERT INTO z_tag_0203LET01 (tag_value, created) values(" & tag_0203LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203LET01 = Nothing
Dim tag_0203LSL01
tag_0203LSL01 = HMIRuntime.Tags("0203LSL01").Read
strSQL = "INSERT INTO z_tag_0203LSL01 (tag_value, created) values(" & tag_0203LSL01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203LSL01 = Nothing
Dim tag_0203M01
tag_0203M01 = HMIRuntime.Tags("0203M01").Read
strSQL = "INSERT INTO z_tag_0203M01 (tag_value, created) values(" & tag_0203M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203M01 = Nothing
Dim tag_0203M02
tag_0203M02 = HMIRuntime.Tags("0203M02").Read
strSQL = "INSERT INTO z_tag_0203M02 (tag_value, created) values(" & tag_0203M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203M02 = Nothing
Dim tag_0203PV01
tag_0203PV01 = HMIRuntime.Tags("0203PV01").Read
strSQL = "INSERT INTO z_tag_0203PV01 (tag_value, created) values(" & tag_0203PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV01 = Nothing
Dim tag_0203PV02
tag_0203PV02 = HMIRuntime.Tags("0203PV02").Read
strSQL = "INSERT INTO z_tag_0203PV02 (tag_value, created) values(" & tag_0203PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV02 = Nothing
Dim tag_0203PV03
tag_0203PV03 = HMIRuntime.Tags("0203PV03").Read
strSQL = "INSERT INTO z_tag_0203PV03 (tag_value, created) values(" & tag_0203PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV03 = Nothing
Dim tag_0203PV04
tag_0203PV04 = HMIRuntime.Tags("0203PV04").Read
strSQL = "INSERT INTO z_tag_0203PV04 (tag_value, created) values(" & tag_0203PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV04 = Nothing
Dim tag_0203PV05
tag_0203PV05 = HMIRuntime.Tags("0203PV05").Read
strSQL = "INSERT INTO z_tag_0203PV05 (tag_value, created) values(" & tag_0203PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV05 = Nothing
Dim tag_0203PV06
tag_0203PV06 = HMIRuntime.Tags("0203PV06").Read
strSQL = "INSERT INTO z_tag_0203PV06 (tag_value, created) values(" & tag_0203PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV06 = Nothing
Dim tag_0203PV07
tag_0203PV07 = HMIRuntime.Tags("0203PV07").Read
strSQL = "INSERT INTO z_tag_0203PV07 (tag_value, created) values(" & tag_0203PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV07 = Nothing
Dim tag_0203PV08
tag_0203PV08 = HMIRuntime.Tags("0203PV08").Read
strSQL = "INSERT INTO z_tag_0203PV08 (tag_value, created) values(" & tag_0203PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV08 = Nothing
Dim tag_0203PV09
tag_0203PV09 = HMIRuntime.Tags("0203PV09").Read
strSQL = "INSERT INTO z_tag_0203PV09 (tag_value, created) values(" & tag_0203PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV09 = Nothing
Dim tag_0203PV10
tag_0203PV10 = HMIRuntime.Tags("0203PV10").Read
strSQL = "INSERT INTO z_tag_0203PV10 (tag_value, created) values(" & tag_0203PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV10 = Nothing
Dim tag_0203PV11
tag_0203PV11 = HMIRuntime.Tags("0203PV11").Read
strSQL = "INSERT INTO z_tag_0203PV11 (tag_value, created) values(" & tag_0203PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV11 = Nothing
Dim tag_0203PV12
tag_0203PV12 = HMIRuntime.Tags("0203PV12").Read
strSQL = "INSERT INTO z_tag_0203PV12 (tag_value, created) values(" & tag_0203PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV12 = Nothing
Dim tag_0203PV13
tag_0203PV13 = HMIRuntime.Tags("0203PV13").Read
strSQL = "INSERT INTO z_tag_0203PV13 (tag_value, created) values(" & tag_0203PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203PV13 = Nothing
Dim tag_0203TET01
tag_0203TET01 = HMIRuntime.Tags("0203TET01").Read
strSQL = "INSERT INTO z_tag_0203TET01 (tag_value, created) values(" & tag_0203TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203TET01 = Nothing
Dim tag_0203TET02
tag_0203TET02 = HMIRuntime.Tags("0203TET02").Read
strSQL = "INSERT INTO z_tag_0203TET02 (tag_value, created) values(" & tag_0203TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0203TET02 = Nothing
Dim tag_0204_1LET01
tag_0204_1LET01 = HMIRuntime.Tags("0204_1LET01").Read
strSQL = "INSERT INTO z_tag_0204_1LET01 (tag_value, created) values(" & tag_0204_1LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204_1LET01 = Nothing
Dim tag_0204_1PV26
tag_0204_1PV26 = HMIRuntime.Tags("0204_1PV26").Read
strSQL = "INSERT INTO z_tag_0204_1PV26 (tag_value, created) values(" & tag_0204_1PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204_1PV26 = Nothing
Dim tag_0204_2PV08
tag_0204_2PV08 = HMIRuntime.Tags("0204_2PV08").Read
strSQL = "INSERT INTO z_tag_0204_2PV08 (tag_value, created) values(" & tag_0204_2PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204_2PV08 = Nothing
Dim tag_0204LAMP01
tag_0204LAMP01 = HMIRuntime.Tags("0204LAMP01").Read
strSQL = "INSERT INTO z_tag_0204LAMP01 (tag_value, created) values(" & tag_0204LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204LAMP01 = Nothing
Dim tag_0204LET01
tag_0204LET01 = HMIRuntime.Tags("0204LET01").Read
strSQL = "INSERT INTO z_tag_0204LET01 (tag_value, created) values(" & tag_0204LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204LET01 = Nothing
Dim tag_0204LSL01
tag_0204LSL01 = HMIRuntime.Tags("0204LSL01").Read
strSQL = "INSERT INTO z_tag_0204LSL01 (tag_value, created) values(" & tag_0204LSL01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204LSL01 = Nothing
Dim tag_0204M01
tag_0204M01 = HMIRuntime.Tags("0204M01").Read
strSQL = "INSERT INTO z_tag_0204M01 (tag_value, created) values(" & tag_0204M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204M01 = Nothing
Dim tag_0204M02
tag_0204M02 = HMIRuntime.Tags("0204M02").Read
strSQL = "INSERT INTO z_tag_0204M02 (tag_value, created) values(" & tag_0204M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204M02 = Nothing
Dim tag_0204PV01
tag_0204PV01 = HMIRuntime.Tags("0204PV01").Read
strSQL = "INSERT INTO z_tag_0204PV01 (tag_value, created) values(" & tag_0204PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV01 = Nothing
Dim tag_0204PV02
tag_0204PV02 = HMIRuntime.Tags("0204PV02").Read
strSQL = "INSERT INTO z_tag_0204PV02 (tag_value, created) values(" & tag_0204PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV02 = Nothing
Dim tag_0204PV03
tag_0204PV03 = HMIRuntime.Tags("0204PV03").Read
strSQL = "INSERT INTO z_tag_0204PV03 (tag_value, created) values(" & tag_0204PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV03 = Nothing
Dim tag_0204PV04
tag_0204PV04 = HMIRuntime.Tags("0204PV04").Read
strSQL = "INSERT INTO z_tag_0204PV04 (tag_value, created) values(" & tag_0204PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV04 = Nothing
Dim tag_0204PV05
tag_0204PV05 = HMIRuntime.Tags("0204PV05").Read
strSQL = "INSERT INTO z_tag_0204PV05 (tag_value, created) values(" & tag_0204PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV05 = Nothing
Dim tag_0204PV06
tag_0204PV06 = HMIRuntime.Tags("0204PV06").Read
strSQL = "INSERT INTO z_tag_0204PV06 (tag_value, created) values(" & tag_0204PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV06 = Nothing
Dim tag_0204PV07
tag_0204PV07 = HMIRuntime.Tags("0204PV07").Read
strSQL = "INSERT INTO z_tag_0204PV07 (tag_value, created) values(" & tag_0204PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV07 = Nothing
Dim tag_0204PV08
tag_0204PV08 = HMIRuntime.Tags("0204PV08").Read
strSQL = "INSERT INTO z_tag_0204PV08 (tag_value, created) values(" & tag_0204PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV08 = Nothing
Dim tag_0204PV09
tag_0204PV09 = HMIRuntime.Tags("0204PV09").Read
strSQL = "INSERT INTO z_tag_0204PV09 (tag_value, created) values(" & tag_0204PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV09 = Nothing
Dim tag_0204PV10
tag_0204PV10 = HMIRuntime.Tags("0204PV10").Read
strSQL = "INSERT INTO z_tag_0204PV10 (tag_value, created) values(" & tag_0204PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV10 = Nothing
Dim tag_0204PV11
tag_0204PV11 = HMIRuntime.Tags("0204PV11").Read
strSQL = "INSERT INTO z_tag_0204PV11 (tag_value, created) values(" & tag_0204PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV11 = Nothing
Dim tag_0204PV12
tag_0204PV12 = HMIRuntime.Tags("0204PV12").Read
strSQL = "INSERT INTO z_tag_0204PV12 (tag_value, created) values(" & tag_0204PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV12 = Nothing
Dim tag_0204PV13
tag_0204PV13 = HMIRuntime.Tags("0204PV13").Read
strSQL = "INSERT INTO z_tag_0204PV13 (tag_value, created) values(" & tag_0204PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV13 = Nothing
Dim tag_0204PV14
tag_0204PV14 = HMIRuntime.Tags("0204PV14").Read
strSQL = "INSERT INTO z_tag_0204PV14 (tag_value, created) values(" & tag_0204PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV14 = Nothing
Dim tag_0204PV15
tag_0204PV15 = HMIRuntime.Tags("0204PV15").Read
strSQL = "INSERT INTO z_tag_0204PV15 (tag_value, created) values(" & tag_0204PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV15 = Nothing
Dim tag_0204PV16
tag_0204PV16 = HMIRuntime.Tags("0204PV16").Read
strSQL = "INSERT INTO z_tag_0204PV16 (tag_value, created) values(" & tag_0204PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV16 = Nothing
Dim tag_0204PV17
tag_0204PV17 = HMIRuntime.Tags("0204PV17").Read
strSQL = "INSERT INTO z_tag_0204PV17 (tag_value, created) values(" & tag_0204PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV17 = Nothing
Dim tag_0204PV18
tag_0204PV18 = HMIRuntime.Tags("0204PV18").Read
strSQL = "INSERT INTO z_tag_0204PV18 (tag_value, created) values(" & tag_0204PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV18 = Nothing
Dim tag_0204PV19
tag_0204PV19 = HMIRuntime.Tags("0204PV19").Read
strSQL = "INSERT INTO z_tag_0204PV19 (tag_value, created) values(" & tag_0204PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV19 = Nothing
Dim tag_0204PV20
tag_0204PV20 = HMIRuntime.Tags("0204PV20").Read
strSQL = "INSERT INTO z_tag_0204PV20 (tag_value, created) values(" & tag_0204PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV20 = Nothing
Dim tag_0204PV21
tag_0204PV21 = HMIRuntime.Tags("0204PV21").Read
strSQL = "INSERT INTO z_tag_0204PV21 (tag_value, created) values(" & tag_0204PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV21 = Nothing
Dim tag_0204PV22
tag_0204PV22 = HMIRuntime.Tags("0204PV22").Read
strSQL = "INSERT INTO z_tag_0204PV22 (tag_value, created) values(" & tag_0204PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV22 = Nothing
Dim tag_0204PV23
tag_0204PV23 = HMIRuntime.Tags("0204PV23").Read
strSQL = "INSERT INTO z_tag_0204PV23 (tag_value, created) values(" & tag_0204PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV23 = Nothing
Dim tag_0204PV24
tag_0204PV24 = HMIRuntime.Tags("0204PV24").Read
strSQL = "INSERT INTO z_tag_0204PV24 (tag_value, created) values(" & tag_0204PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV24 = Nothing
Dim tag_0204PV25
tag_0204PV25 = HMIRuntime.Tags("0204PV25").Read
strSQL = "INSERT INTO z_tag_0204PV25 (tag_value, created) values(" & tag_0204PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV25 = Nothing
Dim tag_0204PV27
tag_0204PV27 = HMIRuntime.Tags("0204PV27").Read
strSQL = "INSERT INTO z_tag_0204PV27 (tag_value, created) values(" & tag_0204PV27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0204PV27 = Nothing
Dim tag_0205LAMP01
tag_0205LAMP01 = HMIRuntime.Tags("0205LAMP01").Read
strSQL = "INSERT INTO z_tag_0205LAMP01 (tag_value, created) values(" & tag_0205LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205LAMP01 = Nothing
Dim tag_0205LSH01
tag_0205LSH01 = HMIRuntime.Tags("0205LSH01").Read
strSQL = "INSERT INTO z_tag_0205LSH01 (tag_value, created) values(" & tag_0205LSH01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205LSH01 = Nothing
Dim tag_0205LSL01
tag_0205LSL01 = HMIRuntime.Tags("0205LSL01").Read
strSQL = "INSERT INTO z_tag_0205LSL01 (tag_value, created) values(" & tag_0205LSL01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205LSL01 = Nothing
Dim tag_0205M01
tag_0205M01 = HMIRuntime.Tags("0205M01").Read
strSQL = "INSERT INTO z_tag_0205M01 (tag_value, created) values(" & tag_0205M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205M01 = Nothing
Dim tag_0205M02
tag_0205M02 = HMIRuntime.Tags("0205M02").Read
strSQL = "INSERT INTO z_tag_0205M02 (tag_value, created) values(" & tag_0205M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205M02 = Nothing
Dim tag_0205PV01
tag_0205PV01 = HMIRuntime.Tags("0205PV01").Read
strSQL = "INSERT INTO z_tag_0205PV01 (tag_value, created) values(" & tag_0205PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV01 = Nothing
Dim tag_0205PV02
tag_0205PV02 = HMIRuntime.Tags("0205PV02").Read
strSQL = "INSERT INTO z_tag_0205PV02 (tag_value, created) values(" & tag_0205PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV02 = Nothing
Dim tag_0205PV03
tag_0205PV03 = HMIRuntime.Tags("0205PV03").Read
strSQL = "INSERT INTO z_tag_0205PV03 (tag_value, created) values(" & tag_0205PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV03 = Nothing
Dim tag_0205PV04
tag_0205PV04 = HMIRuntime.Tags("0205PV04").Read
strSQL = "INSERT INTO z_tag_0205PV04 (tag_value, created) values(" & tag_0205PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV04 = Nothing
Dim tag_0205PV05
tag_0205PV05 = HMIRuntime.Tags("0205PV05").Read
strSQL = "INSERT INTO z_tag_0205PV05 (tag_value, created) values(" & tag_0205PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV05 = Nothing
Dim tag_0205PV06
tag_0205PV06 = HMIRuntime.Tags("0205PV06").Read
strSQL = "INSERT INTO z_tag_0205PV06 (tag_value, created) values(" & tag_0205PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV06 = Nothing
Dim tag_0205PV07
tag_0205PV07 = HMIRuntime.Tags("0205PV07").Read
strSQL = "INSERT INTO z_tag_0205PV07 (tag_value, created) values(" & tag_0205PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV07 = Nothing
Dim tag_0205PV08
tag_0205PV08 = HMIRuntime.Tags("0205PV08").Read
strSQL = "INSERT INTO z_tag_0205PV08 (tag_value, created) values(" & tag_0205PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV08 = Nothing
Dim tag_0205PV09
tag_0205PV09 = HMIRuntime.Tags("0205PV09").Read
strSQL = "INSERT INTO z_tag_0205PV09 (tag_value, created) values(" & tag_0205PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV09 = Nothing
Dim tag_0205PV10
tag_0205PV10 = HMIRuntime.Tags("0205PV10").Read
strSQL = "INSERT INTO z_tag_0205PV10 (tag_value, created) values(" & tag_0205PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV10 = Nothing
Dim tag_0205PV11
tag_0205PV11 = HMIRuntime.Tags("0205PV11").Read
strSQL = "INSERT INTO z_tag_0205PV11 (tag_value, created) values(" & tag_0205PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV11 = Nothing
Dim tag_0205PV12
tag_0205PV12 = HMIRuntime.Tags("0205PV12").Read
strSQL = "INSERT INTO z_tag_0205PV12 (tag_value, created) values(" & tag_0205PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV12 = Nothing
Dim tag_0205PV13
tag_0205PV13 = HMIRuntime.Tags("0205PV13").Read
strSQL = "INSERT INTO z_tag_0205PV13 (tag_value, created) values(" & tag_0205PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV13 = Nothing
Dim tag_0205PV14
tag_0205PV14 = HMIRuntime.Tags("0205PV14").Read
strSQL = "INSERT INTO z_tag_0205PV14 (tag_value, created) values(" & tag_0205PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV14 = Nothing
Dim tag_0205PV15
tag_0205PV15 = HMIRuntime.Tags("0205PV15").Read
strSQL = "INSERT INTO z_tag_0205PV15 (tag_value, created) values(" & tag_0205PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV15 = Nothing
Dim tag_0205PV16
tag_0205PV16 = HMIRuntime.Tags("0205PV16").Read
strSQL = "INSERT INTO z_tag_0205PV16 (tag_value, created) values(" & tag_0205PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV16 = Nothing
Dim tag_0205PV17
tag_0205PV17 = HMIRuntime.Tags("0205PV17").Read
strSQL = "INSERT INTO z_tag_0205PV17 (tag_value, created) values(" & tag_0205PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV17 = Nothing
Dim tag_0205PV18
tag_0205PV18 = HMIRuntime.Tags("0205PV18").Read
strSQL = "INSERT INTO z_tag_0205PV18 (tag_value, created) values(" & tag_0205PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205PV18 = Nothing
Dim tag_0205TET01_new
tag_0205TET01_new = HMIRuntime.Tags("0205TET01_new").Read
strSQL = "INSERT INTO z_tag_0205TET01_new (tag_value, created) values(" & tag_0205TET01_new & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0205TET01_new = Nothing
Dim tag_0206FQET01
tag_0206FQET01 = HMIRuntime.Tags("0206FQET01").Read
strSQL = "INSERT INTO z_tag_0206FQET01 (tag_value, created) values(" & tag_0206FQET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206FQET01 = Nothing
Dim tag_0206LAMP01
tag_0206LAMP01 = HMIRuntime.Tags("0206LAMP01").Read
strSQL = "INSERT INTO z_tag_0206LAMP01 (tag_value, created) values(" & tag_0206LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206LAMP01 = Nothing
Dim tag_0206PV03
tag_0206PV03 = HMIRuntime.Tags("0206PV03").Read
strSQL = "INSERT INTO z_tag_0206PV03 (tag_value, created) values(" & tag_0206PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV03 = Nothing
Dim tag_0206PV04
tag_0206PV04 = HMIRuntime.Tags("0206PV04").Read
strSQL = "INSERT INTO z_tag_0206PV04 (tag_value, created) values(" & tag_0206PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV04 = Nothing
Dim tag_0206PV05
tag_0206PV05 = HMIRuntime.Tags("0206PV05").Read
strSQL = "INSERT INTO z_tag_0206PV05 (tag_value, created) values(" & tag_0206PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV05 = Nothing
Dim tag_0206PV06
tag_0206PV06 = HMIRuntime.Tags("0206PV06").Read
strSQL = "INSERT INTO z_tag_0206PV06 (tag_value, created) values(" & tag_0206PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV06 = Nothing
Dim tag_0206PV07
tag_0206PV07 = HMIRuntime.Tags("0206PV07").Read
strSQL = "INSERT INTO z_tag_0206PV07 (tag_value, created) values(" & tag_0206PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV07 = Nothing
Dim tag_0206PV08
tag_0206PV08 = HMIRuntime.Tags("0206PV08").Read
strSQL = "INSERT INTO z_tag_0206PV08 (tag_value, created) values(" & tag_0206PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV08 = Nothing
Dim tag_0206PV10
tag_0206PV10 = HMIRuntime.Tags("0206PV10").Read
strSQL = "INSERT INTO z_tag_0206PV10 (tag_value, created) values(" & tag_0206PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV10 = Nothing
Dim tag_0206PV011
tag_0206PV011 = HMIRuntime.Tags("0206PV011").Read
strSQL = "INSERT INTO z_tag_0206PV011 (tag_value, created) values(" & tag_0206PV011 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV011 = Nothing
Dim tag_0206PV11
tag_0206PV11 = HMIRuntime.Tags("0206PV11").Read
strSQL = "INSERT INTO z_tag_0206PV11 (tag_value, created) values(" & tag_0206PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV11 = Nothing
Dim tag_0206PV012
tag_0206PV012 = HMIRuntime.Tags("0206PV012").Read
strSQL = "INSERT INTO z_tag_0206PV012 (tag_value, created) values(" & tag_0206PV012 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV012 = Nothing
Dim tag_0206PV013
tag_0206PV013 = HMIRuntime.Tags("0206PV013").Read
strSQL = "INSERT INTO z_tag_0206PV013 (tag_value, created) values(" & tag_0206PV013 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV013 = Nothing
Dim tag_0206PV021
tag_0206PV021 = HMIRuntime.Tags("0206PV021").Read
strSQL = "INSERT INTO z_tag_0206PV021 (tag_value, created) values(" & tag_0206PV021 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV021 = Nothing
Dim tag_0206PV022
tag_0206PV022 = HMIRuntime.Tags("0206PV022").Read
strSQL = "INSERT INTO z_tag_0206PV022 (tag_value, created) values(" & tag_0206PV022 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV022 = Nothing
Dim tag_0206PV023
tag_0206PV023 = HMIRuntime.Tags("0206PV023").Read
strSQL = "INSERT INTO z_tag_0206PV023 (tag_value, created) values(" & tag_0206PV023 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV023 = Nothing
Dim tag_0206PV091
tag_0206PV091 = HMIRuntime.Tags("0206PV091").Read
strSQL = "INSERT INTO z_tag_0206PV091 (tag_value, created) values(" & tag_0206PV091 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV091 = Nothing
Dim tag_0206PV092
tag_0206PV092 = HMIRuntime.Tags("0206PV092").Read
strSQL = "INSERT INTO z_tag_0206PV092 (tag_value, created) values(" & tag_0206PV092 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV092 = Nothing
Dim tag_0206PV093
tag_0206PV093 = HMIRuntime.Tags("0206PV093").Read
strSQL = "INSERT INTO z_tag_0206PV093 (tag_value, created) values(" & tag_0206PV093 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206PV093 = Nothing
Dim tag_0206TET01
tag_0206TET01 = HMIRuntime.Tags("0206TET01").Read
strSQL = "INSERT INTO z_tag_0206TET01 (tag_value, created) values(" & tag_0206TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0206TET01 = Nothing
Dim tag_0207M01
tag_0207M01 = HMIRuntime.Tags("0207M01").Read
strSQL = "INSERT INTO z_tag_0207M01 (tag_value, created) values(" & tag_0207M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0207M01 = Nothing
Dim tag_0208FS01
tag_0208FS01 = HMIRuntime.Tags("0208FS01").Read
strSQL = "INSERT INTO z_tag_0208FS01 (tag_value, created) values(" & tag_0208FS01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208FS01 = Nothing
Dim tag_0208FS02
tag_0208FS02 = HMIRuntime.Tags("0208FS02").Read
strSQL = "INSERT INTO z_tag_0208FS02 (tag_value, created) values(" & tag_0208FS02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208FS02 = Nothing
Dim tag_0208LET01
tag_0208LET01 = HMIRuntime.Tags("0208LET01").Read
strSQL = "INSERT INTO z_tag_0208LET01 (tag_value, created) values(" & tag_0208LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LET01 = Nothing
Dim tag_0208LET02
tag_0208LET02 = HMIRuntime.Tags("0208LET02").Read
strSQL = "INSERT INTO z_tag_0208LET02 (tag_value, created) values(" & tag_0208LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LET02 = Nothing
Dim tag_0208LET03
tag_0208LET03 = HMIRuntime.Tags("0208LET03").Read
strSQL = "INSERT INTO z_tag_0208LET03 (tag_value, created) values(" & tag_0208LET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LET03 = Nothing
Dim tag_0208LSH01
tag_0208LSH01 = HMIRuntime.Tags("0208LSH01").Read
strSQL = "INSERT INTO z_tag_0208LSH01 (tag_value, created) values(" & tag_0208LSH01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LSH01 = Nothing
Dim tag_0208LSH02
tag_0208LSH02 = HMIRuntime.Tags("0208LSH02").Read
strSQL = "INSERT INTO z_tag_0208LSH02 (tag_value, created) values(" & tag_0208LSH02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LSH02 = Nothing
Dim tag_0208LSH03
tag_0208LSH03 = HMIRuntime.Tags("0208LSH03").Read
strSQL = "INSERT INTO z_tag_0208LSH03 (tag_value, created) values(" & tag_0208LSH03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208LSH03 = Nothing
Dim tag_0208M01
tag_0208M01 = HMIRuntime.Tags("0208M01").Read
strSQL = "INSERT INTO z_tag_0208M01 (tag_value, created) values(" & tag_0208M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M01 = Nothing
Dim tag_0208M02
tag_0208M02 = HMIRuntime.Tags("0208M02").Read
strSQL = "INSERT INTO z_tag_0208M02 (tag_value, created) values(" & tag_0208M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M02 = Nothing
Dim tag_0208M03
tag_0208M03 = HMIRuntime.Tags("0208M03").Read
strSQL = "INSERT INTO z_tag_0208M03 (tag_value, created) values(" & tag_0208M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M03 = Nothing
Dim tag_0208M04
tag_0208M04 = HMIRuntime.Tags("0208M04").Read
strSQL = "INSERT INTO z_tag_0208M04 (tag_value, created) values(" & tag_0208M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M04 = Nothing
Dim tag_0208M05
tag_0208M05 = HMIRuntime.Tags("0208M05").Read
strSQL = "INSERT INTO z_tag_0208M05 (tag_value, created) values(" & tag_0208M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M05 = Nothing
Dim tag_0208M06
tag_0208M06 = HMIRuntime.Tags("0208M06").Read
strSQL = "INSERT INTO z_tag_0208M06 (tag_value, created) values(" & tag_0208M06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M06 = Nothing
Dim tag_0208M07
tag_0208M07 = HMIRuntime.Tags("0208M07").Read
strSQL = "INSERT INTO z_tag_0208M07 (tag_value, created) values(" & tag_0208M07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M07 = Nothing
Dim tag_0208M08
tag_0208M08 = HMIRuntime.Tags("0208M08").Read
strSQL = "INSERT INTO z_tag_0208M08 (tag_value, created) values(" & tag_0208M08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M08 = Nothing
Dim tag_0208M09
tag_0208M09 = HMIRuntime.Tags("0208M09").Read
strSQL = "INSERT INTO z_tag_0208M09 (tag_value, created) values(" & tag_0208M09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M09 = Nothing
Dim tag_0208M10
tag_0208M10 = HMIRuntime.Tags("0208M10").Read
strSQL = "INSERT INTO z_tag_0208M10 (tag_value, created) values(" & tag_0208M10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208M10 = Nothing
Dim tag_0208PET01
tag_0208PET01 = HMIRuntime.Tags("0208PET01").Read
strSQL = "INSERT INTO z_tag_0208PET01 (tag_value, created) values(" & tag_0208PET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PET01 = Nothing
Dim tag_0208PET02
tag_0208PET02 = HMIRuntime.Tags("0208PET02").Read
strSQL = "INSERT INTO z_tag_0208PET02 (tag_value, created) values(" & tag_0208PET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PET02 = Nothing
Dim tag_0208PET03
tag_0208PET03 = HMIRuntime.Tags("0208PET03").Read
strSQL = "INSERT INTO z_tag_0208PET03 (tag_value, created) values(" & tag_0208PET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PET03 = Nothing
Dim tag_0208PV01
tag_0208PV01 = HMIRuntime.Tags("0208PV01").Read
strSQL = "INSERT INTO z_tag_0208PV01 (tag_value, created) values(" & tag_0208PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV01 = Nothing
Dim tag_0208PV02
tag_0208PV02 = HMIRuntime.Tags("0208PV02").Read
strSQL = "INSERT INTO z_tag_0208PV02 (tag_value, created) values(" & tag_0208PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV02 = Nothing
Dim tag_0208PV03
tag_0208PV03 = HMIRuntime.Tags("0208PV03").Read
strSQL = "INSERT INTO z_tag_0208PV03 (tag_value, created) values(" & tag_0208PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV03 = Nothing
Dim tag_0208PV04
tag_0208PV04 = HMIRuntime.Tags("0208PV04").Read
strSQL = "INSERT INTO z_tag_0208PV04 (tag_value, created) values(" & tag_0208PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV04 = Nothing
Dim tag_0208PV05
tag_0208PV05 = HMIRuntime.Tags("0208PV05").Read
strSQL = "INSERT INTO z_tag_0208PV05 (tag_value, created) values(" & tag_0208PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV05 = Nothing
Dim tag_0208PV06
tag_0208PV06 = HMIRuntime.Tags("0208PV06").Read
strSQL = "INSERT INTO z_tag_0208PV06 (tag_value, created) values(" & tag_0208PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV06 = Nothing
Dim tag_0208PV07
tag_0208PV07 = HMIRuntime.Tags("0208PV07").Read
strSQL = "INSERT INTO z_tag_0208PV07 (tag_value, created) values(" & tag_0208PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV07 = Nothing
Dim tag_0208PV08
tag_0208PV08 = HMIRuntime.Tags("0208PV08").Read
strSQL = "INSERT INTO z_tag_0208PV08 (tag_value, created) values(" & tag_0208PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV08 = Nothing
Dim tag_0208PV08_status
tag_0208PV08_status = HMIRuntime.Tags("0208PV08_status").Read
strSQL = "INSERT INTO z_tag_0208PV08_status (tag_value, created) values(" & tag_0208PV08_status & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV08_status = Nothing
Dim tag_0208PV09
tag_0208PV09 = HMIRuntime.Tags("0208PV09").Read
strSQL = "INSERT INTO z_tag_0208PV09 (tag_value, created) values(" & tag_0208PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV09 = Nothing
Dim tag_0208PV10
tag_0208PV10 = HMIRuntime.Tags("0208PV10").Read
strSQL = "INSERT INTO z_tag_0208PV10 (tag_value, created) values(" & tag_0208PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV10 = Nothing
Dim tag_0208PV12
tag_0208PV12 = HMIRuntime.Tags("0208PV12").Read
strSQL = "INSERT INTO z_tag_0208PV12 (tag_value, created) values(" & tag_0208PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208PV12 = Nothing
Dim tag_0208TET01
tag_0208TET01 = HMIRuntime.Tags("0208TET01").Read
strSQL = "INSERT INTO z_tag_0208TET01 (tag_value, created) values(" & tag_0208TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TET01 = Nothing
Dim tag_0208TET02
tag_0208TET02 = HMIRuntime.Tags("0208TET02").Read
strSQL = "INSERT INTO z_tag_0208TET02 (tag_value, created) values(" & tag_0208TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TET02 = Nothing
Dim tag_0208TET03
tag_0208TET03 = HMIRuntime.Tags("0208TET03").Read
strSQL = "INSERT INTO z_tag_0208TET03 (tag_value, created) values(" & tag_0208TET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TET03 = Nothing
Dim tag_0208TET04
tag_0208TET04 = HMIRuntime.Tags("0208TET04").Read
strSQL = "INSERT INTO z_tag_0208TET04 (tag_value, created) values(" & tag_0208TET04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TET04 = Nothing
Dim tag_0208TET05
tag_0208TET05 = HMIRuntime.Tags("0208TET05").Read
strSQL = "INSERT INTO z_tag_0208TET05 (tag_value, created) values(" & tag_0208TET05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TET05 = Nothing
Dim tag_0208TSH01
tag_0208TSH01 = HMIRuntime.Tags("0208TSH01").Read
strSQL = "INSERT INTO z_tag_0208TSH01 (tag_value, created) values(" & tag_0208TSH01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0208TSH01 = Nothing
Dim tag_0209FQET01
tag_0209FQET01 = HMIRuntime.Tags("0209FQET01").Read
strSQL = "INSERT INTO z_tag_0209FQET01 (tag_value, created) values(" & tag_0209FQET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209FQET01 = Nothing
Dim tag_0209FQET02
tag_0209FQET02 = HMIRuntime.Tags("0209FQET02").Read
strSQL = "INSERT INTO z_tag_0209FQET02 (tag_value, created) values(" & tag_0209FQET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209FQET02 = Nothing
Dim tag_0209LET01
tag_0209LET01 = HMIRuntime.Tags("0209LET01").Read
strSQL = "INSERT INTO z_tag_0209LET01 (tag_value, created) values(" & tag_0209LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209LET01 = Nothing
Dim tag_0209LET02
tag_0209LET02 = HMIRuntime.Tags("0209LET02").Read
strSQL = "INSERT INTO z_tag_0209LET02 (tag_value, created) values(" & tag_0209LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209LET02 = Nothing
Dim tag_0209LET03
tag_0209LET03 = HMIRuntime.Tags("0209LET03").Read
strSQL = "INSERT INTO z_tag_0209LET03 (tag_value, created) values(" & tag_0209LET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209LET03 = Nothing
Dim tag_0209LET04
tag_0209LET04 = HMIRuntime.Tags("0209LET04").Read
strSQL = "INSERT INTO z_tag_0209LET04 (tag_value, created) values(" & tag_0209LET04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209LET04 = Nothing
Dim tag_0209M01
tag_0209M01 = HMIRuntime.Tags("0209M01").Read
strSQL = "INSERT INTO z_tag_0209M01 (tag_value, created) values(" & tag_0209M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209M01 = Nothing
Dim tag_0209M02
tag_0209M02 = HMIRuntime.Tags("0209M02").Read
strSQL = "INSERT INTO z_tag_0209M02 (tag_value, created) values(" & tag_0209M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209M02 = Nothing
Dim tag_0209M03
tag_0209M03 = HMIRuntime.Tags("0209M03").Read
strSQL = "INSERT INTO z_tag_0209M03 (tag_value, created) values(" & tag_0209M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209M03 = Nothing
Dim tag_0209PV01
tag_0209PV01 = HMIRuntime.Tags("0209PV01").Read
strSQL = "INSERT INTO z_tag_0209PV01 (tag_value, created) values(" & tag_0209PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV01 = Nothing
Dim tag_0209PV02
tag_0209PV02 = HMIRuntime.Tags("0209PV02").Read
strSQL = "INSERT INTO z_tag_0209PV02 (tag_value, created) values(" & tag_0209PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV02 = Nothing
Dim tag_0209PV03
tag_0209PV03 = HMIRuntime.Tags("0209PV03").Read
strSQL = "INSERT INTO z_tag_0209PV03 (tag_value, created) values(" & tag_0209PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV03 = Nothing
Dim tag_0209PV04
tag_0209PV04 = HMIRuntime.Tags("0209PV04").Read
strSQL = "INSERT INTO z_tag_0209PV04 (tag_value, created) values(" & tag_0209PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV04 = Nothing
Dim tag_0209PV05
tag_0209PV05 = HMIRuntime.Tags("0209PV05").Read
strSQL = "INSERT INTO z_tag_0209PV05 (tag_value, created) values(" & tag_0209PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV05 = Nothing
Dim tag_0209PV06
tag_0209PV06 = HMIRuntime.Tags("0209PV06").Read
strSQL = "INSERT INTO z_tag_0209PV06 (tag_value, created) values(" & tag_0209PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV06 = Nothing
Dim tag_0209PV07
tag_0209PV07 = HMIRuntime.Tags("0209PV07").Read
strSQL = "INSERT INTO z_tag_0209PV07 (tag_value, created) values(" & tag_0209PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV07 = Nothing
Dim tag_0209PV08
tag_0209PV08 = HMIRuntime.Tags("0209PV08").Read
strSQL = "INSERT INTO z_tag_0209PV08 (tag_value, created) values(" & tag_0209PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV08 = Nothing
Dim tag_0209PV09
tag_0209PV09 = HMIRuntime.Tags("0209PV09").Read
strSQL = "INSERT INTO z_tag_0209PV09 (tag_value, created) values(" & tag_0209PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV09 = Nothing
Dim tag_0209PV10
tag_0209PV10 = HMIRuntime.Tags("0209PV10").Read
strSQL = "INSERT INTO z_tag_0209PV10 (tag_value, created) values(" & tag_0209PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV10 = Nothing
Dim tag_0209PV11
tag_0209PV11 = HMIRuntime.Tags("0209PV11").Read
strSQL = "INSERT INTO z_tag_0209PV11 (tag_value, created) values(" & tag_0209PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV11 = Nothing
Dim tag_0209PV12
tag_0209PV12 = HMIRuntime.Tags("0209PV12").Read
strSQL = "INSERT INTO z_tag_0209PV12 (tag_value, created) values(" & tag_0209PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV12 = Nothing
Dim tag_0209PV13
tag_0209PV13 = HMIRuntime.Tags("0209PV13").Read
strSQL = "INSERT INTO z_tag_0209PV13 (tag_value, created) values(" & tag_0209PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV13 = Nothing
Dim tag_0209PV14
tag_0209PV14 = HMIRuntime.Tags("0209PV14").Read
strSQL = "INSERT INTO z_tag_0209PV14 (tag_value, created) values(" & tag_0209PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV14 = Nothing
Dim tag_0209PV15
tag_0209PV15 = HMIRuntime.Tags("0209PV15").Read
strSQL = "INSERT INTO z_tag_0209PV15 (tag_value, created) values(" & tag_0209PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV15 = Nothing
Dim tag_0209PV16
tag_0209PV16 = HMIRuntime.Tags("0209PV16").Read
strSQL = "INSERT INTO z_tag_0209PV16 (tag_value, created) values(" & tag_0209PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV16 = Nothing
Dim tag_0209PV17
tag_0209PV17 = HMIRuntime.Tags("0209PV17").Read
strSQL = "INSERT INTO z_tag_0209PV17 (tag_value, created) values(" & tag_0209PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV17 = Nothing
Dim tag_0209PV18
tag_0209PV18 = HMIRuntime.Tags("0209PV18").Read
strSQL = "INSERT INTO z_tag_0209PV18 (tag_value, created) values(" & tag_0209PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV18 = Nothing
Dim tag_0209PV19
tag_0209PV19 = HMIRuntime.Tags("0209PV19").Read
strSQL = "INSERT INTO z_tag_0209PV19 (tag_value, created) values(" & tag_0209PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV19 = Nothing
Dim tag_0209PV20
tag_0209PV20 = HMIRuntime.Tags("0209PV20").Read
strSQL = "INSERT INTO z_tag_0209PV20 (tag_value, created) values(" & tag_0209PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV20 = Nothing
Dim tag_0209PV21
tag_0209PV21 = HMIRuntime.Tags("0209PV21").Read
strSQL = "INSERT INTO z_tag_0209PV21 (tag_value, created) values(" & tag_0209PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV21 = Nothing
Dim tag_0209PV22
tag_0209PV22 = HMIRuntime.Tags("0209PV22").Read
strSQL = "INSERT INTO z_tag_0209PV22 (tag_value, created) values(" & tag_0209PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV22 = Nothing
Dim tag_0209PV23
tag_0209PV23 = HMIRuntime.Tags("0209PV23").Read
strSQL = "INSERT INTO z_tag_0209PV23 (tag_value, created) values(" & tag_0209PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV23 = Nothing
Dim tag_0209PV24
tag_0209PV24 = HMIRuntime.Tags("0209PV24").Read
strSQL = "INSERT INTO z_tag_0209PV24 (tag_value, created) values(" & tag_0209PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV24 = Nothing
Dim tag_0209PV25
tag_0209PV25 = HMIRuntime.Tags("0209PV25").Read
strSQL = "INSERT INTO z_tag_0209PV25 (tag_value, created) values(" & tag_0209PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV25 = Nothing
Dim tag_0209PV26
tag_0209PV26 = HMIRuntime.Tags("0209PV26").Read
strSQL = "INSERT INTO z_tag_0209PV26 (tag_value, created) values(" & tag_0209PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV26 = Nothing
Dim tag_0209PV27
tag_0209PV27 = HMIRuntime.Tags("0209PV27").Read
strSQL = "INSERT INTO z_tag_0209PV27 (tag_value, created) values(" & tag_0209PV27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV27 = Nothing
Dim tag_0209PV28
tag_0209PV28 = HMIRuntime.Tags("0209PV28").Read
strSQL = "INSERT INTO z_tag_0209PV28 (tag_value, created) values(" & tag_0209PV28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV28 = Nothing
Dim tag_0209PV29
tag_0209PV29 = HMIRuntime.Tags("0209PV29").Read
strSQL = "INSERT INTO z_tag_0209PV29 (tag_value, created) values(" & tag_0209PV29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV29 = Nothing
Dim tag_0209PV30
tag_0209PV30 = HMIRuntime.Tags("0209PV30").Read
strSQL = "INSERT INTO z_tag_0209PV30 (tag_value, created) values(" & tag_0209PV30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV30 = Nothing
Dim tag_0209PV31
tag_0209PV31 = HMIRuntime.Tags("0209PV31").Read
strSQL = "INSERT INTO z_tag_0209PV31 (tag_value, created) values(" & tag_0209PV31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV31 = Nothing
Dim tag_0209PV32
tag_0209PV32 = HMIRuntime.Tags("0209PV32").Read
strSQL = "INSERT INTO z_tag_0209PV32 (tag_value, created) values(" & tag_0209PV32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV32 = Nothing
Dim tag_0209PV33
tag_0209PV33 = HMIRuntime.Tags("0209PV33").Read
strSQL = "INSERT INTO z_tag_0209PV33 (tag_value, created) values(" & tag_0209PV33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV33 = Nothing
Dim tag_0209PV34
tag_0209PV34 = HMIRuntime.Tags("0209PV34").Read
strSQL = "INSERT INTO z_tag_0209PV34 (tag_value, created) values(" & tag_0209PV34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV34 = Nothing
Dim tag_0209PV35
tag_0209PV35 = HMIRuntime.Tags("0209PV35").Read
strSQL = "INSERT INTO z_tag_0209PV35 (tag_value, created) values(" & tag_0209PV35 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV35 = Nothing
Dim tag_0209PV36
tag_0209PV36 = HMIRuntime.Tags("0209PV36").Read
strSQL = "INSERT INTO z_tag_0209PV36 (tag_value, created) values(" & tag_0209PV36 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV36 = Nothing
Dim tag_0209PV37
tag_0209PV37 = HMIRuntime.Tags("0209PV37").Read
strSQL = "INSERT INTO z_tag_0209PV37 (tag_value, created) values(" & tag_0209PV37 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV37 = Nothing
Dim tag_0209PV40
tag_0209PV40 = HMIRuntime.Tags("0209PV40").Read
strSQL = "INSERT INTO z_tag_0209PV40 (tag_value, created) values(" & tag_0209PV40 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV40 = Nothing
Dim tag_0209PV41
tag_0209PV41 = HMIRuntime.Tags("0209PV41").Read
strSQL = "INSERT INTO z_tag_0209PV41 (tag_value, created) values(" & tag_0209PV41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV41 = Nothing
Dim tag_0209PV42
tag_0209PV42 = HMIRuntime.Tags("0209PV42").Read
strSQL = "INSERT INTO z_tag_0209PV42 (tag_value, created) values(" & tag_0209PV42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209PV42 = Nothing
Dim tag_0209QET01
tag_0209QET01 = HMIRuntime.Tags("0209QET01").Read
strSQL = "INSERT INTO z_tag_0209QET01 (tag_value, created) values(" & tag_0209QET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209QET01 = Nothing
Dim tag_0209TET01
tag_0209TET01 = HMIRuntime.Tags("0209TET01").Read
strSQL = "INSERT INTO z_tag_0209TET01 (tag_value, created) values(" & tag_0209TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209TET01 = Nothing
Dim tag_0209TET02
tag_0209TET02 = HMIRuntime.Tags("0209TET02").Read
strSQL = "INSERT INTO z_tag_0209TET02 (tag_value, created) values(" & tag_0209TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0209TET02 = Nothing
Dim tag_0210LET01
tag_0210LET01 = HMIRuntime.Tags("0210LET01").Read
strSQL = "INSERT INTO z_tag_0210LET01 (tag_value, created) values(" & tag_0210LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210LET01 = Nothing
Dim tag_0210LSH01
tag_0210LSH01 = HMIRuntime.Tags("0210LSH01").Read
strSQL = "INSERT INTO z_tag_0210LSH01 (tag_value, created) values(" & tag_0210LSH01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210LSH01 = Nothing
Dim tag_0210M01
tag_0210M01 = HMIRuntime.Tags("0210M01").Read
strSQL = "INSERT INTO z_tag_0210M01 (tag_value, created) values(" & tag_0210M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210M01 = Nothing
Dim tag_0210M02
tag_0210M02 = HMIRuntime.Tags("0210M02").Read
strSQL = "INSERT INTO z_tag_0210M02 (tag_value, created) values(" & tag_0210M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210M02 = Nothing
Dim tag_0210PV01
tag_0210PV01 = HMIRuntime.Tags("0210PV01").Read
strSQL = "INSERT INTO z_tag_0210PV01 (tag_value, created) values(" & tag_0210PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210PV01 = Nothing
Dim tag_0210PV02
tag_0210PV02 = HMIRuntime.Tags("0210PV02").Read
strSQL = "INSERT INTO z_tag_0210PV02 (tag_value, created) values(" & tag_0210PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210PV02 = Nothing
Dim tag_0210PV03
tag_0210PV03 = HMIRuntime.Tags("0210PV03").Read
strSQL = "INSERT INTO z_tag_0210PV03 (tag_value, created) values(" & tag_0210PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210PV03 = Nothing
Dim tag_0210TET01
tag_0210TET01 = HMIRuntime.Tags("0210TET01").Read
strSQL = "INSERT INTO z_tag_0210TET01 (tag_value, created) values(" & tag_0210TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210TET01 = Nothing
Dim tag_0210TET02
tag_0210TET02 = HMIRuntime.Tags("0210TET02").Read
strSQL = "INSERT INTO z_tag_0210TET02 (tag_value, created) values(" & tag_0210TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210TET02 = Nothing
Dim tag_0210TET03
tag_0210TET03 = HMIRuntime.Tags("0210TET03").Read
strSQL = "INSERT INTO z_tag_0210TET03 (tag_value, created) values(" & tag_0210TET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0210TET03 = Nothing
Dim tag_0401_1FQET01
tag_0401_1FQET01 = HMIRuntime.Tags("0401_1FQET01").Read
strSQL = "INSERT INTO z_tag_0401_1FQET01 (tag_value, created) values(" & tag_0401_1FQET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_1FQET01 = Nothing
Dim tag_0601TET05
tag_0601TET05 = HMIRuntime.Tags("0601TET05").Read
strSQL = "INSERT INTO z_tag_0601TET05 (tag_value, created) values(" & tag_0601TET05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601TET05 = Nothing
Dim tag_0601TET06
tag_0601TET06 = HMIRuntime.Tags("0601TET06").Read
strSQL = "INSERT INTO z_tag_0601TET06 (tag_value, created) values(" & tag_0601TET06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601TET06 = Nothing
Dim tag_02041FQET01
tag_02041FQET01 = HMIRuntime.Tags("02041FQET01").Read
strSQL = "INSERT INTO z_tag_02041FQET01 (tag_value, created) values(" & tag_02041FQET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02041FQET01 = Nothing
Dim tag_02041PET01
tag_02041PET01 = HMIRuntime.Tags("02041PET01").Read
strSQL = "INSERT INTO z_tag_02041PET01 (tag_value, created) values(" & tag_02041PET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02041PET01 = Nothing
Dim tag_02041TET01
tag_02041TET01 = HMIRuntime.Tags("02041TET01").Read
strSQL = "INSERT INTO z_tag_02041TET01 (tag_value, created) values(" & tag_02041TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02041TET01 = Nothing
Dim tag_02041TET02
tag_02041TET02 = HMIRuntime.Tags("02041TET02").Read
strSQL = "INSERT INTO z_tag_02041TET02 (tag_value, created) values(" & tag_02041TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02041TET02 = Nothing
Dim tag_02041TET03
tag_02041TET03 = HMIRuntime.Tags("02041TET03").Read
strSQL = "INSERT INTO z_tag_02041TET03 (tag_value, created) values(" & tag_02041TET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02041TET03 = Nothing
Dim tag_02042LAMP01
tag_02042LAMP01 = HMIRuntime.Tags("02042LAMP01").Read
strSQL = "INSERT INTO z_tag_02042LAMP01 (tag_value, created) values(" & tag_02042LAMP01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042LAMP01 = Nothing
Dim tag_02042PV01
tag_02042PV01 = HMIRuntime.Tags("02042PV01").Read
strSQL = "INSERT INTO z_tag_02042PV01 (tag_value, created) values(" & tag_02042PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV01 = Nothing
Dim tag_02042PV02
tag_02042PV02 = HMIRuntime.Tags("02042PV02").Read
strSQL = "INSERT INTO z_tag_02042PV02 (tag_value, created) values(" & tag_02042PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV02 = Nothing
Dim tag_02042PV03
tag_02042PV03 = HMIRuntime.Tags("02042PV03").Read
strSQL = "INSERT INTO z_tag_02042PV03 (tag_value, created) values(" & tag_02042PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV03 = Nothing
Dim tag_02042PV04
tag_02042PV04 = HMIRuntime.Tags("02042PV04").Read
strSQL = "INSERT INTO z_tag_02042PV04 (tag_value, created) values(" & tag_02042PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV04 = Nothing
Dim tag_02042PV05
tag_02042PV05 = HMIRuntime.Tags("02042PV05").Read
strSQL = "INSERT INTO z_tag_02042PV05 (tag_value, created) values(" & tag_02042PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV05 = Nothing
Dim tag_02042PV06
tag_02042PV06 = HMIRuntime.Tags("02042PV06").Read
strSQL = "INSERT INTO z_tag_02042PV06 (tag_value, created) values(" & tag_02042PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042PV06 = Nothing
Dim tag_02042TET01
tag_02042TET01 = HMIRuntime.Tags("02042TET01").Read
strSQL = "INSERT INTO z_tag_02042TET01 (tag_value, created) values(" & tag_02042TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042TET01 = Nothing
Dim tag_02042TET02
tag_02042TET02 = HMIRuntime.Tags("02042TET02").Read
strSQL = "INSERT INTO z_tag_02042TET02 (tag_value, created) values(" & tag_02042TET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_02042TET02 = Nothing
Dim tag_Bao_cao_canh_khuay
tag_Bao_cao_canh_khuay = HMIRuntime.Tags("Bao_cao_canh_khuay").Read
strSQL = "INSERT INTO z_tag_Bao_cao_canh_khuay (tag_value, created) values(" & tag_Bao_cao_canh_khuay & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Bao_cao_canh_khuay = Nothing
Dim tag_Bao_thap_canh_khuay
tag_Bao_thap_canh_khuay = HMIRuntime.Tags("Bao_thap_canh_khuay").Read
strSQL = "INSERT INTO z_tag_Bao_thap_canh_khuay (tag_value, created) values(" & tag_Bao_thap_canh_khuay & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Bao_thap_canh_khuay = Nothing
Dim tag_BaoaskhiLN
tag_BaoaskhiLN = HMIRuntime.Tags("BaoaskhiLN").Read
strSQL = "INSERT INTO z_tag_BaoaskhiLN (tag_value, created) values(" & tag_BaoaskhiLN & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BaoaskhiLN = Nothing
Dim tag_baocaphoiNLX
tag_baocaphoiNLX = HMIRuntime.Tags("baocaphoiNLX").Read
strSQL = "INSERT INTO z_tag_baocaphoiNLX (tag_value, created) values(" & tag_baocaphoiNLX & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baocaphoiNLX = Nothing
Dim tag_BBT05PV
tag_BBT05PV = HMIRuntime.Tags("BBT05PV").Read
strSQL = "INSERT INTO z_tag_BBT05PV (tag_value, created) values(" & tag_BBT05PV & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BBT05PV = Nothing
Dim tag_BBT06PV
tag_BBT06PV = HMIRuntime.Tags("BBT06PV").Read
strSQL = "INSERT INTO z_tag_BBT06PV (tag_value, created) values(" & tag_BBT06PV & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_BBT06PV = Nothing
Dim tag_bomcipnau_st
tag_bomcipnau_st = HMIRuntime.Tags("bomcipnau_st").Read
strSQL = "INSERT INTO z_tag_bomcipnau_st (tag_value, created) values(" & tag_bomcipnau_st & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bomcipnau_st = Nothing
Dim tag_CIPBrewhouse_run
tag_CIPBrewhouse_run = HMIRuntime.Tags("CIPBrewhouse_run").Read
strSQL = "INSERT INTO z_tag_CIPBrewhouse_run (tag_value, created) values(" & tag_CIPBrewhouse_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIPBrewhouse_run = Nothing
Dim tag_CIPBrewhouse_seq
tag_CIPBrewhouse_seq = HMIRuntime.Tags("CIPBrewhouse_seq").Read
strSQL = "INSERT INTO z_tag_CIPBrewhouse_seq (tag_value, created) values(" & tag_CIPBrewhouse_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIPBrewhouse_seq = Nothing
Dim tag_CV01_out
tag_CV01_out = HMIRuntime.Tags("CV01_out").Read
strSQL = "INSERT INTO z_tag_CV01_out (tag_value, created) values(" & tag_CV01_out & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CV01_out = Nothing
Dim tag_Densisy_WK
tag_Densisy_WK = HMIRuntime.Tags("Densisy_WK").Read
strSQL = "INSERT INTO z_tag_Densisy_WK (tag_value, created) values(" & tag_Densisy_WK & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Densisy_WK = Nothing
Dim tag_dien_ap
tag_dien_ap = HMIRuntime.Tags("dien_ap").Read
strSQL = "INSERT INTO z_tag_dien_ap (tag_value, created) values(" & tag_dien_ap & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_dien_ap = Nothing
Dim tag_Elphapy_flow_pv
tag_Elphapy_flow_pv = HMIRuntime.Tags("Elphapy_flow_pv").Read
strSQL = "INSERT INTO z_tag_Elphapy_flow_pv (tag_value, created) values(" & tag_Elphapy_flow_pv & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Elphapy_flow_pv = Nothing
Dim tag_elthapy_flow
tag_elthapy_flow = HMIRuntime.Tags("elthapy_flow").Read
strSQL = "INSERT INTO z_tag_elthapy_flow (tag_value, created) values(" & tag_elthapy_flow & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_elthapy_flow = Nothing
Dim tag_Highlimit
tag_Highlimit = HMIRuntime.Tags("Highlimit").Read
strSQL = "INSERT INTO z_tag_Highlimit (tag_value, created) values(" & tag_Highlimit & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Highlimit = Nothing
Dim tag_Hl_test
tag_Hl_test = HMIRuntime.Tags("Hl_test").Read
strSQL = "INSERT INTO z_tag_Hl_test (tag_value, created) values(" & tag_Hl_test & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Hl_test = Nothing
Dim tag_HoldingVessel_Auto
tag_HoldingVessel_Auto = HMIRuntime.Tags("HoldingVessel_Auto").Read
strSQL = "INSERT INTO z_tag_HoldingVessel_Auto (tag_value, created) values(" & tag_HoldingVessel_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_HoldingVessel_Auto = Nothing
Dim tag_HoldingVessel_CIP
tag_HoldingVessel_CIP = HMIRuntime.Tags("HoldingVessel_CIP").Read
strSQL = "INSERT INTO z_tag_HoldingVessel_CIP (tag_value, created) values(" & tag_HoldingVessel_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_HoldingVessel_CIP = Nothing
Dim tag_HoldingVessel_run
tag_HoldingVessel_run = HMIRuntime.Tags("HoldingVessel_run").Read
strSQL = "INSERT INTO z_tag_HoldingVessel_run (tag_value, created) values(" & tag_HoldingVessel_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_HoldingVessel_run = Nothing
Dim tag_HoldingVessel_seq
tag_HoldingVessel_seq = HMIRuntime.Tags("HoldingVessel_seq").Read
strSQL = "INSERT INTO z_tag_HoldingVessel_seq (tag_value, created) values(" & tag_HoldingVessel_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_HoldingVessel_seq = Nothing
Dim tag_HoldingVessel_seq_add
tag_HoldingVessel_seq_add = HMIRuntime.Tags("HoldingVessel_seq_add").Read
strSQL = "INSERT INTO z_tag_HoldingVessel_seq_add (tag_value, created) values(" & tag_HoldingVessel_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_HoldingVessel_seq_add = Nothing
Dim tag_Hop01_run
tag_Hop01_run = HMIRuntime.Tags("Hop01_run").Read
strSQL = "INSERT INTO z_tag_Hop01_run (tag_value, created) values(" & tag_Hop01_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Hop01_run = Nothing
Dim tag_Hop01_seq
tag_Hop01_seq = HMIRuntime.Tags("Hop01_seq").Read
strSQL = "INSERT INTO z_tag_Hop01_seq (tag_value, created) values(" & tag_Hop01_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Hop01_seq = Nothing
Dim tag_Hop02_run
tag_Hop02_run = HMIRuntime.Tags("Hop02_run").Read
strSQL = "INSERT INTO z_tag_Hop02_run (tag_value, created) values(" & tag_Hop02_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Hop02_run = Nothing
Dim tag_Hop02_seq
tag_Hop02_seq = HMIRuntime.Tags("Hop02_seq").Read
strSQL = "INSERT INTO z_tag_Hop02_seq (tag_value, created) values(" & tag_Hop02_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Hop02_seq = Nothing
Dim tag_LauterTun_Auto
tag_LauterTun_Auto = HMIRuntime.Tags("LauterTun_Auto").Read
strSQL = "INSERT INTO z_tag_LauterTun_Auto (tag_value, created) values(" & tag_LauterTun_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LauterTun_Auto = Nothing
Dim tag_LauterTun_CIP
tag_LauterTun_CIP = HMIRuntime.Tags("LauterTun_CIP").Read
strSQL = "INSERT INTO z_tag_LauterTun_CIP (tag_value, created) values(" & tag_LauterTun_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LauterTun_CIP = Nothing
Dim tag_LauterTun_run
tag_LauterTun_run = HMIRuntime.Tags("LauterTun_run").Read
strSQL = "INSERT INTO z_tag_LauterTun_run (tag_value, created) values(" & tag_LauterTun_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LauterTun_run = Nothing
Dim tag_LauterTun_seq
tag_LauterTun_seq = HMIRuntime.Tags("LauterTun_seq").Read
strSQL = "INSERT INTO z_tag_LauterTun_seq (tag_value, created) values(" & tag_LauterTun_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LauterTun_seq = Nothing
Dim tag_LauterTun_seq_add
tag_LauterTun_seq_add = HMIRuntime.Tags("LauterTun_seq_add").Read
strSQL = "INSERT INTO z_tag_LauterTun_seq_add (tag_value, created) values(" & tag_LauterTun_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LauterTun_seq_add = Nothing
Dim tag_Lowlimit
tag_Lowlimit = HMIRuntime.Tags("Lowlimit").Read
strSQL = "INSERT INTO z_tag_Lowlimit (tag_value, created) values(" & tag_Lowlimit & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Lowlimit = Nothing
Dim tag_LTSpargingwater
tag_LTSpargingwater = HMIRuntime.Tags("LTSpargingwater").Read
strSQL = "INSERT INTO z_tag_LTSpargingwater (tag_value, created) values(" & tag_LTSpargingwater & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LTSpargingwater = Nothing
Dim tag_LTStepWort
tag_LTStepWort = HMIRuntime.Tags("LTStepWort").Read
strSQL = "INSERT INTO z_tag_LTStepWort (tag_value, created) values(" & tag_LTStepWort & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_LTStepWort = Nothing
Dim tag_MaltBin_Auto
tag_MaltBin_Auto = HMIRuntime.Tags("MaltBin_Auto").Read
strSQL = "INSERT INTO z_tag_MaltBin_Auto (tag_value, created) values(" & tag_MaltBin_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltBin_Auto = Nothing
Dim tag_MaltBin_run
tag_MaltBin_run = HMIRuntime.Tags("MaltBin_run").Read
strSQL = "INSERT INTO z_tag_MaltBin_run (tag_value, created) values(" & tag_MaltBin_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltBin_run = Nothing
Dim tag_MaltBin_seq
tag_MaltBin_seq = HMIRuntime.Tags("MaltBin_seq").Read
strSQL = "INSERT INTO z_tag_MaltBin_seq (tag_value, created) values(" & tag_MaltBin_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltBin_seq = Nothing
Dim tag_MaltBin_seq_add
tag_MaltBin_seq_add = HMIRuntime.Tags("MaltBin_seq_add").Read
strSQL = "INSERT INTO z_tag_MaltBin_seq_add (tag_value, created) values(" & tag_MaltBin_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltBin_seq_add = Nothing
Dim tag_MaltIntake_run
tag_MaltIntake_run = HMIRuntime.Tags("MaltIntake_run").Read
strSQL = "INSERT INTO z_tag_MaltIntake_run (tag_value, created) values(" & tag_MaltIntake_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltIntake_run = Nothing
Dim tag_MaltIntake_seq
tag_MaltIntake_seq = HMIRuntime.Tags("MaltIntake_seq").Read
strSQL = "INSERT INTO z_tag_MaltIntake_seq (tag_value, created) values(" & tag_MaltIntake_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltIntake_seq = Nothing
Dim tag_MaltMilling_run
tag_MaltMilling_run = HMIRuntime.Tags("MaltMilling_run").Read
strSQL = "INSERT INTO z_tag_MaltMilling_run (tag_value, created) values(" & tag_MaltMilling_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltMilling_run = Nothing
Dim tag_MaltMilling_seq
tag_MaltMilling_seq = HMIRuntime.Tags("MaltMilling_seq").Read
strSQL = "INSERT INTO z_tag_MaltMilling_seq (tag_value, created) values(" & tag_MaltMilling_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MaltMilling_seq = Nothing
Dim tag_MashTun_Auto
tag_MashTun_Auto = HMIRuntime.Tags("MashTun_Auto").Read
strSQL = "INSERT INTO z_tag_MashTun_Auto (tag_value, created) values(" & tag_MashTun_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MashTun_Auto = Nothing
Dim tag_MashTun_CIP
tag_MashTun_CIP = HMIRuntime.Tags("MashTun_CIP").Read
strSQL = "INSERT INTO z_tag_MashTun_CIP (tag_value, created) values(" & tag_MashTun_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MashTun_CIP = Nothing
Dim tag_MashTun_run
tag_MashTun_run = HMIRuntime.Tags("MashTun_run").Read
strSQL = "INSERT INTO z_tag_MashTun_run (tag_value, created) values(" & tag_MashTun_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MashTun_run = Nothing
Dim tag_MashTun_seq
tag_MashTun_seq = HMIRuntime.Tags("MashTun_seq").Read
strSQL = "INSERT INTO z_tag_MashTun_seq (tag_value, created) values(" & tag_MashTun_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MashTun_seq = Nothing
Dim tag_MashTun_seq_add
tag_MashTun_seq_add = HMIRuntime.Tags("MashTun_seq_add").Read
strSQL = "INSERT INTO z_tag_MashTun_seq_add (tag_value, created) values(" & tag_MashTun_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MashTun_seq_add = Nothing
Dim tag_Mostart
tag_Mostart = HMIRuntime.Tags("Mostart").Read
strSQL = "INSERT INTO z_tag_Mostart (tag_value, created) values(" & tag_Mostart & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mostart = Nothing
Dim tag_Mucthap_TBF05
tag_Mucthap_TBF05 = HMIRuntime.Tags("Mucthap_TBF05").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF05 (tag_value, created) values(" & tag_Mucthap_TBF05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF05 = Nothing
Dim tag_Mucthap_TBF06
tag_Mucthap_TBF06 = HMIRuntime.Tags("Mucthap_TBF06").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF06 (tag_value, created) values(" & tag_Mucthap_TBF06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF06 = Nothing
Dim tag_nghiengao_st
tag_nghiengao_st = HMIRuntime.Tags("nghiengao_st").Read
strSQL = "INSERT INTO z_tag_nghiengao_st (tag_value, created) values(" & tag_nghiengao_st & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nghiengao_st = Nothing
Dim tag_nghienmalt_st
tag_nghienmalt_st = HMIRuntime.Tags("nghienmalt_st").Read
strSQL = "INSERT INTO z_tag_nghienmalt_st (tag_value, created) values(" & tag_nghienmalt_st & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nghienmalt_st = Nothing
Dim tag_PID_0201M01
tag_PID_0201M01 = HMIRuntime.Tags("PID_0201M01").Read
strSQL = "INSERT INTO z_tag_PID_0201M01 (tag_value, created) values(" & tag_PID_0201M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0201M01 = Nothing
Dim tag_PID_0201M02
tag_PID_0201M02 = HMIRuntime.Tags("PID_0201M02").Read
strSQL = "INSERT INTO z_tag_PID_0201M02 (tag_value, created) values(" & tag_PID_0201M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0201M02 = Nothing
Dim tag_PID_0201M03
tag_PID_0201M03 = HMIRuntime.Tags("PID_0201M03").Read
strSQL = "INSERT INTO z_tag_PID_0201M03 (tag_value, created) values(" & tag_PID_0201M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0201M03 = Nothing
Dim tag_PID_0201M04
tag_PID_0201M04 = HMIRuntime.Tags("PID_0201M04").Read
strSQL = "INSERT INTO z_tag_PID_0201M04 (tag_value, created) values(" & tag_PID_0201M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0201M04 = Nothing
Dim tag_PID_0204M01
tag_PID_0204M01 = HMIRuntime.Tags("PID_0204M01").Read
strSQL = "INSERT INTO z_tag_PID_0204M01 (tag_value, created) values(" & tag_PID_0204M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0204M01 = Nothing
Dim tag_PID_CV01
tag_PID_CV01 = HMIRuntime.Tags("PID_CV01").Read
strSQL = "INSERT INTO z_tag_PID_CV01 (tag_value, created) values(" & tag_PID_CV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_CV01 = Nothing
Dim tag_PID0202CV04
tag_PID0202CV04 = HMIRuntime.Tags("PID0202CV04").Read
strSQL = "INSERT INTO z_tag_PID0202CV04 (tag_value, created) values(" & tag_PID0202CV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID0202CV04 = Nothing
Dim tag_PID0202M05
tag_PID0202M05 = HMIRuntime.Tags("PID0202M05").Read
strSQL = "INSERT INTO z_tag_PID0202M05 (tag_value, created) values(" & tag_PID0202M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID0202M05 = Nothing
Dim tag_PID0208CV02
tag_PID0208CV02 = HMIRuntime.Tags("PID0208CV02").Read
strSQL = "INSERT INTO z_tag_PID0208CV02 (tag_value, created) values(" & tag_PID0208CV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID0208CV02 = Nothing
Dim tag_PID0208CV03
tag_PID0208CV03 = HMIRuntime.Tags("PID0208CV03").Read
strSQL = "INSERT INTO z_tag_PID0208CV03 (tag_value, created) values(" & tag_PID0208CV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID0208CV03 = Nothing
Dim tag_PID0208M10
tag_PID0208M10 = HMIRuntime.Tags("PID0208M10").Read
strSQL = "INSERT INTO z_tag_PID0208M10 (tag_value, created) values(" & tag_PID0208M10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID0208M10 = Nothing
Dim tag_Quatkholanh 1
tag_Quatkholanh 1 = HMIRuntime.Tags("Quatkholanh 1").Read
strSQL = "INSERT INTO z_tag_Quatkholanh 1 (tag_value, created) values(" & tag_Quatkholanh 1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Quatkholanh 1 = Nothing
Dim tag_Quatkholanh 2
tag_Quatkholanh 2 = HMIRuntime.Tags("Quatkholanh 2").Read
strSQL = "INSERT INTO z_tag_Quatkholanh 2 (tag_value, created) values(" & tag_Quatkholanh 2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Quatkholanh 2 = Nothing
Dim tag_RiceCooker_Auto
tag_RiceCooker_Auto = HMIRuntime.Tags("RiceCooker_Auto").Read
strSQL = "INSERT INTO z_tag_RiceCooker_Auto (tag_value, created) values(" & tag_RiceCooker_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceCooker_Auto = Nothing
Dim tag_RiceCooker_CIP
tag_RiceCooker_CIP = HMIRuntime.Tags("RiceCooker_CIP").Read
strSQL = "INSERT INTO z_tag_RiceCooker_CIP (tag_value, created) values(" & tag_RiceCooker_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceCooker_CIP = Nothing
Dim tag_RiceCooker_run
tag_RiceCooker_run = HMIRuntime.Tags("RiceCooker_run").Read
strSQL = "INSERT INTO z_tag_RiceCooker_run (tag_value, created) values(" & tag_RiceCooker_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceCooker_run = Nothing
Dim tag_RiceCooker_seq
tag_RiceCooker_seq = HMIRuntime.Tags("RiceCooker_seq").Read
strSQL = "INSERT INTO z_tag_RiceCooker_seq (tag_value, created) values(" & tag_RiceCooker_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceCooker_seq = Nothing
Dim tag_RiceCooker_seq_add
tag_RiceCooker_seq_add = HMIRuntime.Tags("RiceCooker_seq_add").Read
strSQL = "INSERT INTO z_tag_RiceCooker_seq_add (tag_value, created) values(" & tag_RiceCooker_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceCooker_seq_add = Nothing
Dim tag_RiceIntake_run
tag_RiceIntake_run = HMIRuntime.Tags("RiceIntake_run").Read
strSQL = "INSERT INTO z_tag_RiceIntake_run (tag_value, created) values(" & tag_RiceIntake_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceIntake_run = Nothing
Dim tag_RiceIntake_seq
tag_RiceIntake_seq = HMIRuntime.Tags("RiceIntake_seq").Read
strSQL = "INSERT INTO z_tag_RiceIntake_seq (tag_value, created) values(" & tag_RiceIntake_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceIntake_seq = Nothing
Dim tag_RiceMilling_run
tag_RiceMilling_run = HMIRuntime.Tags("RiceMilling_run").Read
strSQL = "INSERT INTO z_tag_RiceMilling_run (tag_value, created) values(" & tag_RiceMilling_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceMilling_run = Nothing
Dim tag_RiceMilling_seq
tag_RiceMilling_seq = HMIRuntime.Tags("RiceMilling_seq").Read
strSQL = "INSERT INTO z_tag_RiceMilling_seq (tag_value, created) values(" & tag_RiceMilling_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_RiceMilling_seq = Nothing
Dim tag_Saykholanh
tag_Saykholanh = HMIRuntime.Tags("Saykholanh").Read
strSQL = "INSERT INTO z_tag_Saykholanh (tag_value, created) values(" & tag_Saykholanh & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Saykholanh = Nothing
Dim tag_select_pid
tag_select_pid = HMIRuntime.Tags("select_pid").Read
strSQL = "INSERT INTO z_tag_select_pid (tag_value, created) values(" & tag_select_pid & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_select_pid = Nothing
Dim tag_SpentGrain_run
tag_SpentGrain_run = HMIRuntime.Tags("SpentGrain_run").Read
strSQL = "INSERT INTO z_tag_SpentGrain_run (tag_value, created) values(" & tag_SpentGrain_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_SpentGrain_run = Nothing
Dim tag_SpentGrain_seq
tag_SpentGrain_seq = HMIRuntime.Tags("SpentGrain_seq").Read
strSQL = "INSERT INTO z_tag_SpentGrain_seq (tag_value, created) values(" & tag_SpentGrain_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_SpentGrain_seq = Nothing
Dim tag_Status_th_caphoinlx
tag_Status_th_caphoinlx = HMIRuntime.Tags("Status_th_caphoinlx").Read
strSQL = "INSERT INTO z_tag_Status_th_caphoinlx (tag_value, created) values(" & tag_Status_th_caphoinlx & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_th_caphoinlx = Nothing
Dim tag_Status_vanxaba
tag_Status_vanxaba = HMIRuntime.Tags("Status_vanxaba").Read
strSQL = "INSERT INTO z_tag_Status_vanxaba (tag_value, created) values(" & tag_Status_vanxaba & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Status_vanxaba = Nothing
Dim tag_steam_flow
tag_steam_flow = HMIRuntime.Tags("steam_flow").Read
strSQL = "INSERT INTO z_tag_steam_flow (tag_value, created) values(" & tag_steam_flow & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_steam_flow = Nothing
Dim tag_steam_pressure
tag_steam_pressure = HMIRuntime.Tags("steam_pressure").Read
strSQL = "INSERT INTO z_tag_steam_pressure (tag_value, created) values(" & tag_steam_pressure & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_steam_pressure = Nothing
Dim tag_Steam_wk_ist
tag_Steam_wk_ist = HMIRuntime.Tags("Steam_wk_ist").Read
strSQL = "INSERT INTO z_tag_Steam_wk_ist (tag_value, created) values(" & tag_Steam_wk_ist & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Steam_wk_ist = Nothing
Dim tag_Tanknumber
tag_Tanknumber = HMIRuntime.Tags("Tanknumber").Read
strSQL = "INSERT INTO z_tag_Tanknumber (tag_value, created) values(" & tag_Tanknumber & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tanknumber = Nothing
Dim tag_TBF05_run
tag_TBF05_run = HMIRuntime.Tags("TBF05_run").Read
strSQL = "INSERT INTO z_tag_TBF05_run (tag_value, created) values(" & tag_TBF05_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF05_run = Nothing
Dim tag_TBF05_seq
tag_TBF05_seq = HMIRuntime.Tags("TBF05_seq").Read
strSQL = "INSERT INTO z_tag_TBF05_seq (tag_value, created) values(" & tag_TBF05_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF05_seq = Nothing
Dim tag_TBF06_run
tag_TBF06_run = HMIRuntime.Tags("TBF06_run").Read
strSQL = "INSERT INTO z_tag_TBF06_run (tag_value, created) values(" & tag_TBF06_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF06_run = Nothing
Dim tag_TBF06_seq
tag_TBF06_seq = HMIRuntime.Tags("TBF06_seq").Read
strSQL = "INSERT INTO z_tag_TBF06_seq (tag_value, created) values(" & tag_TBF06_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF06_seq = Nothing
Dim tag_TETlanhnhanh_in
tag_TETlanhnhanh_in = HMIRuntime.Tags("TETlanhnhanh_in").Read
strSQL = "INSERT INTO z_tag_TETlanhnhanh_in (tag_value, created) values(" & tag_TETlanhnhanh_in & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TETlanhnhanh_in = Nothing
Dim tag_TETlow
tag_TETlow = HMIRuntime.Tags("TETlow").Read
strSQL = "INSERT INTO z_tag_TETlow (tag_value, created) values(" & tag_TETlow & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TETlow = Nothing
Dim tag_tinhieucaphoi
tag_tinhieucaphoi = HMIRuntime.Tags("tinhieucaphoi").Read
strSQL = "INSERT INTO z_tag_tinhieucaphoi (tag_value, created) values(" & tag_tinhieucaphoi & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_tinhieucaphoi = Nothing
Dim tag_TotalWort
tag_TotalWort = HMIRuntime.Tags("TotalWort").Read
strSQL = "INSERT INTO z_tag_TotalWort (tag_value, created) values(" & tag_TotalWort & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalWort = Nothing
Dim tag_TrubTank_run
tag_TrubTank_run = HMIRuntime.Tags("TrubTank_run").Read
strSQL = "INSERT INTO z_tag_TrubTank_run (tag_value, created) values(" & tag_TrubTank_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TrubTank_run = Nothing
Dim tag_TrubTank_seq
tag_TrubTank_seq = HMIRuntime.Tags("TrubTank_seq").Read
strSQL = "INSERT INTO z_tag_TrubTank_seq (tag_value, created) values(" & tag_TrubTank_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TrubTank_seq = Nothing
Dim tag_Valvecap Glycol
tag_Valvecap Glycol = HMIRuntime.Tags("Valvecap Glycol").Read
strSQL = "INSERT INTO z_tag_Valvecap Glycol (tag_value, created) values(" & tag_Valvecap Glycol & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Valvecap Glycol = Nothing
Dim tag_vancapnuocLN
tag_vancapnuocLN = HMIRuntime.Tags("vancapnuocLN").Read
strSQL = "INSERT INTO z_tag_vancapnuocLN (tag_value, created) values(" & tag_vancapnuocLN & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_vancapnuocLN = Nothing
Dim tag_vanhoi_noihoa_st
tag_vanhoi_noihoa_st = HMIRuntime.Tags("vanhoi_noihoa_st").Read
strSQL = "INSERT INTO z_tag_vanhoi_noihoa_st (tag_value, created) values(" & tag_vanhoi_noihoa_st & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_vanhoi_noihoa_st = Nothing
Dim tag_WhirlPool_Auto
tag_WhirlPool_Auto = HMIRuntime.Tags("WhirlPool_Auto").Read
strSQL = "INSERT INTO z_tag_WhirlPool_Auto (tag_value, created) values(" & tag_WhirlPool_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WhirlPool_Auto = Nothing
Dim tag_WhirlPool_CIP
tag_WhirlPool_CIP = HMIRuntime.Tags("WhirlPool_CIP").Read
strSQL = "INSERT INTO z_tag_WhirlPool_CIP (tag_value, created) values(" & tag_WhirlPool_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WhirlPool_CIP = Nothing
Dim tag_WhirlPool_run
tag_WhirlPool_run = HMIRuntime.Tags("WhirlPool_run").Read
strSQL = "INSERT INTO z_tag_WhirlPool_run (tag_value, created) values(" & tag_WhirlPool_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WhirlPool_run = Nothing
Dim tag_WhirlPool_seq
tag_WhirlPool_seq = HMIRuntime.Tags("WhirlPool_seq").Read
strSQL = "INSERT INTO z_tag_WhirlPool_seq (tag_value, created) values(" & tag_WhirlPool_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WhirlPool_seq = Nothing
Dim tag_WhirlPool_seq_add
tag_WhirlPool_seq_add = HMIRuntime.Tags("WhirlPool_seq_add").Read
strSQL = "INSERT INTO z_tag_WhirlPool_seq_add (tag_value, created) values(" & tag_WhirlPool_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WhirlPool_seq_add = Nothing
Dim tag_WortCooler_Auto
tag_WortCooler_Auto = HMIRuntime.Tags("WortCooler_Auto").Read
strSQL = "INSERT INTO z_tag_WortCooler_Auto (tag_value, created) values(" & tag_WortCooler_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortCooler_Auto = Nothing
Dim tag_WortCooler_CIP
tag_WortCooler_CIP = HMIRuntime.Tags("WortCooler_CIP").Read
strSQL = "INSERT INTO z_tag_WortCooler_CIP (tag_value, created) values(" & tag_WortCooler_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortCooler_CIP = Nothing
Dim tag_WortCooler_run
tag_WortCooler_run = HMIRuntime.Tags("WortCooler_run").Read
strSQL = "INSERT INTO z_tag_WortCooler_run (tag_value, created) values(" & tag_WortCooler_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortCooler_run = Nothing
Dim tag_WortCooler_seq
tag_WortCooler_seq = HMIRuntime.Tags("WortCooler_seq").Read
strSQL = "INSERT INTO z_tag_WortCooler_seq (tag_value, created) values(" & tag_WortCooler_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortCooler_seq = Nothing
Dim tag_WortCooler_seq_add
tag_WortCooler_seq_add = HMIRuntime.Tags("WortCooler_seq_add").Read
strSQL = "INSERT INTO z_tag_WortCooler_seq_add (tag_value, created) values(" & tag_WortCooler_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortCooler_seq_add = Nothing
Dim tag_WortKettle_Auto
tag_WortKettle_Auto = HMIRuntime.Tags("WortKettle_Auto").Read
strSQL = "INSERT INTO z_tag_WortKettle_Auto (tag_value, created) values(" & tag_WortKettle_Auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortKettle_Auto = Nothing
Dim tag_WortKettle_CIP
tag_WortKettle_CIP = HMIRuntime.Tags("WortKettle_CIP").Read
strSQL = "INSERT INTO z_tag_WortKettle_CIP (tag_value, created) values(" & tag_WortKettle_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortKettle_CIP = Nothing
Dim tag_WortKettle_run
tag_WortKettle_run = HMIRuntime.Tags("WortKettle_run").Read
strSQL = "INSERT INTO z_tag_WortKettle_run (tag_value, created) values(" & tag_WortKettle_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortKettle_run = Nothing
Dim tag_WortKettle_seq
tag_WortKettle_seq = HMIRuntime.Tags("WortKettle_seq").Read
strSQL = "INSERT INTO z_tag_WortKettle_seq (tag_value, created) values(" & tag_WortKettle_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortKettle_seq = Nothing
Dim tag_WortKettle_seq_add
tag_WortKettle_seq_add = HMIRuntime.Tags("WortKettle_seq_add").Read
strSQL = "INSERT INTO z_tag_WortKettle_seq_add (tag_value, created) values(" & tag_WortKettle_seq_add & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WortKettle_seq_add = Nothing
Dim tag_0301_TET_01
tag_0301_TET_01 = HMIRuntime.Tags("0301_TET_01").Read
strSQL = "INSERT INTO z_tag_0301_TET_01 (tag_value, created) values(" & tag_0301_TET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_TET_01 = Nothing
Dim tag_0301_TET_02
tag_0301_TET_02 = HMIRuntime.Tags("0301_TET_02").Read
strSQL = "INSERT INTO z_tag_0301_TET_02 (tag_value, created) values(" & tag_0301_TET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_TET_02 = Nothing
Dim tag_0301_WET_01
tag_0301_WET_01 = HMIRuntime.Tags("0301_WET_01").Read
strSQL = "INSERT INTO z_tag_0301_WET_01 (tag_value, created) values(" & tag_0301_WET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_WET_01 = Nothing
Dim tag_0301_WET_01_OUT
tag_0301_WET_01_OUT = HMIRuntime.Tags("0301_WET_01_OUT").Read
strSQL = "INSERT INTO z_tag_0301_WET_01_OUT (tag_value, created) values(" & tag_0301_WET_01_OUT & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_WET_01_OUT = Nothing
Dim tag_0301_WET_02
tag_0301_WET_02 = HMIRuntime.Tags("0301_WET_02").Read
strSQL = "INSERT INTO z_tag_0301_WET_02 (tag_value, created) values(" & tag_0301_WET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_WET_02 = Nothing
Dim tag_0301_WET_02_OUT
tag_0301_WET_02_OUT = HMIRuntime.Tags("0301_WET_02_OUT").Read
strSQL = "INSERT INTO z_tag_0301_WET_02_OUT (tag_value, created) values(" & tag_0301_WET_02_OUT & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301_WET_02_OUT = Nothing
Dim tag_0301M01
tag_0301M01 = HMIRuntime.Tags("0301M01").Read
strSQL = "INSERT INTO z_tag_0301M01 (tag_value, created) values(" & tag_0301M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301M01 = Nothing
Dim tag_0301M02
tag_0301M02 = HMIRuntime.Tags("0301M02").Read
strSQL = "INSERT INTO z_tag_0301M02 (tag_value, created) values(" & tag_0301M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301M02 = Nothing
Dim tag_0301M03
tag_0301M03 = HMIRuntime.Tags("0301M03").Read
strSQL = "INSERT INTO z_tag_0301M03 (tag_value, created) values(" & tag_0301M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301M03 = Nothing
Dim tag_0301M04
tag_0301M04 = HMIRuntime.Tags("0301M04").Read
strSQL = "INSERT INTO z_tag_0301M04 (tag_value, created) values(" & tag_0301M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301M04 = Nothing
Dim tag_0301PV01
tag_0301PV01 = HMIRuntime.Tags("0301PV01").Read
strSQL = "INSERT INTO z_tag_0301PV01 (tag_value, created) values(" & tag_0301PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV01 = Nothing
Dim tag_0301PV02
tag_0301PV02 = HMIRuntime.Tags("0301PV02").Read
strSQL = "INSERT INTO z_tag_0301PV02 (tag_value, created) values(" & tag_0301PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV02 = Nothing
Dim tag_0301PV03
tag_0301PV03 = HMIRuntime.Tags("0301PV03").Read
strSQL = "INSERT INTO z_tag_0301PV03 (tag_value, created) values(" & tag_0301PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV03 = Nothing
Dim tag_0301PV04
tag_0301PV04 = HMIRuntime.Tags("0301PV04").Read
strSQL = "INSERT INTO z_tag_0301PV04 (tag_value, created) values(" & tag_0301PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV04 = Nothing
Dim tag_0301PV05
tag_0301PV05 = HMIRuntime.Tags("0301PV05").Read
strSQL = "INSERT INTO z_tag_0301PV05 (tag_value, created) values(" & tag_0301PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV05 = Nothing
Dim tag_0301PV06
tag_0301PV06 = HMIRuntime.Tags("0301PV06").Read
strSQL = "INSERT INTO z_tag_0301PV06 (tag_value, created) values(" & tag_0301PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV06 = Nothing
Dim tag_0301PV07
tag_0301PV07 = HMIRuntime.Tags("0301PV07").Read
strSQL = "INSERT INTO z_tag_0301PV07 (tag_value, created) values(" & tag_0301PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV07 = Nothing
Dim tag_0301PV08
tag_0301PV08 = HMIRuntime.Tags("0301PV08").Read
strSQL = "INSERT INTO z_tag_0301PV08 (tag_value, created) values(" & tag_0301PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV08 = Nothing
Dim tag_0301PV09
tag_0301PV09 = HMIRuntime.Tags("0301PV09").Read
strSQL = "INSERT INTO z_tag_0301PV09 (tag_value, created) values(" & tag_0301PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV09 = Nothing
Dim tag_0301PV10
tag_0301PV10 = HMIRuntime.Tags("0301PV10").Read
strSQL = "INSERT INTO z_tag_0301PV10 (tag_value, created) values(" & tag_0301PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV10 = Nothing
Dim tag_0301PV11
tag_0301PV11 = HMIRuntime.Tags("0301PV11").Read
strSQL = "INSERT INTO z_tag_0301PV11 (tag_value, created) values(" & tag_0301PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV11 = Nothing
Dim tag_0301PV13
tag_0301PV13 = HMIRuntime.Tags("0301PV13").Read
strSQL = "INSERT INTO z_tag_0301PV13 (tag_value, created) values(" & tag_0301PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV13 = Nothing
Dim tag_0301PV14
tag_0301PV14 = HMIRuntime.Tags("0301PV14").Read
strSQL = "INSERT INTO z_tag_0301PV14 (tag_value, created) values(" & tag_0301PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0301PV14 = Nothing
Dim tag_0401_FQET_02
tag_0401_FQET_02 = HMIRuntime.Tags("0401_FQET_02").Read
strSQL = "INSERT INTO z_tag_0401_FQET_02 (tag_value, created) values(" & tag_0401_FQET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_FQET_02 = Nothing
Dim tag_0401_TET_11
tag_0401_TET_11 = HMIRuntime.Tags("0401_TET_11").Read
strSQL = "INSERT INTO z_tag_0401_TET_11 (tag_value, created) values(" & tag_0401_TET_11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_11 = Nothing
Dim tag_0401_TET_12
tag_0401_TET_12 = HMIRuntime.Tags("0401_TET_12").Read
strSQL = "INSERT INTO z_tag_0401_TET_12 (tag_value, created) values(" & tag_0401_TET_12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_12 = Nothing
Dim tag_0401_TET_21
tag_0401_TET_21 = HMIRuntime.Tags("0401_TET_21").Read
strSQL = "INSERT INTO z_tag_0401_TET_21 (tag_value, created) values(" & tag_0401_TET_21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_21 = Nothing
Dim tag_0401_TET_22
tag_0401_TET_22 = HMIRuntime.Tags("0401_TET_22").Read
strSQL = "INSERT INTO z_tag_0401_TET_22 (tag_value, created) values(" & tag_0401_TET_22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_22 = Nothing
Dim tag_0401_TET_31
tag_0401_TET_31 = HMIRuntime.Tags("0401_TET_31").Read
strSQL = "INSERT INTO z_tag_0401_TET_31 (tag_value, created) values(" & tag_0401_TET_31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_31 = Nothing
Dim tag_0401_TET_32
tag_0401_TET_32 = HMIRuntime.Tags("0401_TET_32").Read
strSQL = "INSERT INTO z_tag_0401_TET_32 (tag_value, created) values(" & tag_0401_TET_32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_32 = Nothing
Dim tag_0401_TET_41
tag_0401_TET_41 = HMIRuntime.Tags("0401_TET_41").Read
strSQL = "INSERT INTO z_tag_0401_TET_41 (tag_value, created) values(" & tag_0401_TET_41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_41 = Nothing
Dim tag_0401_TET_42
tag_0401_TET_42 = HMIRuntime.Tags("0401_TET_42").Read
strSQL = "INSERT INTO z_tag_0401_TET_42 (tag_value, created) values(" & tag_0401_TET_42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_42 = Nothing
Dim tag_0401_TET_51
tag_0401_TET_51 = HMIRuntime.Tags("0401_TET_51").Read
strSQL = "INSERT INTO z_tag_0401_TET_51 (tag_value, created) values(" & tag_0401_TET_51 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_51 = Nothing
Dim tag_0401_TET_52
tag_0401_TET_52 = HMIRuntime.Tags("0401_TET_52").Read
strSQL = "INSERT INTO z_tag_0401_TET_52 (tag_value, created) values(" & tag_0401_TET_52 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_52 = Nothing
Dim tag_0401_TET_61
tag_0401_TET_61 = HMIRuntime.Tags("0401_TET_61").Read
strSQL = "INSERT INTO z_tag_0401_TET_61 (tag_value, created) values(" & tag_0401_TET_61 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_61 = Nothing
Dim tag_0401_TET_62
tag_0401_TET_62 = HMIRuntime.Tags("0401_TET_62").Read
strSQL = "INSERT INTO z_tag_0401_TET_62 (tag_value, created) values(" & tag_0401_TET_62 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_62 = Nothing
Dim tag_0401_TET_71
tag_0401_TET_71 = HMIRuntime.Tags("0401_TET_71").Read
strSQL = "INSERT INTO z_tag_0401_TET_71 (tag_value, created) values(" & tag_0401_TET_71 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_71 = Nothing
Dim tag_0401_TET_72
tag_0401_TET_72 = HMIRuntime.Tags("0401_TET_72").Read
strSQL = "INSERT INTO z_tag_0401_TET_72 (tag_value, created) values(" & tag_0401_TET_72 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_72 = Nothing
Dim tag_0401_TET_81
tag_0401_TET_81 = HMIRuntime.Tags("0401_TET_81").Read
strSQL = "INSERT INTO z_tag_0401_TET_81 (tag_value, created) values(" & tag_0401_TET_81 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_81 = Nothing
Dim tag_0401_TET_82
tag_0401_TET_82 = HMIRuntime.Tags("0401_TET_82").Read
strSQL = "INSERT INTO z_tag_0401_TET_82 (tag_value, created) values(" & tag_0401_TET_82 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_82 = Nothing
Dim tag_0401_TET_91
tag_0401_TET_91 = HMIRuntime.Tags("0401_TET_91").Read
strSQL = "INSERT INTO z_tag_0401_TET_91 (tag_value, created) values(" & tag_0401_TET_91 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_91 = Nothing
Dim tag_0401_TET_92
tag_0401_TET_92 = HMIRuntime.Tags("0401_TET_92").Read
strSQL = "INSERT INTO z_tag_0401_TET_92 (tag_value, created) values(" & tag_0401_TET_92 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_92 = Nothing
Dim tag_0401_TET_101
tag_0401_TET_101 = HMIRuntime.Tags("0401_TET_101").Read
strSQL = "INSERT INTO z_tag_0401_TET_101 (tag_value, created) values(" & tag_0401_TET_101 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_101 = Nothing
Dim tag_0401_TET_102
tag_0401_TET_102 = HMIRuntime.Tags("0401_TET_102").Read
strSQL = "INSERT INTO z_tag_0401_TET_102 (tag_value, created) values(" & tag_0401_TET_102 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_102 = Nothing
Dim tag_0401_TET_111
tag_0401_TET_111 = HMIRuntime.Tags("0401_TET_111").Read
strSQL = "INSERT INTO z_tag_0401_TET_111 (tag_value, created) values(" & tag_0401_TET_111 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_111 = Nothing
Dim tag_0401_TET_112
tag_0401_TET_112 = HMIRuntime.Tags("0401_TET_112").Read
strSQL = "INSERT INTO z_tag_0401_TET_112 (tag_value, created) values(" & tag_0401_TET_112 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_112 = Nothing
Dim tag_0401_TET_121
tag_0401_TET_121 = HMIRuntime.Tags("0401_TET_121").Read
strSQL = "INSERT INTO z_tag_0401_TET_121 (tag_value, created) values(" & tag_0401_TET_121 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_121 = Nothing
Dim tag_0401_TET_122
tag_0401_TET_122 = HMIRuntime.Tags("0401_TET_122").Read
strSQL = "INSERT INTO z_tag_0401_TET_122 (tag_value, created) values(" & tag_0401_TET_122 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_122 = Nothing
Dim tag_0401_TET_131
tag_0401_TET_131 = HMIRuntime.Tags("0401_TET_131").Read
strSQL = "INSERT INTO z_tag_0401_TET_131 (tag_value, created) values(" & tag_0401_TET_131 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_131 = Nothing
Dim tag_0401_TET_132
tag_0401_TET_132 = HMIRuntime.Tags("0401_TET_132").Read
strSQL = "INSERT INTO z_tag_0401_TET_132 (tag_value, created) values(" & tag_0401_TET_132 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_132 = Nothing
Dim tag_0401_TET_141
tag_0401_TET_141 = HMIRuntime.Tags("0401_TET_141").Read
strSQL = "INSERT INTO z_tag_0401_TET_141 (tag_value, created) values(" & tag_0401_TET_141 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_141 = Nothing
Dim tag_0401_TET_142
tag_0401_TET_142 = HMIRuntime.Tags("0401_TET_142").Read
strSQL = "INSERT INTO z_tag_0401_TET_142 (tag_value, created) values(" & tag_0401_TET_142 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_142 = Nothing
Dim tag_0401_TET_151
tag_0401_TET_151 = HMIRuntime.Tags("0401_TET_151").Read
strSQL = "INSERT INTO z_tag_0401_TET_151 (tag_value, created) values(" & tag_0401_TET_151 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_151 = Nothing
Dim tag_0401_TET_152
tag_0401_TET_152 = HMIRuntime.Tags("0401_TET_152").Read
strSQL = "INSERT INTO z_tag_0401_TET_152 (tag_value, created) values(" & tag_0401_TET_152 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_152 = Nothing
Dim tag_0401_TET_161
tag_0401_TET_161 = HMIRuntime.Tags("0401_TET_161").Read
strSQL = "INSERT INTO z_tag_0401_TET_161 (tag_value, created) values(" & tag_0401_TET_161 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_161 = Nothing
Dim tag_0401_TET_162
tag_0401_TET_162 = HMIRuntime.Tags("0401_TET_162").Read
strSQL = "INSERT INTO z_tag_0401_TET_162 (tag_value, created) values(" & tag_0401_TET_162 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_162 = Nothing
Dim tag_0401_TET_171
tag_0401_TET_171 = HMIRuntime.Tags("0401_TET_171").Read
strSQL = "INSERT INTO z_tag_0401_TET_171 (tag_value, created) values(" & tag_0401_TET_171 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_171 = Nothing
Dim tag_0401_TET_172
tag_0401_TET_172 = HMIRuntime.Tags("0401_TET_172").Read
strSQL = "INSERT INTO z_tag_0401_TET_172 (tag_value, created) values(" & tag_0401_TET_172 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_172 = Nothing
Dim tag_0401_TET_181
tag_0401_TET_181 = HMIRuntime.Tags("0401_TET_181").Read
strSQL = "INSERT INTO z_tag_0401_TET_181 (tag_value, created) values(" & tag_0401_TET_181 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_181 = Nothing
Dim tag_0401_TET_182
tag_0401_TET_182 = HMIRuntime.Tags("0401_TET_182").Read
strSQL = "INSERT INTO z_tag_0401_TET_182 (tag_value, created) values(" & tag_0401_TET_182 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_182 = Nothing
Dim tag_0401_TET_191
tag_0401_TET_191 = HMIRuntime.Tags("0401_TET_191").Read
strSQL = "INSERT INTO z_tag_0401_TET_191 (tag_value, created) values(" & tag_0401_TET_191 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_191 = Nothing
Dim tag_0401_TET_192
tag_0401_TET_192 = HMIRuntime.Tags("0401_TET_192").Read
strSQL = "INSERT INTO z_tag_0401_TET_192 (tag_value, created) values(" & tag_0401_TET_192 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_192 = Nothing
Dim tag_0401_TET_201
tag_0401_TET_201 = HMIRuntime.Tags("0401_TET_201").Read
strSQL = "INSERT INTO z_tag_0401_TET_201 (tag_value, created) values(" & tag_0401_TET_201 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_201 = Nothing
Dim tag_0401_TET_202
tag_0401_TET_202 = HMIRuntime.Tags("0401_TET_202").Read
strSQL = "INSERT INTO z_tag_0401_TET_202 (tag_value, created) values(" & tag_0401_TET_202 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_202 = Nothing
Dim tag_0401_TET_211
tag_0401_TET_211 = HMIRuntime.Tags("0401_TET_211").Read
strSQL = "INSERT INTO z_tag_0401_TET_211 (tag_value, created) values(" & tag_0401_TET_211 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_211 = Nothing
Dim tag_0401_TET_212
tag_0401_TET_212 = HMIRuntime.Tags("0401_TET_212").Read
strSQL = "INSERT INTO z_tag_0401_TET_212 (tag_value, created) values(" & tag_0401_TET_212 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_212 = Nothing
Dim tag_0401_TET_221
tag_0401_TET_221 = HMIRuntime.Tags("0401_TET_221").Read
strSQL = "INSERT INTO z_tag_0401_TET_221 (tag_value, created) values(" & tag_0401_TET_221 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_221 = Nothing
Dim tag_0401_TET_222
tag_0401_TET_222 = HMIRuntime.Tags("0401_TET_222").Read
strSQL = "INSERT INTO z_tag_0401_TET_222 (tag_value, created) values(" & tag_0401_TET_222 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_222 = Nothing
Dim tag_0401_TET_231
tag_0401_TET_231 = HMIRuntime.Tags("0401_TET_231").Read
strSQL = "INSERT INTO z_tag_0401_TET_231 (tag_value, created) values(" & tag_0401_TET_231 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_231 = Nothing
Dim tag_0401_TET_232
tag_0401_TET_232 = HMIRuntime.Tags("0401_TET_232").Read
strSQL = "INSERT INTO z_tag_0401_TET_232 (tag_value, created) values(" & tag_0401_TET_232 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_232 = Nothing
Dim tag_0401_TET_241
tag_0401_TET_241 = HMIRuntime.Tags("0401_TET_241").Read
strSQL = "INSERT INTO z_tag_0401_TET_241 (tag_value, created) values(" & tag_0401_TET_241 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_241 = Nothing
Dim tag_0401_TET_242
tag_0401_TET_242 = HMIRuntime.Tags("0401_TET_242").Read
strSQL = "INSERT INTO z_tag_0401_TET_242 (tag_value, created) values(" & tag_0401_TET_242 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_242 = Nothing
Dim tag_0401_TET_251
tag_0401_TET_251 = HMIRuntime.Tags("0401_TET_251").Read
strSQL = "INSERT INTO z_tag_0401_TET_251 (tag_value, created) values(" & tag_0401_TET_251 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_251 = Nothing
Dim tag_0401_TET_252
tag_0401_TET_252 = HMIRuntime.Tags("0401_TET_252").Read
strSQL = "INSERT INTO z_tag_0401_TET_252 (tag_value, created) values(" & tag_0401_TET_252 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_252 = Nothing
Dim tag_0401_TET_261
tag_0401_TET_261 = HMIRuntime.Tags("0401_TET_261").Read
strSQL = "INSERT INTO z_tag_0401_TET_261 (tag_value, created) values(" & tag_0401_TET_261 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_261 = Nothing
Dim tag_0401_TET_262
tag_0401_TET_262 = HMIRuntime.Tags("0401_TET_262").Read
strSQL = "INSERT INTO z_tag_0401_TET_262 (tag_value, created) values(" & tag_0401_TET_262 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_262 = Nothing
Dim tag_0401_TET_271
tag_0401_TET_271 = HMIRuntime.Tags("0401_TET_271").Read
strSQL = "INSERT INTO z_tag_0401_TET_271 (tag_value, created) values(" & tag_0401_TET_271 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_271 = Nothing
Dim tag_0401_TET_272
tag_0401_TET_272 = HMIRuntime.Tags("0401_TET_272").Read
strSQL = "INSERT INTO z_tag_0401_TET_272 (tag_value, created) values(" & tag_0401_TET_272 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_272 = Nothing
Dim tag_0401_TET_281
tag_0401_TET_281 = HMIRuntime.Tags("0401_TET_281").Read
strSQL = "INSERT INTO z_tag_0401_TET_281 (tag_value, created) values(" & tag_0401_TET_281 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_281 = Nothing
Dim tag_0401_TET_282
tag_0401_TET_282 = HMIRuntime.Tags("0401_TET_282").Read
strSQL = "INSERT INTO z_tag_0401_TET_282 (tag_value, created) values(" & tag_0401_TET_282 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET_282 = Nothing
Dim tag_0401_TET29
tag_0401_TET29 = HMIRuntime.Tags("0401_TET29").Read
strSQL = "INSERT INTO z_tag_0401_TET29 (tag_value, created) values(" & tag_0401_TET29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET29 = Nothing
Dim tag_0401_TET30
tag_0401_TET30 = HMIRuntime.Tags("0401_TET30").Read
strSQL = "INSERT INTO z_tag_0401_TET30 (tag_value, created) values(" & tag_0401_TET30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET30 = Nothing
Dim tag_0401_TET31_D
tag_0401_TET31_D = HMIRuntime.Tags("0401_TET31_D").Read
strSQL = "INSERT INTO z_tag_0401_TET31_D (tag_value, created) values(" & tag_0401_TET31_D & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET31_D = Nothing
Dim tag_0401_TET31_T
tag_0401_TET31_T = HMIRuntime.Tags("0401_TET31_T").Read
strSQL = "INSERT INTO z_tag_0401_TET31_T (tag_value, created) values(" & tag_0401_TET31_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET31_T = Nothing
Dim tag_0401_TET32_D
tag_0401_TET32_D = HMIRuntime.Tags("0401_TET32_D").Read
strSQL = "INSERT INTO z_tag_0401_TET32_D (tag_value, created) values(" & tag_0401_TET32_D & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET32_D = Nothing
Dim tag_0401_TET32_T
tag_0401_TET32_T = HMIRuntime.Tags("0401_TET32_T").Read
strSQL = "INSERT INTO z_tag_0401_TET32_T (tag_value, created) values(" & tag_0401_TET32_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET32_T = Nothing
Dim tag_0401_TET33_D
tag_0401_TET33_D = HMIRuntime.Tags("0401_TET33_D").Read
strSQL = "INSERT INTO z_tag_0401_TET33_D (tag_value, created) values(" & tag_0401_TET33_D & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET33_D = Nothing
Dim tag_0401_TET33_T
tag_0401_TET33_T = HMIRuntime.Tags("0401_TET33_T").Read
strSQL = "INSERT INTO z_tag_0401_TET33_T (tag_value, created) values(" & tag_0401_TET33_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET33_T = Nothing
Dim tag_0401_TET34_D
tag_0401_TET34_D = HMIRuntime.Tags("0401_TET34_D").Read
strSQL = "INSERT INTO z_tag_0401_TET34_D (tag_value, created) values(" & tag_0401_TET34_D & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET34_D = Nothing
Dim tag_0401_TET34_T
tag_0401_TET34_T = HMIRuntime.Tags("0401_TET34_T").Read
strSQL = "INSERT INTO z_tag_0401_TET34_T (tag_value, created) values(" & tag_0401_TET34_T & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401_TET34_T = Nothing
Dim tag_0401LSL01
tag_0401LSL01 = HMIRuntime.Tags("0401LSL01").Read
strSQL = "INSERT INTO z_tag_0401LSL01 (tag_value, created) values(" & tag_0401LSL01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL01 = Nothing
Dim tag_0401LSL02
tag_0401LSL02 = HMIRuntime.Tags("0401LSL02").Read
strSQL = "INSERT INTO z_tag_0401LSL02 (tag_value, created) values(" & tag_0401LSL02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL02 = Nothing
Dim tag_0401LSL03
tag_0401LSL03 = HMIRuntime.Tags("0401LSL03").Read
strSQL = "INSERT INTO z_tag_0401LSL03 (tag_value, created) values(" & tag_0401LSL03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL03 = Nothing
Dim tag_0401LSL04
tag_0401LSL04 = HMIRuntime.Tags("0401LSL04").Read
strSQL = "INSERT INTO z_tag_0401LSL04 (tag_value, created) values(" & tag_0401LSL04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL04 = Nothing
Dim tag_0401LSL05
tag_0401LSL05 = HMIRuntime.Tags("0401LSL05").Read
strSQL = "INSERT INTO z_tag_0401LSL05 (tag_value, created) values(" & tag_0401LSL05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL05 = Nothing
Dim tag_0401LSL06
tag_0401LSL06 = HMIRuntime.Tags("0401LSL06").Read
strSQL = "INSERT INTO z_tag_0401LSL06 (tag_value, created) values(" & tag_0401LSL06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL06 = Nothing
Dim tag_0401LSL07
tag_0401LSL07 = HMIRuntime.Tags("0401LSL07").Read
strSQL = "INSERT INTO z_tag_0401LSL07 (tag_value, created) values(" & tag_0401LSL07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL07 = Nothing
Dim tag_0401LSL08
tag_0401LSL08 = HMIRuntime.Tags("0401LSL08").Read
strSQL = "INSERT INTO z_tag_0401LSL08 (tag_value, created) values(" & tag_0401LSL08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL08 = Nothing
Dim tag_0401LSL09
tag_0401LSL09 = HMIRuntime.Tags("0401LSL09").Read
strSQL = "INSERT INTO z_tag_0401LSL09 (tag_value, created) values(" & tag_0401LSL09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL09 = Nothing
Dim tag_0401LSL10
tag_0401LSL10 = HMIRuntime.Tags("0401LSL10").Read
strSQL = "INSERT INTO z_tag_0401LSL10 (tag_value, created) values(" & tag_0401LSL10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL10 = Nothing
Dim tag_0401LSL11
tag_0401LSL11 = HMIRuntime.Tags("0401LSL11").Read
strSQL = "INSERT INTO z_tag_0401LSL11 (tag_value, created) values(" & tag_0401LSL11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL11 = Nothing
Dim tag_0401LSL12
tag_0401LSL12 = HMIRuntime.Tags("0401LSL12").Read
strSQL = "INSERT INTO z_tag_0401LSL12 (tag_value, created) values(" & tag_0401LSL12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL12 = Nothing
Dim tag_0401LSL13
tag_0401LSL13 = HMIRuntime.Tags("0401LSL13").Read
strSQL = "INSERT INTO z_tag_0401LSL13 (tag_value, created) values(" & tag_0401LSL13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL13 = Nothing
Dim tag_0401LSL14
tag_0401LSL14 = HMIRuntime.Tags("0401LSL14").Read
strSQL = "INSERT INTO z_tag_0401LSL14 (tag_value, created) values(" & tag_0401LSL14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL14 = Nothing
Dim tag_0401LSL15
tag_0401LSL15 = HMIRuntime.Tags("0401LSL15").Read
strSQL = "INSERT INTO z_tag_0401LSL15 (tag_value, created) values(" & tag_0401LSL15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL15 = Nothing
Dim tag_0401LSL16
tag_0401LSL16 = HMIRuntime.Tags("0401LSL16").Read
strSQL = "INSERT INTO z_tag_0401LSL16 (tag_value, created) values(" & tag_0401LSL16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL16 = Nothing
Dim tag_0401LSL17
tag_0401LSL17 = HMIRuntime.Tags("0401LSL17").Read
strSQL = "INSERT INTO z_tag_0401LSL17 (tag_value, created) values(" & tag_0401LSL17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL17 = Nothing
Dim tag_0401LSL18
tag_0401LSL18 = HMIRuntime.Tags("0401LSL18").Read
strSQL = "INSERT INTO z_tag_0401LSL18 (tag_value, created) values(" & tag_0401LSL18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL18 = Nothing
Dim tag_0401LSL19
tag_0401LSL19 = HMIRuntime.Tags("0401LSL19").Read
strSQL = "INSERT INTO z_tag_0401LSL19 (tag_value, created) values(" & tag_0401LSL19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL19 = Nothing
Dim tag_0401LSL20
tag_0401LSL20 = HMIRuntime.Tags("0401LSL20").Read
strSQL = "INSERT INTO z_tag_0401LSL20 (tag_value, created) values(" & tag_0401LSL20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL20 = Nothing
Dim tag_0401LSL21
tag_0401LSL21 = HMIRuntime.Tags("0401LSL21").Read
strSQL = "INSERT INTO z_tag_0401LSL21 (tag_value, created) values(" & tag_0401LSL21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL21 = Nothing
Dim tag_0401LSL22
tag_0401LSL22 = HMIRuntime.Tags("0401LSL22").Read
strSQL = "INSERT INTO z_tag_0401LSL22 (tag_value, created) values(" & tag_0401LSL22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL22 = Nothing
Dim tag_0401LSL23
tag_0401LSL23 = HMIRuntime.Tags("0401LSL23").Read
strSQL = "INSERT INTO z_tag_0401LSL23 (tag_value, created) values(" & tag_0401LSL23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL23 = Nothing
Dim tag_0401LSL24
tag_0401LSL24 = HMIRuntime.Tags("0401LSL24").Read
strSQL = "INSERT INTO z_tag_0401LSL24 (tag_value, created) values(" & tag_0401LSL24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL24 = Nothing
Dim tag_0401LSL25
tag_0401LSL25 = HMIRuntime.Tags("0401LSL25").Read
strSQL = "INSERT INTO z_tag_0401LSL25 (tag_value, created) values(" & tag_0401LSL25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL25 = Nothing
Dim tag_0401LSL26
tag_0401LSL26 = HMIRuntime.Tags("0401LSL26").Read
strSQL = "INSERT INTO z_tag_0401LSL26 (tag_value, created) values(" & tag_0401LSL26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL26 = Nothing
Dim tag_0401LSL27
tag_0401LSL27 = HMIRuntime.Tags("0401LSL27").Read
strSQL = "INSERT INTO z_tag_0401LSL27 (tag_value, created) values(" & tag_0401LSL27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL27 = Nothing
Dim tag_0401LSL28
tag_0401LSL28 = HMIRuntime.Tags("0401LSL28").Read
strSQL = "INSERT INTO z_tag_0401LSL28 (tag_value, created) values(" & tag_0401LSL28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL28 = Nothing
Dim tag_0401LSL29
tag_0401LSL29 = HMIRuntime.Tags("0401LSL29").Read
strSQL = "INSERT INTO z_tag_0401LSL29 (tag_value, created) values(" & tag_0401LSL29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL29 = Nothing
Dim tag_0401LSL30
tag_0401LSL30 = HMIRuntime.Tags("0401LSL30").Read
strSQL = "INSERT INTO z_tag_0401LSL30 (tag_value, created) values(" & tag_0401LSL30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401LSL30 = Nothing
Dim tag_0401M01
tag_0401M01 = HMIRuntime.Tags("0401M01").Read
strSQL = "INSERT INTO z_tag_0401M01 (tag_value, created) values(" & tag_0401M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401M01 = Nothing
Dim tag_0401M02
tag_0401M02 = HMIRuntime.Tags("0401M02").Read
strSQL = "INSERT INTO z_tag_0401M02 (tag_value, created) values(" & tag_0401M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401M02 = Nothing
Dim tag_0401M03
tag_0401M03 = HMIRuntime.Tags("0401M03").Read
strSQL = "INSERT INTO z_tag_0401M03 (tag_value, created) values(" & tag_0401M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401M03 = Nothing
Dim tag_0401M05
tag_0401M05 = HMIRuntime.Tags("0401M05").Read
strSQL = "INSERT INTO z_tag_0401M05 (tag_value, created) values(" & tag_0401M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401M05 = Nothing
Dim tag_0401PV01
tag_0401PV01 = HMIRuntime.Tags("0401PV01").Read
strSQL = "INSERT INTO z_tag_0401PV01 (tag_value, created) values(" & tag_0401PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV01 = Nothing
Dim tag_0401PV1_1
tag_0401PV1_1 = HMIRuntime.Tags("0401PV1_1").Read
strSQL = "INSERT INTO z_tag_0401PV1_1 (tag_value, created) values(" & tag_0401PV1_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV1_1 = Nothing
Dim tag_0401PV1_2
tag_0401PV1_2 = HMIRuntime.Tags("0401PV1_2").Read
strSQL = "INSERT INTO z_tag_0401PV1_2 (tag_value, created) values(" & tag_0401PV1_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV1_2 = Nothing
Dim tag_0401PV1_3
tag_0401PV1_3 = HMIRuntime.Tags("0401PV1_3").Read
strSQL = "INSERT INTO z_tag_0401PV1_3 (tag_value, created) values(" & tag_0401PV1_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV1_3 = Nothing
Dim tag_0401PV1_4
tag_0401PV1_4 = HMIRuntime.Tags("0401PV1_4").Read
strSQL = "INSERT INTO z_tag_0401PV1_4 (tag_value, created) values(" & tag_0401PV1_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV1_4 = Nothing
Dim tag_0401PV1_5
tag_0401PV1_5 = HMIRuntime.Tags("0401PV1_5").Read
strSQL = "INSERT INTO z_tag_0401PV1_5 (tag_value, created) values(" & tag_0401PV1_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV1_5 = Nothing
Dim tag_0401PV02
tag_0401PV02 = HMIRuntime.Tags("0401PV02").Read
strSQL = "INSERT INTO z_tag_0401PV02 (tag_value, created) values(" & tag_0401PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV02 = Nothing
Dim tag_0401PV2_1
tag_0401PV2_1 = HMIRuntime.Tags("0401PV2_1").Read
strSQL = "INSERT INTO z_tag_0401PV2_1 (tag_value, created) values(" & tag_0401PV2_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV2_1 = Nothing
Dim tag_0401PV2_2
tag_0401PV2_2 = HMIRuntime.Tags("0401PV2_2").Read
strSQL = "INSERT INTO z_tag_0401PV2_2 (tag_value, created) values(" & tag_0401PV2_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV2_2 = Nothing
Dim tag_0401PV2_3
tag_0401PV2_3 = HMIRuntime.Tags("0401PV2_3").Read
strSQL = "INSERT INTO z_tag_0401PV2_3 (tag_value, created) values(" & tag_0401PV2_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV2_3 = Nothing
Dim tag_0401PV2_4
tag_0401PV2_4 = HMIRuntime.Tags("0401PV2_4").Read
strSQL = "INSERT INTO z_tag_0401PV2_4 (tag_value, created) values(" & tag_0401PV2_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV2_4 = Nothing
Dim tag_0401PV2_5
tag_0401PV2_5 = HMIRuntime.Tags("0401PV2_5").Read
strSQL = "INSERT INTO z_tag_0401PV2_5 (tag_value, created) values(" & tag_0401PV2_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV2_5 = Nothing
Dim tag_0401PV03
tag_0401PV03 = HMIRuntime.Tags("0401PV03").Read
strSQL = "INSERT INTO z_tag_0401PV03 (tag_value, created) values(" & tag_0401PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV03 = Nothing
Dim tag_0401PV3_1
tag_0401PV3_1 = HMIRuntime.Tags("0401PV3_1").Read
strSQL = "INSERT INTO z_tag_0401PV3_1 (tag_value, created) values(" & tag_0401PV3_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV3_1 = Nothing
Dim tag_0401PV3_2
tag_0401PV3_2 = HMIRuntime.Tags("0401PV3_2").Read
strSQL = "INSERT INTO z_tag_0401PV3_2 (tag_value, created) values(" & tag_0401PV3_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV3_2 = Nothing
Dim tag_0401PV3_3
tag_0401PV3_3 = HMIRuntime.Tags("0401PV3_3").Read
strSQL = "INSERT INTO z_tag_0401PV3_3 (tag_value, created) values(" & tag_0401PV3_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV3_3 = Nothing
Dim tag_0401PV3_4
tag_0401PV3_4 = HMIRuntime.Tags("0401PV3_4").Read
strSQL = "INSERT INTO z_tag_0401PV3_4 (tag_value, created) values(" & tag_0401PV3_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV3_4 = Nothing
Dim tag_0401PV3_5
tag_0401PV3_5 = HMIRuntime.Tags("0401PV3_5").Read
strSQL = "INSERT INTO z_tag_0401PV3_5 (tag_value, created) values(" & tag_0401PV3_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV3_5 = Nothing
Dim tag_0401PV04
tag_0401PV04 = HMIRuntime.Tags("0401PV04").Read
strSQL = "INSERT INTO z_tag_0401PV04 (tag_value, created) values(" & tag_0401PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV04 = Nothing
Dim tag_0401PV4_1
tag_0401PV4_1 = HMIRuntime.Tags("0401PV4_1").Read
strSQL = "INSERT INTO z_tag_0401PV4_1 (tag_value, created) values(" & tag_0401PV4_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV4_1 = Nothing
Dim tag_0401PV4_2
tag_0401PV4_2 = HMIRuntime.Tags("0401PV4_2").Read
strSQL = "INSERT INTO z_tag_0401PV4_2 (tag_value, created) values(" & tag_0401PV4_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV4_2 = Nothing
Dim tag_0401PV4_3
tag_0401PV4_3 = HMIRuntime.Tags("0401PV4_3").Read
strSQL = "INSERT INTO z_tag_0401PV4_3 (tag_value, created) values(" & tag_0401PV4_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV4_3 = Nothing
Dim tag_0401PV4_4
tag_0401PV4_4 = HMIRuntime.Tags("0401PV4_4").Read
strSQL = "INSERT INTO z_tag_0401PV4_4 (tag_value, created) values(" & tag_0401PV4_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV4_4 = Nothing
Dim tag_0401PV4_5
tag_0401PV4_5 = HMIRuntime.Tags("0401PV4_5").Read
strSQL = "INSERT INTO z_tag_0401PV4_5 (tag_value, created) values(" & tag_0401PV4_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV4_5 = Nothing
Dim tag_0401PV05
tag_0401PV05 = HMIRuntime.Tags("0401PV05").Read
strSQL = "INSERT INTO z_tag_0401PV05 (tag_value, created) values(" & tag_0401PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV05 = Nothing
Dim tag_0401PV5_1
tag_0401PV5_1 = HMIRuntime.Tags("0401PV5_1").Read
strSQL = "INSERT INTO z_tag_0401PV5_1 (tag_value, created) values(" & tag_0401PV5_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV5_1 = Nothing
Dim tag_0401PV5_2
tag_0401PV5_2 = HMIRuntime.Tags("0401PV5_2").Read
strSQL = "INSERT INTO z_tag_0401PV5_2 (tag_value, created) values(" & tag_0401PV5_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV5_2 = Nothing
Dim tag_0401PV5_3
tag_0401PV5_3 = HMIRuntime.Tags("0401PV5_3").Read
strSQL = "INSERT INTO z_tag_0401PV5_3 (tag_value, created) values(" & tag_0401PV5_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV5_3 = Nothing
Dim tag_0401PV5_4
tag_0401PV5_4 = HMIRuntime.Tags("0401PV5_4").Read
strSQL = "INSERT INTO z_tag_0401PV5_4 (tag_value, created) values(" & tag_0401PV5_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV5_4 = Nothing
Dim tag_0401PV5_5
tag_0401PV5_5 = HMIRuntime.Tags("0401PV5_5").Read
strSQL = "INSERT INTO z_tag_0401PV5_5 (tag_value, created) values(" & tag_0401PV5_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV5_5 = Nothing
Dim tag_0401PV06
tag_0401PV06 = HMIRuntime.Tags("0401PV06").Read
strSQL = "INSERT INTO z_tag_0401PV06 (tag_value, created) values(" & tag_0401PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV06 = Nothing
Dim tag_0401PV6_1
tag_0401PV6_1 = HMIRuntime.Tags("0401PV6_1").Read
strSQL = "INSERT INTO z_tag_0401PV6_1 (tag_value, created) values(" & tag_0401PV6_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV6_1 = Nothing
Dim tag_0401PV6_2
tag_0401PV6_2 = HMIRuntime.Tags("0401PV6_2").Read
strSQL = "INSERT INTO z_tag_0401PV6_2 (tag_value, created) values(" & tag_0401PV6_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV6_2 = Nothing
Dim tag_0401PV6_3
tag_0401PV6_3 = HMIRuntime.Tags("0401PV6_3").Read
strSQL = "INSERT INTO z_tag_0401PV6_3 (tag_value, created) values(" & tag_0401PV6_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV6_3 = Nothing
Dim tag_0401PV6_4
tag_0401PV6_4 = HMIRuntime.Tags("0401PV6_4").Read
strSQL = "INSERT INTO z_tag_0401PV6_4 (tag_value, created) values(" & tag_0401PV6_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV6_4 = Nothing
Dim tag_0401PV6_5
tag_0401PV6_5 = HMIRuntime.Tags("0401PV6_5").Read
strSQL = "INSERT INTO z_tag_0401PV6_5 (tag_value, created) values(" & tag_0401PV6_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV6_5 = Nothing
Dim tag_0401PV07
tag_0401PV07 = HMIRuntime.Tags("0401PV07").Read
strSQL = "INSERT INTO z_tag_0401PV07 (tag_value, created) values(" & tag_0401PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV07 = Nothing
Dim tag_0401PV7_1
tag_0401PV7_1 = HMIRuntime.Tags("0401PV7_1").Read
strSQL = "INSERT INTO z_tag_0401PV7_1 (tag_value, created) values(" & tag_0401PV7_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV7_1 = Nothing
Dim tag_0401PV7_2
tag_0401PV7_2 = HMIRuntime.Tags("0401PV7_2").Read
strSQL = "INSERT INTO z_tag_0401PV7_2 (tag_value, created) values(" & tag_0401PV7_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV7_2 = Nothing
Dim tag_0401PV7_3
tag_0401PV7_3 = HMIRuntime.Tags("0401PV7_3").Read
strSQL = "INSERT INTO z_tag_0401PV7_3 (tag_value, created) values(" & tag_0401PV7_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV7_3 = Nothing
Dim tag_0401PV7_4
tag_0401PV7_4 = HMIRuntime.Tags("0401PV7_4").Read
strSQL = "INSERT INTO z_tag_0401PV7_4 (tag_value, created) values(" & tag_0401PV7_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV7_4 = Nothing
Dim tag_0401PV7_5
tag_0401PV7_5 = HMIRuntime.Tags("0401PV7_5").Read
strSQL = "INSERT INTO z_tag_0401PV7_5 (tag_value, created) values(" & tag_0401PV7_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV7_5 = Nothing
Dim tag_0401PV08
tag_0401PV08 = HMIRuntime.Tags("0401PV08").Read
strSQL = "INSERT INTO z_tag_0401PV08 (tag_value, created) values(" & tag_0401PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV08 = Nothing
Dim tag_0401PV8_1
tag_0401PV8_1 = HMIRuntime.Tags("0401PV8_1").Read
strSQL = "INSERT INTO z_tag_0401PV8_1 (tag_value, created) values(" & tag_0401PV8_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV8_1 = Nothing
Dim tag_0401PV8_2
tag_0401PV8_2 = HMIRuntime.Tags("0401PV8_2").Read
strSQL = "INSERT INTO z_tag_0401PV8_2 (tag_value, created) values(" & tag_0401PV8_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV8_2 = Nothing
Dim tag_0401PV8_3
tag_0401PV8_3 = HMIRuntime.Tags("0401PV8_3").Read
strSQL = "INSERT INTO z_tag_0401PV8_3 (tag_value, created) values(" & tag_0401PV8_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV8_3 = Nothing
Dim tag_0401PV8_4
tag_0401PV8_4 = HMIRuntime.Tags("0401PV8_4").Read
strSQL = "INSERT INTO z_tag_0401PV8_4 (tag_value, created) values(" & tag_0401PV8_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV8_4 = Nothing
Dim tag_0401PV8_5
tag_0401PV8_5 = HMIRuntime.Tags("0401PV8_5").Read
strSQL = "INSERT INTO z_tag_0401PV8_5 (tag_value, created) values(" & tag_0401PV8_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV8_5 = Nothing
Dim tag_0401PV09
tag_0401PV09 = HMIRuntime.Tags("0401PV09").Read
strSQL = "INSERT INTO z_tag_0401PV09 (tag_value, created) values(" & tag_0401PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV09 = Nothing
Dim tag_0401PV9_1
tag_0401PV9_1 = HMIRuntime.Tags("0401PV9_1").Read
strSQL = "INSERT INTO z_tag_0401PV9_1 (tag_value, created) values(" & tag_0401PV9_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV9_1 = Nothing
Dim tag_0401PV9_2
tag_0401PV9_2 = HMIRuntime.Tags("0401PV9_2").Read
strSQL = "INSERT INTO z_tag_0401PV9_2 (tag_value, created) values(" & tag_0401PV9_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV9_2 = Nothing
Dim tag_0401PV9_3
tag_0401PV9_3 = HMIRuntime.Tags("0401PV9_3").Read
strSQL = "INSERT INTO z_tag_0401PV9_3 (tag_value, created) values(" & tag_0401PV9_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV9_3 = Nothing
Dim tag_0401PV9_4
tag_0401PV9_4 = HMIRuntime.Tags("0401PV9_4").Read
strSQL = "INSERT INTO z_tag_0401PV9_4 (tag_value, created) values(" & tag_0401PV9_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV9_4 = Nothing
Dim tag_0401PV9_5
tag_0401PV9_5 = HMIRuntime.Tags("0401PV9_5").Read
strSQL = "INSERT INTO z_tag_0401PV9_5 (tag_value, created) values(" & tag_0401PV9_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV9_5 = Nothing
Dim tag_0401PV10
tag_0401PV10 = HMIRuntime.Tags("0401PV10").Read
strSQL = "INSERT INTO z_tag_0401PV10 (tag_value, created) values(" & tag_0401PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10 = Nothing
Dim tag_0401PV10_1
tag_0401PV10_1 = HMIRuntime.Tags("0401PV10_1").Read
strSQL = "INSERT INTO z_tag_0401PV10_1 (tag_value, created) values(" & tag_0401PV10_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10_1 = Nothing
Dim tag_0401PV10_2
tag_0401PV10_2 = HMIRuntime.Tags("0401PV10_2").Read
strSQL = "INSERT INTO z_tag_0401PV10_2 (tag_value, created) values(" & tag_0401PV10_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10_2 = Nothing
Dim tag_0401PV10_3
tag_0401PV10_3 = HMIRuntime.Tags("0401PV10_3").Read
strSQL = "INSERT INTO z_tag_0401PV10_3 (tag_value, created) values(" & tag_0401PV10_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10_3 = Nothing
Dim tag_0401PV10_4
tag_0401PV10_4 = HMIRuntime.Tags("0401PV10_4").Read
strSQL = "INSERT INTO z_tag_0401PV10_4 (tag_value, created) values(" & tag_0401PV10_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10_4 = Nothing
Dim tag_0401PV10_5
tag_0401PV10_5 = HMIRuntime.Tags("0401PV10_5").Read
strSQL = "INSERT INTO z_tag_0401PV10_5 (tag_value, created) values(" & tag_0401PV10_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV10_5 = Nothing
Dim tag_0401PV11
tag_0401PV11 = HMIRuntime.Tags("0401PV11").Read
strSQL = "INSERT INTO z_tag_0401PV11 (tag_value, created) values(" & tag_0401PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11 = Nothing
Dim tag_0401PV11_1
tag_0401PV11_1 = HMIRuntime.Tags("0401PV11_1").Read
strSQL = "INSERT INTO z_tag_0401PV11_1 (tag_value, created) values(" & tag_0401PV11_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11_1 = Nothing
Dim tag_0401PV11_2
tag_0401PV11_2 = HMIRuntime.Tags("0401PV11_2").Read
strSQL = "INSERT INTO z_tag_0401PV11_2 (tag_value, created) values(" & tag_0401PV11_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11_2 = Nothing
Dim tag_0401PV11_3
tag_0401PV11_3 = HMIRuntime.Tags("0401PV11_3").Read
strSQL = "INSERT INTO z_tag_0401PV11_3 (tag_value, created) values(" & tag_0401PV11_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11_3 = Nothing
Dim tag_0401PV11_4
tag_0401PV11_4 = HMIRuntime.Tags("0401PV11_4").Read
strSQL = "INSERT INTO z_tag_0401PV11_4 (tag_value, created) values(" & tag_0401PV11_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11_4 = Nothing
Dim tag_0401PV11_5
tag_0401PV11_5 = HMIRuntime.Tags("0401PV11_5").Read
strSQL = "INSERT INTO z_tag_0401PV11_5 (tag_value, created) values(" & tag_0401PV11_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV11_5 = Nothing
Dim tag_0401PV12_1
tag_0401PV12_1 = HMIRuntime.Tags("0401PV12_1").Read
strSQL = "INSERT INTO z_tag_0401PV12_1 (tag_value, created) values(" & tag_0401PV12_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV12_1 = Nothing
Dim tag_0401PV12_2
tag_0401PV12_2 = HMIRuntime.Tags("0401PV12_2").Read
strSQL = "INSERT INTO z_tag_0401PV12_2 (tag_value, created) values(" & tag_0401PV12_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV12_2 = Nothing
Dim tag_0401PV12_3
tag_0401PV12_3 = HMIRuntime.Tags("0401PV12_3").Read
strSQL = "INSERT INTO z_tag_0401PV12_3 (tag_value, created) values(" & tag_0401PV12_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV12_3 = Nothing
Dim tag_0401PV12_4
tag_0401PV12_4 = HMIRuntime.Tags("0401PV12_4").Read
strSQL = "INSERT INTO z_tag_0401PV12_4 (tag_value, created) values(" & tag_0401PV12_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV12_4 = Nothing
Dim tag_0401PV12_5
tag_0401PV12_5 = HMIRuntime.Tags("0401PV12_5").Read
strSQL = "INSERT INTO z_tag_0401PV12_5 (tag_value, created) values(" & tag_0401PV12_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV12_5 = Nothing
Dim tag_0401PV13_1
tag_0401PV13_1 = HMIRuntime.Tags("0401PV13_1").Read
strSQL = "INSERT INTO z_tag_0401PV13_1 (tag_value, created) values(" & tag_0401PV13_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV13_1 = Nothing
Dim tag_0401PV13_2
tag_0401PV13_2 = HMIRuntime.Tags("0401PV13_2").Read
strSQL = "INSERT INTO z_tag_0401PV13_2 (tag_value, created) values(" & tag_0401PV13_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV13_2 = Nothing
Dim tag_0401PV13_3
tag_0401PV13_3 = HMIRuntime.Tags("0401PV13_3").Read
strSQL = "INSERT INTO z_tag_0401PV13_3 (tag_value, created) values(" & tag_0401PV13_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV13_3 = Nothing
Dim tag_0401PV13_4
tag_0401PV13_4 = HMIRuntime.Tags("0401PV13_4").Read
strSQL = "INSERT INTO z_tag_0401PV13_4 (tag_value, created) values(" & tag_0401PV13_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV13_4 = Nothing
Dim tag_0401PV13_5
tag_0401PV13_5 = HMIRuntime.Tags("0401PV13_5").Read
strSQL = "INSERT INTO z_tag_0401PV13_5 (tag_value, created) values(" & tag_0401PV13_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV13_5 = Nothing
Dim tag_0401PV14_1
tag_0401PV14_1 = HMIRuntime.Tags("0401PV14_1").Read
strSQL = "INSERT INTO z_tag_0401PV14_1 (tag_value, created) values(" & tag_0401PV14_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV14_1 = Nothing
Dim tag_0401PV14_2
tag_0401PV14_2 = HMIRuntime.Tags("0401PV14_2").Read
strSQL = "INSERT INTO z_tag_0401PV14_2 (tag_value, created) values(" & tag_0401PV14_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV14_2 = Nothing
Dim tag_0401PV14_3
tag_0401PV14_3 = HMIRuntime.Tags("0401PV14_3").Read
strSQL = "INSERT INTO z_tag_0401PV14_3 (tag_value, created) values(" & tag_0401PV14_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV14_3 = Nothing
Dim tag_0401PV14_4
tag_0401PV14_4 = HMIRuntime.Tags("0401PV14_4").Read
strSQL = "INSERT INTO z_tag_0401PV14_4 (tag_value, created) values(" & tag_0401PV14_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV14_4 = Nothing
Dim tag_0401PV14_5
tag_0401PV14_5 = HMIRuntime.Tags("0401PV14_5").Read
strSQL = "INSERT INTO z_tag_0401PV14_5 (tag_value, created) values(" & tag_0401PV14_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV14_5 = Nothing
Dim tag_0401PV15_1
tag_0401PV15_1 = HMIRuntime.Tags("0401PV15_1").Read
strSQL = "INSERT INTO z_tag_0401PV15_1 (tag_value, created) values(" & tag_0401PV15_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV15_1 = Nothing
Dim tag_0401PV15_2
tag_0401PV15_2 = HMIRuntime.Tags("0401PV15_2").Read
strSQL = "INSERT INTO z_tag_0401PV15_2 (tag_value, created) values(" & tag_0401PV15_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV15_2 = Nothing
Dim tag_0401PV15_3
tag_0401PV15_3 = HMIRuntime.Tags("0401PV15_3").Read
strSQL = "INSERT INTO z_tag_0401PV15_3 (tag_value, created) values(" & tag_0401PV15_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV15_3 = Nothing
Dim tag_0401PV15_4
tag_0401PV15_4 = HMIRuntime.Tags("0401PV15_4").Read
strSQL = "INSERT INTO z_tag_0401PV15_4 (tag_value, created) values(" & tag_0401PV15_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV15_4 = Nothing
Dim tag_0401PV15_5
tag_0401PV15_5 = HMIRuntime.Tags("0401PV15_5").Read
strSQL = "INSERT INTO z_tag_0401PV15_5 (tag_value, created) values(" & tag_0401PV15_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV15_5 = Nothing
Dim tag_0401PV16_1
tag_0401PV16_1 = HMIRuntime.Tags("0401PV16_1").Read
strSQL = "INSERT INTO z_tag_0401PV16_1 (tag_value, created) values(" & tag_0401PV16_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV16_1 = Nothing
Dim tag_0401PV16_2
tag_0401PV16_2 = HMIRuntime.Tags("0401PV16_2").Read
strSQL = "INSERT INTO z_tag_0401PV16_2 (tag_value, created) values(" & tag_0401PV16_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV16_2 = Nothing
Dim tag_0401PV16_3
tag_0401PV16_3 = HMIRuntime.Tags("0401PV16_3").Read
strSQL = "INSERT INTO z_tag_0401PV16_3 (tag_value, created) values(" & tag_0401PV16_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV16_3 = Nothing
Dim tag_0401PV16_4
tag_0401PV16_4 = HMIRuntime.Tags("0401PV16_4").Read
strSQL = "INSERT INTO z_tag_0401PV16_4 (tag_value, created) values(" & tag_0401PV16_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV16_4 = Nothing
Dim tag_0401PV16_5
tag_0401PV16_5 = HMIRuntime.Tags("0401PV16_5").Read
strSQL = "INSERT INTO z_tag_0401PV16_5 (tag_value, created) values(" & tag_0401PV16_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV16_5 = Nothing
Dim tag_0401PV17_1
tag_0401PV17_1 = HMIRuntime.Tags("0401PV17_1").Read
strSQL = "INSERT INTO z_tag_0401PV17_1 (tag_value, created) values(" & tag_0401PV17_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV17_1 = Nothing
Dim tag_0401PV17_2
tag_0401PV17_2 = HMIRuntime.Tags("0401PV17_2").Read
strSQL = "INSERT INTO z_tag_0401PV17_2 (tag_value, created) values(" & tag_0401PV17_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV17_2 = Nothing
Dim tag_0401PV17_3
tag_0401PV17_3 = HMIRuntime.Tags("0401PV17_3").Read
strSQL = "INSERT INTO z_tag_0401PV17_3 (tag_value, created) values(" & tag_0401PV17_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV17_3 = Nothing
Dim tag_0401PV17_4
tag_0401PV17_4 = HMIRuntime.Tags("0401PV17_4").Read
strSQL = "INSERT INTO z_tag_0401PV17_4 (tag_value, created) values(" & tag_0401PV17_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV17_4 = Nothing
Dim tag_0401PV17_5
tag_0401PV17_5 = HMIRuntime.Tags("0401PV17_5").Read
strSQL = "INSERT INTO z_tag_0401PV17_5 (tag_value, created) values(" & tag_0401PV17_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV17_5 = Nothing
Dim tag_0401PV18_1
tag_0401PV18_1 = HMIRuntime.Tags("0401PV18_1").Read
strSQL = "INSERT INTO z_tag_0401PV18_1 (tag_value, created) values(" & tag_0401PV18_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV18_1 = Nothing
Dim tag_0401PV18_2
tag_0401PV18_2 = HMIRuntime.Tags("0401PV18_2").Read
strSQL = "INSERT INTO z_tag_0401PV18_2 (tag_value, created) values(" & tag_0401PV18_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV18_2 = Nothing
Dim tag_0401PV18_3
tag_0401PV18_3 = HMIRuntime.Tags("0401PV18_3").Read
strSQL = "INSERT INTO z_tag_0401PV18_3 (tag_value, created) values(" & tag_0401PV18_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV18_3 = Nothing
Dim tag_0401PV18_4
tag_0401PV18_4 = HMIRuntime.Tags("0401PV18_4").Read
strSQL = "INSERT INTO z_tag_0401PV18_4 (tag_value, created) values(" & tag_0401PV18_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV18_4 = Nothing
Dim tag_0401PV18_5
tag_0401PV18_5 = HMIRuntime.Tags("0401PV18_5").Read
strSQL = "INSERT INTO z_tag_0401PV18_5 (tag_value, created) values(" & tag_0401PV18_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV18_5 = Nothing
Dim tag_0401PV19_1
tag_0401PV19_1 = HMIRuntime.Tags("0401PV19_1").Read
strSQL = "INSERT INTO z_tag_0401PV19_1 (tag_value, created) values(" & tag_0401PV19_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV19_1 = Nothing
Dim tag_0401PV19_2
tag_0401PV19_2 = HMIRuntime.Tags("0401PV19_2").Read
strSQL = "INSERT INTO z_tag_0401PV19_2 (tag_value, created) values(" & tag_0401PV19_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV19_2 = Nothing
Dim tag_0401PV19_3
tag_0401PV19_3 = HMIRuntime.Tags("0401PV19_3").Read
strSQL = "INSERT INTO z_tag_0401PV19_3 (tag_value, created) values(" & tag_0401PV19_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV19_3 = Nothing
Dim tag_0401PV19_4
tag_0401PV19_4 = HMIRuntime.Tags("0401PV19_4").Read
strSQL = "INSERT INTO z_tag_0401PV19_4 (tag_value, created) values(" & tag_0401PV19_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV19_4 = Nothing
Dim tag_0401PV19_5
tag_0401PV19_5 = HMIRuntime.Tags("0401PV19_5").Read
strSQL = "INSERT INTO z_tag_0401PV19_5 (tag_value, created) values(" & tag_0401PV19_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV19_5 = Nothing
Dim tag_0401PV20_1
tag_0401PV20_1 = HMIRuntime.Tags("0401PV20_1").Read
strSQL = "INSERT INTO z_tag_0401PV20_1 (tag_value, created) values(" & tag_0401PV20_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV20_1 = Nothing
Dim tag_0401PV20_2
tag_0401PV20_2 = HMIRuntime.Tags("0401PV20_2").Read
strSQL = "INSERT INTO z_tag_0401PV20_2 (tag_value, created) values(" & tag_0401PV20_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV20_2 = Nothing
Dim tag_0401PV20_3
tag_0401PV20_3 = HMIRuntime.Tags("0401PV20_3").Read
strSQL = "INSERT INTO z_tag_0401PV20_3 (tag_value, created) values(" & tag_0401PV20_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV20_3 = Nothing
Dim tag_0401PV20_4
tag_0401PV20_4 = HMIRuntime.Tags("0401PV20_4").Read
strSQL = "INSERT INTO z_tag_0401PV20_4 (tag_value, created) values(" & tag_0401PV20_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV20_4 = Nothing
Dim tag_0401PV20_5
tag_0401PV20_5 = HMIRuntime.Tags("0401PV20_5").Read
strSQL = "INSERT INTO z_tag_0401PV20_5 (tag_value, created) values(" & tag_0401PV20_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV20_5 = Nothing
Dim tag_0401PV21_1
tag_0401PV21_1 = HMIRuntime.Tags("0401PV21_1").Read
strSQL = "INSERT INTO z_tag_0401PV21_1 (tag_value, created) values(" & tag_0401PV21_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV21_1 = Nothing
Dim tag_0401PV21_2
tag_0401PV21_2 = HMIRuntime.Tags("0401PV21_2").Read
strSQL = "INSERT INTO z_tag_0401PV21_2 (tag_value, created) values(" & tag_0401PV21_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV21_2 = Nothing
Dim tag_0401PV21_3
tag_0401PV21_3 = HMIRuntime.Tags("0401PV21_3").Read
strSQL = "INSERT INTO z_tag_0401PV21_3 (tag_value, created) values(" & tag_0401PV21_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV21_3 = Nothing
Dim tag_0401PV21_4
tag_0401PV21_4 = HMIRuntime.Tags("0401PV21_4").Read
strSQL = "INSERT INTO z_tag_0401PV21_4 (tag_value, created) values(" & tag_0401PV21_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV21_4 = Nothing
Dim tag_0401PV21_5
tag_0401PV21_5 = HMIRuntime.Tags("0401PV21_5").Read
strSQL = "INSERT INTO z_tag_0401PV21_5 (tag_value, created) values(" & tag_0401PV21_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV21_5 = Nothing
Dim tag_0401PV22_1
tag_0401PV22_1 = HMIRuntime.Tags("0401PV22_1").Read
strSQL = "INSERT INTO z_tag_0401PV22_1 (tag_value, created) values(" & tag_0401PV22_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV22_1 = Nothing
Dim tag_0401PV22_2
tag_0401PV22_2 = HMIRuntime.Tags("0401PV22_2").Read
strSQL = "INSERT INTO z_tag_0401PV22_2 (tag_value, created) values(" & tag_0401PV22_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV22_2 = Nothing
Dim tag_0401PV22_3
tag_0401PV22_3 = HMIRuntime.Tags("0401PV22_3").Read
strSQL = "INSERT INTO z_tag_0401PV22_3 (tag_value, created) values(" & tag_0401PV22_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV22_3 = Nothing
Dim tag_0401PV22_4
tag_0401PV22_4 = HMIRuntime.Tags("0401PV22_4").Read
strSQL = "INSERT INTO z_tag_0401PV22_4 (tag_value, created) values(" & tag_0401PV22_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV22_4 = Nothing
Dim tag_0401PV22_5
tag_0401PV22_5 = HMIRuntime.Tags("0401PV22_5").Read
strSQL = "INSERT INTO z_tag_0401PV22_5 (tag_value, created) values(" & tag_0401PV22_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV22_5 = Nothing
Dim tag_0401PV23_1
tag_0401PV23_1 = HMIRuntime.Tags("0401PV23_1").Read
strSQL = "INSERT INTO z_tag_0401PV23_1 (tag_value, created) values(" & tag_0401PV23_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV23_1 = Nothing
Dim tag_0401PV23_2
tag_0401PV23_2 = HMIRuntime.Tags("0401PV23_2").Read
strSQL = "INSERT INTO z_tag_0401PV23_2 (tag_value, created) values(" & tag_0401PV23_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV23_2 = Nothing
Dim tag_0401PV23_3
tag_0401PV23_3 = HMIRuntime.Tags("0401PV23_3").Read
strSQL = "INSERT INTO z_tag_0401PV23_3 (tag_value, created) values(" & tag_0401PV23_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV23_3 = Nothing
Dim tag_0401PV23_4
tag_0401PV23_4 = HMIRuntime.Tags("0401PV23_4").Read
strSQL = "INSERT INTO z_tag_0401PV23_4 (tag_value, created) values(" & tag_0401PV23_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV23_4 = Nothing
Dim tag_0401PV23_5
tag_0401PV23_5 = HMIRuntime.Tags("0401PV23_5").Read
strSQL = "INSERT INTO z_tag_0401PV23_5 (tag_value, created) values(" & tag_0401PV23_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV23_5 = Nothing
Dim tag_0401PV24_1
tag_0401PV24_1 = HMIRuntime.Tags("0401PV24_1").Read
strSQL = "INSERT INTO z_tag_0401PV24_1 (tag_value, created) values(" & tag_0401PV24_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV24_1 = Nothing
Dim tag_0401PV24_2
tag_0401PV24_2 = HMIRuntime.Tags("0401PV24_2").Read
strSQL = "INSERT INTO z_tag_0401PV24_2 (tag_value, created) values(" & tag_0401PV24_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV24_2 = Nothing
Dim tag_0401PV24_3
tag_0401PV24_3 = HMIRuntime.Tags("0401PV24_3").Read
strSQL = "INSERT INTO z_tag_0401PV24_3 (tag_value, created) values(" & tag_0401PV24_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV24_3 = Nothing
Dim tag_0401PV24_4
tag_0401PV24_4 = HMIRuntime.Tags("0401PV24_4").Read
strSQL = "INSERT INTO z_tag_0401PV24_4 (tag_value, created) values(" & tag_0401PV24_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV24_4 = Nothing
Dim tag_0401PV24_5
tag_0401PV24_5 = HMIRuntime.Tags("0401PV24_5").Read
strSQL = "INSERT INTO z_tag_0401PV24_5 (tag_value, created) values(" & tag_0401PV24_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV24_5 = Nothing
Dim tag_0401PV25_1
tag_0401PV25_1 = HMIRuntime.Tags("0401PV25_1").Read
strSQL = "INSERT INTO z_tag_0401PV25_1 (tag_value, created) values(" & tag_0401PV25_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV25_1 = Nothing
Dim tag_0401PV25_2
tag_0401PV25_2 = HMIRuntime.Tags("0401PV25_2").Read
strSQL = "INSERT INTO z_tag_0401PV25_2 (tag_value, created) values(" & tag_0401PV25_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV25_2 = Nothing
Dim tag_0401PV25_3
tag_0401PV25_3 = HMIRuntime.Tags("0401PV25_3").Read
strSQL = "INSERT INTO z_tag_0401PV25_3 (tag_value, created) values(" & tag_0401PV25_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV25_3 = Nothing
Dim tag_0401PV25_4
tag_0401PV25_4 = HMIRuntime.Tags("0401PV25_4").Read
strSQL = "INSERT INTO z_tag_0401PV25_4 (tag_value, created) values(" & tag_0401PV25_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV25_4 = Nothing
Dim tag_0401PV25_5
tag_0401PV25_5 = HMIRuntime.Tags("0401PV25_5").Read
strSQL = "INSERT INTO z_tag_0401PV25_5 (tag_value, created) values(" & tag_0401PV25_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV25_5 = Nothing
Dim tag_0401PV26_1
tag_0401PV26_1 = HMIRuntime.Tags("0401PV26_1").Read
strSQL = "INSERT INTO z_tag_0401PV26_1 (tag_value, created) values(" & tag_0401PV26_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV26_1 = Nothing
Dim tag_0401PV26_2
tag_0401PV26_2 = HMIRuntime.Tags("0401PV26_2").Read
strSQL = "INSERT INTO z_tag_0401PV26_2 (tag_value, created) values(" & tag_0401PV26_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV26_2 = Nothing
Dim tag_0401PV26_3
tag_0401PV26_3 = HMIRuntime.Tags("0401PV26_3").Read
strSQL = "INSERT INTO z_tag_0401PV26_3 (tag_value, created) values(" & tag_0401PV26_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV26_3 = Nothing
Dim tag_0401PV26_4
tag_0401PV26_4 = HMIRuntime.Tags("0401PV26_4").Read
strSQL = "INSERT INTO z_tag_0401PV26_4 (tag_value, created) values(" & tag_0401PV26_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV26_4 = Nothing
Dim tag_0401PV26_5
tag_0401PV26_5 = HMIRuntime.Tags("0401PV26_5").Read
strSQL = "INSERT INTO z_tag_0401PV26_5 (tag_value, created) values(" & tag_0401PV26_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV26_5 = Nothing
Dim tag_0401PV27_1
tag_0401PV27_1 = HMIRuntime.Tags("0401PV27_1").Read
strSQL = "INSERT INTO z_tag_0401PV27_1 (tag_value, created) values(" & tag_0401PV27_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV27_1 = Nothing
Dim tag_0401PV27_2
tag_0401PV27_2 = HMIRuntime.Tags("0401PV27_2").Read
strSQL = "INSERT INTO z_tag_0401PV27_2 (tag_value, created) values(" & tag_0401PV27_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV27_2 = Nothing
Dim tag_0401PV27_3
tag_0401PV27_3 = HMIRuntime.Tags("0401PV27_3").Read
strSQL = "INSERT INTO z_tag_0401PV27_3 (tag_value, created) values(" & tag_0401PV27_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV27_3 = Nothing
Dim tag_0401PV27_4
tag_0401PV27_4 = HMIRuntime.Tags("0401PV27_4").Read
strSQL = "INSERT INTO z_tag_0401PV27_4 (tag_value, created) values(" & tag_0401PV27_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV27_4 = Nothing
Dim tag_0401PV27_5
tag_0401PV27_5 = HMIRuntime.Tags("0401PV27_5").Read
strSQL = "INSERT INTO z_tag_0401PV27_5 (tag_value, created) values(" & tag_0401PV27_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV27_5 = Nothing
Dim tag_0401PV28_1
tag_0401PV28_1 = HMIRuntime.Tags("0401PV28_1").Read
strSQL = "INSERT INTO z_tag_0401PV28_1 (tag_value, created) values(" & tag_0401PV28_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV28_1 = Nothing
Dim tag_0401PV28_2
tag_0401PV28_2 = HMIRuntime.Tags("0401PV28_2").Read
strSQL = "INSERT INTO z_tag_0401PV28_2 (tag_value, created) values(" & tag_0401PV28_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV28_2 = Nothing
Dim tag_0401PV28_3
tag_0401PV28_3 = HMIRuntime.Tags("0401PV28_3").Read
strSQL = "INSERT INTO z_tag_0401PV28_3 (tag_value, created) values(" & tag_0401PV28_3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV28_3 = Nothing
Dim tag_0401PV28_4
tag_0401PV28_4 = HMIRuntime.Tags("0401PV28_4").Read
strSQL = "INSERT INTO z_tag_0401PV28_4 (tag_value, created) values(" & tag_0401PV28_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV28_4 = Nothing
Dim tag_0401PV28_5
tag_0401PV28_5 = HMIRuntime.Tags("0401PV28_5").Read
strSQL = "INSERT INTO z_tag_0401PV28_5 (tag_value, created) values(" & tag_0401PV28_5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV28_5 = Nothing
Dim tag_0401PV31_4
tag_0401PV31_4 = HMIRuntime.Tags("0401PV31_4").Read
strSQL = "INSERT INTO z_tag_0401PV31_4 (tag_value, created) values(" & tag_0401PV31_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV31_4 = Nothing
Dim tag_0401PV32_4
tag_0401PV32_4 = HMIRuntime.Tags("0401PV32_4").Read
strSQL = "INSERT INTO z_tag_0401PV32_4 (tag_value, created) values(" & tag_0401PV32_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV32_4 = Nothing
Dim tag_0401PV33_4
tag_0401PV33_4 = HMIRuntime.Tags("0401PV33_4").Read
strSQL = "INSERT INTO z_tag_0401PV33_4 (tag_value, created) values(" & tag_0401PV33_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV33_4 = Nothing
Dim tag_0401PV34_4
tag_0401PV34_4 = HMIRuntime.Tags("0401PV34_4").Read
strSQL = "INSERT INTO z_tag_0401PV34_4 (tag_value, created) values(" & tag_0401PV34_4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0401PV34_4 = Nothing
Dim tag_0404_PET01
tag_0404_PET01 = HMIRuntime.Tags("0404_PET01").Read
strSQL = "INSERT INTO z_tag_0404_PET01 (tag_value, created) values(" & tag_0404_PET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0404_PET01 = Nothing
Dim tag_0601_POT_01
tag_0601_POT_01 = HMIRuntime.Tags("0601_POT_01").Read
strSQL = "INSERT INTO z_tag_0601_POT_01 (tag_value, created) values(" & tag_0601_POT_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_POT_01 = Nothing
Dim tag_0601_POT_02
tag_0601_POT_02 = HMIRuntime.Tags("0601_POT_02").Read
strSQL = "INSERT INTO z_tag_0601_POT_02 (tag_value, created) values(" & tag_0601_POT_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_POT_02 = Nothing
Dim tag_0601_POT_03
tag_0601_POT_03 = HMIRuntime.Tags("0601_POT_03").Read
strSQL = "INSERT INTO z_tag_0601_POT_03 (tag_value, created) values(" & tag_0601_POT_03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_POT_03 = Nothing
Dim tag_0601_POT_04
tag_0601_POT_04 = HMIRuntime.Tags("0601_POT_04").Read
strSQL = "INSERT INTO z_tag_0601_POT_04 (tag_value, created) values(" & tag_0601_POT_04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_POT_04 = Nothing
Dim tag_0601_TET_01
tag_0601_TET_01 = HMIRuntime.Tags("0601_TET_01").Read
strSQL = "INSERT INTO z_tag_0601_TET_01 (tag_value, created) values(" & tag_0601_TET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_TET_01 = Nothing
Dim tag_0601_TET_02
tag_0601_TET_02 = HMIRuntime.Tags("0601_TET_02").Read
strSQL = "INSERT INTO z_tag_0601_TET_02 (tag_value, created) values(" & tag_0601_TET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_TET_02 = Nothing
Dim tag_0601_TET_03
tag_0601_TET_03 = HMIRuntime.Tags("0601_TET_03").Read
strSQL = "INSERT INTO z_tag_0601_TET_03 (tag_value, created) values(" & tag_0601_TET_03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_TET_03 = Nothing
Dim tag_0601_TET_04
tag_0601_TET_04 = HMIRuntime.Tags("0601_TET_04").Read
strSQL = "INSERT INTO z_tag_0601_TET_04 (tag_value, created) values(" & tag_0601_TET_04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601_TET_04 = Nothing
Dim tag_0601M01
tag_0601M01 = HMIRuntime.Tags("0601M01").Read
strSQL = "INSERT INTO z_tag_0601M01 (tag_value, created) values(" & tag_0601M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601M01 = Nothing
Dim tag_0601M05
tag_0601M05 = HMIRuntime.Tags("0601M05").Read
strSQL = "INSERT INTO z_tag_0601M05 (tag_value, created) values(" & tag_0601M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601M05 = Nothing
Dim tag_0601M06
tag_0601M06 = HMIRuntime.Tags("0601M06").Read
strSQL = "INSERT INTO z_tag_0601M06 (tag_value, created) values(" & tag_0601M06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601M06 = Nothing
Dim tag_0601PV01
tag_0601PV01 = HMIRuntime.Tags("0601PV01").Read
strSQL = "INSERT INTO z_tag_0601PV01 (tag_value, created) values(" & tag_0601PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV01 = Nothing
Dim tag_0601PV01N
tag_0601PV01N = HMIRuntime.Tags("0601PV01N").Read
strSQL = "INSERT INTO z_tag_0601PV01N (tag_value, created) values(" & tag_0601PV01N & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV01N = Nothing
Dim tag_0601PV02
tag_0601PV02 = HMIRuntime.Tags("0601PV02").Read
strSQL = "INSERT INTO z_tag_0601PV02 (tag_value, created) values(" & tag_0601PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV02 = Nothing
Dim tag_0601PV03
tag_0601PV03 = HMIRuntime.Tags("0601PV03").Read
strSQL = "INSERT INTO z_tag_0601PV03 (tag_value, created) values(" & tag_0601PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV03 = Nothing
Dim tag_0601PV04
tag_0601PV04 = HMIRuntime.Tags("0601PV04").Read
strSQL = "INSERT INTO z_tag_0601PV04 (tag_value, created) values(" & tag_0601PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV04 = Nothing
Dim tag_0601PV05
tag_0601PV05 = HMIRuntime.Tags("0601PV05").Read
strSQL = "INSERT INTO z_tag_0601PV05 (tag_value, created) values(" & tag_0601PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV05 = Nothing
Dim tag_0601PV06
tag_0601PV06 = HMIRuntime.Tags("0601PV06").Read
strSQL = "INSERT INTO z_tag_0601PV06 (tag_value, created) values(" & tag_0601PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV06 = Nothing
Dim tag_0601PV07
tag_0601PV07 = HMIRuntime.Tags("0601PV07").Read
strSQL = "INSERT INTO z_tag_0601PV07 (tag_value, created) values(" & tag_0601PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV07 = Nothing
Dim tag_0601PV08
tag_0601PV08 = HMIRuntime.Tags("0601PV08").Read
strSQL = "INSERT INTO z_tag_0601PV08 (tag_value, created) values(" & tag_0601PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV08 = Nothing
Dim tag_0601PV09
tag_0601PV09 = HMIRuntime.Tags("0601PV09").Read
strSQL = "INSERT INTO z_tag_0601PV09 (tag_value, created) values(" & tag_0601PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV09 = Nothing
Dim tag_0601PV10
tag_0601PV10 = HMIRuntime.Tags("0601PV10").Read
strSQL = "INSERT INTO z_tag_0601PV10 (tag_value, created) values(" & tag_0601PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV10 = Nothing
Dim tag_0601PV11
tag_0601PV11 = HMIRuntime.Tags("0601PV11").Read
strSQL = "INSERT INTO z_tag_0601PV11 (tag_value, created) values(" & tag_0601PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV11 = Nothing
Dim tag_0601PV12
tag_0601PV12 = HMIRuntime.Tags("0601PV12").Read
strSQL = "INSERT INTO z_tag_0601PV12 (tag_value, created) values(" & tag_0601PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV12 = Nothing
Dim tag_0601PV13
tag_0601PV13 = HMIRuntime.Tags("0601PV13").Read
strSQL = "INSERT INTO z_tag_0601PV13 (tag_value, created) values(" & tag_0601PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV13 = Nothing
Dim tag_0601PV14
tag_0601PV14 = HMIRuntime.Tags("0601PV14").Read
strSQL = "INSERT INTO z_tag_0601PV14 (tag_value, created) values(" & tag_0601PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV14 = Nothing
Dim tag_0601PV15
tag_0601PV15 = HMIRuntime.Tags("0601PV15").Read
strSQL = "INSERT INTO z_tag_0601PV15 (tag_value, created) values(" & tag_0601PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV15 = Nothing
Dim tag_0601PV16
tag_0601PV16 = HMIRuntime.Tags("0601PV16").Read
strSQL = "INSERT INTO z_tag_0601PV16 (tag_value, created) values(" & tag_0601PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV16 = Nothing
Dim tag_0601PV17
tag_0601PV17 = HMIRuntime.Tags("0601PV17").Read
strSQL = "INSERT INTO z_tag_0601PV17 (tag_value, created) values(" & tag_0601PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV17 = Nothing
Dim tag_0601PV18
tag_0601PV18 = HMIRuntime.Tags("0601PV18").Read
strSQL = "INSERT INTO z_tag_0601PV18 (tag_value, created) values(" & tag_0601PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV18 = Nothing
Dim tag_0601PV19
tag_0601PV19 = HMIRuntime.Tags("0601PV19").Read
strSQL = "INSERT INTO z_tag_0601PV19 (tag_value, created) values(" & tag_0601PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV19 = Nothing
Dim tag_0601PV20
tag_0601PV20 = HMIRuntime.Tags("0601PV20").Read
strSQL = "INSERT INTO z_tag_0601PV20 (tag_value, created) values(" & tag_0601PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV20 = Nothing
Dim tag_0601PV21
tag_0601PV21 = HMIRuntime.Tags("0601PV21").Read
strSQL = "INSERT INTO z_tag_0601PV21 (tag_value, created) values(" & tag_0601PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV21 = Nothing
Dim tag_0601PV22
tag_0601PV22 = HMIRuntime.Tags("0601PV22").Read
strSQL = "INSERT INTO z_tag_0601PV22 (tag_value, created) values(" & tag_0601PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV22 = Nothing
Dim tag_0601PV092
tag_0601PV092 = HMIRuntime.Tags("0601PV092").Read
strSQL = "INSERT INTO z_tag_0601PV092 (tag_value, created) values(" & tag_0601PV092 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0601PV092 = Nothing
Dim tag_0701_FQET_01
tag_0701_FQET_01 = HMIRuntime.Tags("0701_FQET_01").Read
strSQL = "INSERT INTO z_tag_0701_FQET_01 (tag_value, created) values(" & tag_0701_FQET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701_FQET_01 = Nothing
Dim tag_0701_FQET_02
tag_0701_FQET_02 = HMIRuntime.Tags("0701_FQET_02").Read
strSQL = "INSERT INTO z_tag_0701_FQET_02 (tag_value, created) values(" & tag_0701_FQET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701_FQET_02 = Nothing
Dim tag_0701_QET_01
tag_0701_QET_01 = HMIRuntime.Tags("0701_QET_01").Read
strSQL = "INSERT INTO z_tag_0701_QET_01 (tag_value, created) values(" & tag_0701_QET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701_QET_01 = Nothing
Dim tag_0701_QET_02
tag_0701_QET_02 = HMIRuntime.Tags("0701_QET_02").Read
strSQL = "INSERT INTO z_tag_0701_QET_02 (tag_value, created) values(" & tag_0701_QET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701_QET_02 = Nothing
Dim tag_0701_TET_01
tag_0701_TET_01 = HMIRuntime.Tags("0701_TET_01").Read
strSQL = "INSERT INTO z_tag_0701_TET_01 (tag_value, created) values(" & tag_0701_TET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701_TET_01 = Nothing
Dim tag_0701LET01
tag_0701LET01 = HMIRuntime.Tags("0701LET01").Read
strSQL = "INSERT INTO z_tag_0701LET01 (tag_value, created) values(" & tag_0701LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701LET01 = Nothing
Dim tag_0701LET02
tag_0701LET02 = HMIRuntime.Tags("0701LET02").Read
strSQL = "INSERT INTO z_tag_0701LET02 (tag_value, created) values(" & tag_0701LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701LET02 = Nothing
Dim tag_0701LET03
tag_0701LET03 = HMIRuntime.Tags("0701LET03").Read
strSQL = "INSERT INTO z_tag_0701LET03 (tag_value, created) values(" & tag_0701LET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701LET03 = Nothing
Dim tag_0701LET04
tag_0701LET04 = HMIRuntime.Tags("0701LET04").Read
strSQL = "INSERT INTO z_tag_0701LET04 (tag_value, created) values(" & tag_0701LET04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701LET04 = Nothing
Dim tag_0701M01
tag_0701M01 = HMIRuntime.Tags("0701M01").Read
strSQL = "INSERT INTO z_tag_0701M01 (tag_value, created) values(" & tag_0701M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701M01 = Nothing
Dim tag_0701M02
tag_0701M02 = HMIRuntime.Tags("0701M02").Read
strSQL = "INSERT INTO z_tag_0701M02 (tag_value, created) values(" & tag_0701M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701M02 = Nothing
Dim tag_0701M03
tag_0701M03 = HMIRuntime.Tags("0701M03").Read
strSQL = "INSERT INTO z_tag_0701M03 (tag_value, created) values(" & tag_0701M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701M03 = Nothing
Dim tag_0701M04
tag_0701M04 = HMIRuntime.Tags("0701M04").Read
strSQL = "INSERT INTO z_tag_0701M04 (tag_value, created) values(" & tag_0701M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701M04 = Nothing
Dim tag_0701M05
tag_0701M05 = HMIRuntime.Tags("0701M05").Read
strSQL = "INSERT INTO z_tag_0701M05 (tag_value, created) values(" & tag_0701M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701M05 = Nothing
Dim tag_0701PV01
tag_0701PV01 = HMIRuntime.Tags("0701PV01").Read
strSQL = "INSERT INTO z_tag_0701PV01 (tag_value, created) values(" & tag_0701PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV01 = Nothing
Dim tag_0701PV02
tag_0701PV02 = HMIRuntime.Tags("0701PV02").Read
strSQL = "INSERT INTO z_tag_0701PV02 (tag_value, created) values(" & tag_0701PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV02 = Nothing
Dim tag_0701PV03
tag_0701PV03 = HMIRuntime.Tags("0701PV03").Read
strSQL = "INSERT INTO z_tag_0701PV03 (tag_value, created) values(" & tag_0701PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV03 = Nothing
Dim tag_0701PV04
tag_0701PV04 = HMIRuntime.Tags("0701PV04").Read
strSQL = "INSERT INTO z_tag_0701PV04 (tag_value, created) values(" & tag_0701PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV04 = Nothing
Dim tag_0701PV05
tag_0701PV05 = HMIRuntime.Tags("0701PV05").Read
strSQL = "INSERT INTO z_tag_0701PV05 (tag_value, created) values(" & tag_0701PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV05 = Nothing
Dim tag_0701PV06
tag_0701PV06 = HMIRuntime.Tags("0701PV06").Read
strSQL = "INSERT INTO z_tag_0701PV06 (tag_value, created) values(" & tag_0701PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV06 = Nothing
Dim tag_0701PV07
tag_0701PV07 = HMIRuntime.Tags("0701PV07").Read
strSQL = "INSERT INTO z_tag_0701PV07 (tag_value, created) values(" & tag_0701PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV07 = Nothing
Dim tag_0701PV08
tag_0701PV08 = HMIRuntime.Tags("0701PV08").Read
strSQL = "INSERT INTO z_tag_0701PV08 (tag_value, created) values(" & tag_0701PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV08 = Nothing
Dim tag_0701PV09
tag_0701PV09 = HMIRuntime.Tags("0701PV09").Read
strSQL = "INSERT INTO z_tag_0701PV09 (tag_value, created) values(" & tag_0701PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV09 = Nothing
Dim tag_0701PV10
tag_0701PV10 = HMIRuntime.Tags("0701PV10").Read
strSQL = "INSERT INTO z_tag_0701PV10 (tag_value, created) values(" & tag_0701PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV10 = Nothing
Dim tag_0701PV11
tag_0701PV11 = HMIRuntime.Tags("0701PV11").Read
strSQL = "INSERT INTO z_tag_0701PV11 (tag_value, created) values(" & tag_0701PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV11 = Nothing
Dim tag_0701PV12
tag_0701PV12 = HMIRuntime.Tags("0701PV12").Read
strSQL = "INSERT INTO z_tag_0701PV12 (tag_value, created) values(" & tag_0701PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV12 = Nothing
Dim tag_0701PV13
tag_0701PV13 = HMIRuntime.Tags("0701PV13").Read
strSQL = "INSERT INTO z_tag_0701PV13 (tag_value, created) values(" & tag_0701PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV13 = Nothing
Dim tag_0701PV14
tag_0701PV14 = HMIRuntime.Tags("0701PV14").Read
strSQL = "INSERT INTO z_tag_0701PV14 (tag_value, created) values(" & tag_0701PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV14 = Nothing
Dim tag_0701PV15
tag_0701PV15 = HMIRuntime.Tags("0701PV15").Read
strSQL = "INSERT INTO z_tag_0701PV15 (tag_value, created) values(" & tag_0701PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV15 = Nothing
Dim tag_0701PV16
tag_0701PV16 = HMIRuntime.Tags("0701PV16").Read
strSQL = "INSERT INTO z_tag_0701PV16 (tag_value, created) values(" & tag_0701PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV16 = Nothing
Dim tag_0701PV17
tag_0701PV17 = HMIRuntime.Tags("0701PV17").Read
strSQL = "INSERT INTO z_tag_0701PV17 (tag_value, created) values(" & tag_0701PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV17 = Nothing
Dim tag_0701PV18
tag_0701PV18 = HMIRuntime.Tags("0701PV18").Read
strSQL = "INSERT INTO z_tag_0701PV18 (tag_value, created) values(" & tag_0701PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV18 = Nothing
Dim tag_0701PV19
tag_0701PV19 = HMIRuntime.Tags("0701PV19").Read
strSQL = "INSERT INTO z_tag_0701PV19 (tag_value, created) values(" & tag_0701PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV19 = Nothing
Dim tag_0701PV20
tag_0701PV20 = HMIRuntime.Tags("0701PV20").Read
strSQL = "INSERT INTO z_tag_0701PV20 (tag_value, created) values(" & tag_0701PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV20 = Nothing
Dim tag_0701PV21
tag_0701PV21 = HMIRuntime.Tags("0701PV21").Read
strSQL = "INSERT INTO z_tag_0701PV21 (tag_value, created) values(" & tag_0701PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV21 = Nothing
Dim tag_0701PV22
tag_0701PV22 = HMIRuntime.Tags("0701PV22").Read
strSQL = "INSERT INTO z_tag_0701PV22 (tag_value, created) values(" & tag_0701PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV22 = Nothing
Dim tag_0701PV23
tag_0701PV23 = HMIRuntime.Tags("0701PV23").Read
strSQL = "INSERT INTO z_tag_0701PV23 (tag_value, created) values(" & tag_0701PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV23 = Nothing
Dim tag_0701PV24
tag_0701PV24 = HMIRuntime.Tags("0701PV24").Read
strSQL = "INSERT INTO z_tag_0701PV24 (tag_value, created) values(" & tag_0701PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV24 = Nothing
Dim tag_0701PV25
tag_0701PV25 = HMIRuntime.Tags("0701PV25").Read
strSQL = "INSERT INTO z_tag_0701PV25 (tag_value, created) values(" & tag_0701PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV25 = Nothing
Dim tag_0701PV26
tag_0701PV26 = HMIRuntime.Tags("0701PV26").Read
strSQL = "INSERT INTO z_tag_0701PV26 (tag_value, created) values(" & tag_0701PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV26 = Nothing
Dim tag_0701PV27
tag_0701PV27 = HMIRuntime.Tags("0701PV27").Read
strSQL = "INSERT INTO z_tag_0701PV27 (tag_value, created) values(" & tag_0701PV27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV27 = Nothing
Dim tag_0701PV28
tag_0701PV28 = HMIRuntime.Tags("0701PV28").Read
strSQL = "INSERT INTO z_tag_0701PV28 (tag_value, created) values(" & tag_0701PV28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV28 = Nothing
Dim tag_0701PV29
tag_0701PV29 = HMIRuntime.Tags("0701PV29").Read
strSQL = "INSERT INTO z_tag_0701PV29 (tag_value, created) values(" & tag_0701PV29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV29 = Nothing
Dim tag_0701PV30
tag_0701PV30 = HMIRuntime.Tags("0701PV30").Read
strSQL = "INSERT INTO z_tag_0701PV30 (tag_value, created) values(" & tag_0701PV30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV30 = Nothing
Dim tag_0701PV31
tag_0701PV31 = HMIRuntime.Tags("0701PV31").Read
strSQL = "INSERT INTO z_tag_0701PV31 (tag_value, created) values(" & tag_0701PV31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV31 = Nothing
Dim tag_0701PV32
tag_0701PV32 = HMIRuntime.Tags("0701PV32").Read
strSQL = "INSERT INTO z_tag_0701PV32 (tag_value, created) values(" & tag_0701PV32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV32 = Nothing
Dim tag_0701PV33
tag_0701PV33 = HMIRuntime.Tags("0701PV33").Read
strSQL = "INSERT INTO z_tag_0701PV33 (tag_value, created) values(" & tag_0701PV33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV33 = Nothing
Dim tag_0701PV34
tag_0701PV34 = HMIRuntime.Tags("0701PV34").Read
strSQL = "INSERT INTO z_tag_0701PV34 (tag_value, created) values(" & tag_0701PV34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV34 = Nothing
Dim tag_0701PV35
tag_0701PV35 = HMIRuntime.Tags("0701PV35").Read
strSQL = "INSERT INTO z_tag_0701PV35 (tag_value, created) values(" & tag_0701PV35 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV35 = Nothing
Dim tag_0701PV36
tag_0701PV36 = HMIRuntime.Tags("0701PV36").Read
strSQL = "INSERT INTO z_tag_0701PV36 (tag_value, created) values(" & tag_0701PV36 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV36 = Nothing
Dim tag_0701PV37
tag_0701PV37 = HMIRuntime.Tags("0701PV37").Read
strSQL = "INSERT INTO z_tag_0701PV37 (tag_value, created) values(" & tag_0701PV37 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV37 = Nothing
Dim tag_0701PV38
tag_0701PV38 = HMIRuntime.Tags("0701PV38").Read
strSQL = "INSERT INTO z_tag_0701PV38 (tag_value, created) values(" & tag_0701PV38 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV38 = Nothing
Dim tag_0701PV39
tag_0701PV39 = HMIRuntime.Tags("0701PV39").Read
strSQL = "INSERT INTO z_tag_0701PV39 (tag_value, created) values(" & tag_0701PV39 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV39 = Nothing
Dim tag_0701PV40
tag_0701PV40 = HMIRuntime.Tags("0701PV40").Read
strSQL = "INSERT INTO z_tag_0701PV40 (tag_value, created) values(" & tag_0701PV40 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV40 = Nothing
Dim tag_0701PV41
tag_0701PV41 = HMIRuntime.Tags("0701PV41").Read
strSQL = "INSERT INTO z_tag_0701PV41 (tag_value, created) values(" & tag_0701PV41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV41 = Nothing
Dim tag_0701PV42
tag_0701PV42 = HMIRuntime.Tags("0701PV42").Read
strSQL = "INSERT INTO z_tag_0701PV42 (tag_value, created) values(" & tag_0701PV42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV42 = Nothing
Dim tag_0701PV43
tag_0701PV43 = HMIRuntime.Tags("0701PV43").Read
strSQL = "INSERT INTO z_tag_0701PV43 (tag_value, created) values(" & tag_0701PV43 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV43 = Nothing
Dim tag_0701PV44
tag_0701PV44 = HMIRuntime.Tags("0701PV44").Read
strSQL = "INSERT INTO z_tag_0701PV44 (tag_value, created) values(" & tag_0701PV44 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV44 = Nothing
Dim tag_0701PV45
tag_0701PV45 = HMIRuntime.Tags("0701PV45").Read
strSQL = "INSERT INTO z_tag_0701PV45 (tag_value, created) values(" & tag_0701PV45 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV45 = Nothing
Dim tag_0701PV46
tag_0701PV46 = HMIRuntime.Tags("0701PV46").Read
strSQL = "INSERT INTO z_tag_0701PV46 (tag_value, created) values(" & tag_0701PV46 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV46 = Nothing
Dim tag_0701PV47
tag_0701PV47 = HMIRuntime.Tags("0701PV47").Read
strSQL = "INSERT INTO z_tag_0701PV47 (tag_value, created) values(" & tag_0701PV47 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV47 = Nothing
Dim tag_0701PV48
tag_0701PV48 = HMIRuntime.Tags("0701PV48").Read
strSQL = "INSERT INTO z_tag_0701PV48 (tag_value, created) values(" & tag_0701PV48 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV48 = Nothing
Dim tag_0701PV49
tag_0701PV49 = HMIRuntime.Tags("0701PV49").Read
strSQL = "INSERT INTO z_tag_0701PV49 (tag_value, created) values(" & tag_0701PV49 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV49 = Nothing
Dim tag_0701PV50
tag_0701PV50 = HMIRuntime.Tags("0701PV50").Read
strSQL = "INSERT INTO z_tag_0701PV50 (tag_value, created) values(" & tag_0701PV50 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV50 = Nothing
Dim tag_0701PV51
tag_0701PV51 = HMIRuntime.Tags("0701PV51").Read
strSQL = "INSERT INTO z_tag_0701PV51 (tag_value, created) values(" & tag_0701PV51 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV51 = Nothing
Dim tag_0701PV52
tag_0701PV52 = HMIRuntime.Tags("0701PV52").Read
strSQL = "INSERT INTO z_tag_0701PV52 (tag_value, created) values(" & tag_0701PV52 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV52 = Nothing
Dim tag_0701PV54
tag_0701PV54 = HMIRuntime.Tags("0701PV54").Read
strSQL = "INSERT INTO z_tag_0701PV54 (tag_value, created) values(" & tag_0701PV54 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV54 = Nothing
Dim tag_0701PV55
tag_0701PV55 = HMIRuntime.Tags("0701PV55").Read
strSQL = "INSERT INTO z_tag_0701PV55 (tag_value, created) values(" & tag_0701PV55 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV55 = Nothing
Dim tag_0701PV56
tag_0701PV56 = HMIRuntime.Tags("0701PV56").Read
strSQL = "INSERT INTO z_tag_0701PV56 (tag_value, created) values(" & tag_0701PV56 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0701PV56 = Nothing
Dim tag_0702_FQET_01
tag_0702_FQET_01 = HMIRuntime.Tags("0702_FQET_01").Read
strSQL = "INSERT INTO z_tag_0702_FQET_01 (tag_value, created) values(" & tag_0702_FQET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_FQET_01 = Nothing
Dim tag_0702_FQET_02
tag_0702_FQET_02 = HMIRuntime.Tags("0702_FQET_02").Read
strSQL = "INSERT INTO z_tag_0702_FQET_02 (tag_value, created) values(" & tag_0702_FQET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_FQET_02 = Nothing
Dim tag_0702_QET_01
tag_0702_QET_01 = HMIRuntime.Tags("0702_QET_01").Read
strSQL = "INSERT INTO z_tag_0702_QET_01 (tag_value, created) values(" & tag_0702_QET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_QET_01 = Nothing
Dim tag_0702_QET_02
tag_0702_QET_02 = HMIRuntime.Tags("0702_QET_02").Read
strSQL = "INSERT INTO z_tag_0702_QET_02 (tag_value, created) values(" & tag_0702_QET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_QET_02 = Nothing
Dim tag_0702_TET_01
tag_0702_TET_01 = HMIRuntime.Tags("0702_TET_01").Read
strSQL = "INSERT INTO z_tag_0702_TET_01 (tag_value, created) values(" & tag_0702_TET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_TET_01 = Nothing
Dim tag_0702_TET_02
tag_0702_TET_02 = HMIRuntime.Tags("0702_TET_02").Read
strSQL = "INSERT INTO z_tag_0702_TET_02 (tag_value, created) values(" & tag_0702_TET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702_TET_02 = Nothing
Dim tag_0702LET01
tag_0702LET01 = HMIRuntime.Tags("0702LET01").Read
strSQL = "INSERT INTO z_tag_0702LET01 (tag_value, created) values(" & tag_0702LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702LET01 = Nothing
Dim tag_0702LET02
tag_0702LET02 = HMIRuntime.Tags("0702LET02").Read
strSQL = "INSERT INTO z_tag_0702LET02 (tag_value, created) values(" & tag_0702LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702LET02 = Nothing
Dim tag_0702LET03
tag_0702LET03 = HMIRuntime.Tags("0702LET03").Read
strSQL = "INSERT INTO z_tag_0702LET03 (tag_value, created) values(" & tag_0702LET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702LET03 = Nothing
Dim tag_0702M01
tag_0702M01 = HMIRuntime.Tags("0702M01").Read
strSQL = "INSERT INTO z_tag_0702M01 (tag_value, created) values(" & tag_0702M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702M01 = Nothing
Dim tag_0702M02
tag_0702M02 = HMIRuntime.Tags("0702M02").Read
strSQL = "INSERT INTO z_tag_0702M02 (tag_value, created) values(" & tag_0702M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702M02 = Nothing
Dim tag_0702M03
tag_0702M03 = HMIRuntime.Tags("0702M03").Read
strSQL = "INSERT INTO z_tag_0702M03 (tag_value, created) values(" & tag_0702M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702M03 = Nothing
Dim tag_0702PV01
tag_0702PV01 = HMIRuntime.Tags("0702PV01").Read
strSQL = "INSERT INTO z_tag_0702PV01 (tag_value, created) values(" & tag_0702PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV01 = Nothing
Dim tag_0702PV02
tag_0702PV02 = HMIRuntime.Tags("0702PV02").Read
strSQL = "INSERT INTO z_tag_0702PV02 (tag_value, created) values(" & tag_0702PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV02 = Nothing
Dim tag_0702PV03
tag_0702PV03 = HMIRuntime.Tags("0702PV03").Read
strSQL = "INSERT INTO z_tag_0702PV03 (tag_value, created) values(" & tag_0702PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV03 = Nothing
Dim tag_0702PV04
tag_0702PV04 = HMIRuntime.Tags("0702PV04").Read
strSQL = "INSERT INTO z_tag_0702PV04 (tag_value, created) values(" & tag_0702PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV04 = Nothing
Dim tag_0702PV05
tag_0702PV05 = HMIRuntime.Tags("0702PV05").Read
strSQL = "INSERT INTO z_tag_0702PV05 (tag_value, created) values(" & tag_0702PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV05 = Nothing
Dim tag_0702PV06
tag_0702PV06 = HMIRuntime.Tags("0702PV06").Read
strSQL = "INSERT INTO z_tag_0702PV06 (tag_value, created) values(" & tag_0702PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV06 = Nothing
Dim tag_0702PV07
tag_0702PV07 = HMIRuntime.Tags("0702PV07").Read
strSQL = "INSERT INTO z_tag_0702PV07 (tag_value, created) values(" & tag_0702PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV07 = Nothing
Dim tag_0702PV08
tag_0702PV08 = HMIRuntime.Tags("0702PV08").Read
strSQL = "INSERT INTO z_tag_0702PV08 (tag_value, created) values(" & tag_0702PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV08 = Nothing
Dim tag_0702PV09
tag_0702PV09 = HMIRuntime.Tags("0702PV09").Read
strSQL = "INSERT INTO z_tag_0702PV09 (tag_value, created) values(" & tag_0702PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV09 = Nothing
Dim tag_0702PV10
tag_0702PV10 = HMIRuntime.Tags("0702PV10").Read
strSQL = "INSERT INTO z_tag_0702PV10 (tag_value, created) values(" & tag_0702PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV10 = Nothing
Dim tag_0702PV11
tag_0702PV11 = HMIRuntime.Tags("0702PV11").Read
strSQL = "INSERT INTO z_tag_0702PV11 (tag_value, created) values(" & tag_0702PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV11 = Nothing
Dim tag_0702PV12
tag_0702PV12 = HMIRuntime.Tags("0702PV12").Read
strSQL = "INSERT INTO z_tag_0702PV12 (tag_value, created) values(" & tag_0702PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV12 = Nothing
Dim tag_0702PV13
tag_0702PV13 = HMIRuntime.Tags("0702PV13").Read
strSQL = "INSERT INTO z_tag_0702PV13 (tag_value, created) values(" & tag_0702PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV13 = Nothing
Dim tag_0702PV14
tag_0702PV14 = HMIRuntime.Tags("0702PV14").Read
strSQL = "INSERT INTO z_tag_0702PV14 (tag_value, created) values(" & tag_0702PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV14 = Nothing
Dim tag_0702PV15
tag_0702PV15 = HMIRuntime.Tags("0702PV15").Read
strSQL = "INSERT INTO z_tag_0702PV15 (tag_value, created) values(" & tag_0702PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV15 = Nothing
Dim tag_0702PV16
tag_0702PV16 = HMIRuntime.Tags("0702PV16").Read
strSQL = "INSERT INTO z_tag_0702PV16 (tag_value, created) values(" & tag_0702PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV16 = Nothing
Dim tag_0702PV17
tag_0702PV17 = HMIRuntime.Tags("0702PV17").Read
strSQL = "INSERT INTO z_tag_0702PV17 (tag_value, created) values(" & tag_0702PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV17 = Nothing
Dim tag_0702PV18
tag_0702PV18 = HMIRuntime.Tags("0702PV18").Read
strSQL = "INSERT INTO z_tag_0702PV18 (tag_value, created) values(" & tag_0702PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV18 = Nothing
Dim tag_0702PV19
tag_0702PV19 = HMIRuntime.Tags("0702PV19").Read
strSQL = "INSERT INTO z_tag_0702PV19 (tag_value, created) values(" & tag_0702PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV19 = Nothing
Dim tag_0702PV20
tag_0702PV20 = HMIRuntime.Tags("0702PV20").Read
strSQL = "INSERT INTO z_tag_0702PV20 (tag_value, created) values(" & tag_0702PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV20 = Nothing
Dim tag_0702PV21
tag_0702PV21 = HMIRuntime.Tags("0702PV21").Read
strSQL = "INSERT INTO z_tag_0702PV21 (tag_value, created) values(" & tag_0702PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV21 = Nothing
Dim tag_0702PV22
tag_0702PV22 = HMIRuntime.Tags("0702PV22").Read
strSQL = "INSERT INTO z_tag_0702PV22 (tag_value, created) values(" & tag_0702PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV22 = Nothing
Dim tag_0702PV23
tag_0702PV23 = HMIRuntime.Tags("0702PV23").Read
strSQL = "INSERT INTO z_tag_0702PV23 (tag_value, created) values(" & tag_0702PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV23 = Nothing
Dim tag_0702PV24
tag_0702PV24 = HMIRuntime.Tags("0702PV24").Read
strSQL = "INSERT INTO z_tag_0702PV24 (tag_value, created) values(" & tag_0702PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV24 = Nothing
Dim tag_0702PV25
tag_0702PV25 = HMIRuntime.Tags("0702PV25").Read
strSQL = "INSERT INTO z_tag_0702PV25 (tag_value, created) values(" & tag_0702PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV25 = Nothing
Dim tag_0702PV26
tag_0702PV26 = HMIRuntime.Tags("0702PV26").Read
strSQL = "INSERT INTO z_tag_0702PV26 (tag_value, created) values(" & tag_0702PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV26 = Nothing
Dim tag_0702PV27
tag_0702PV27 = HMIRuntime.Tags("0702PV27").Read
strSQL = "INSERT INTO z_tag_0702PV27 (tag_value, created) values(" & tag_0702PV27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV27 = Nothing
Dim tag_0702PV28
tag_0702PV28 = HMIRuntime.Tags("0702PV28").Read
strSQL = "INSERT INTO z_tag_0702PV28 (tag_value, created) values(" & tag_0702PV28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV28 = Nothing
Dim tag_0702PV29
tag_0702PV29 = HMIRuntime.Tags("0702PV29").Read
strSQL = "INSERT INTO z_tag_0702PV29 (tag_value, created) values(" & tag_0702PV29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV29 = Nothing
Dim tag_0702PV30
tag_0702PV30 = HMIRuntime.Tags("0702PV30").Read
strSQL = "INSERT INTO z_tag_0702PV30 (tag_value, created) values(" & tag_0702PV30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV30 = Nothing
Dim tag_0702PV31
tag_0702PV31 = HMIRuntime.Tags("0702PV31").Read
strSQL = "INSERT INTO z_tag_0702PV31 (tag_value, created) values(" & tag_0702PV31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV31 = Nothing
Dim tag_0702PV32
tag_0702PV32 = HMIRuntime.Tags("0702PV32").Read
strSQL = "INSERT INTO z_tag_0702PV32 (tag_value, created) values(" & tag_0702PV32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV32 = Nothing
Dim tag_0702PV33
tag_0702PV33 = HMIRuntime.Tags("0702PV33").Read
strSQL = "INSERT INTO z_tag_0702PV33 (tag_value, created) values(" & tag_0702PV33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV33 = Nothing
Dim tag_0702PV34
tag_0702PV34 = HMIRuntime.Tags("0702PV34").Read
strSQL = "INSERT INTO z_tag_0702PV34 (tag_value, created) values(" & tag_0702PV34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV34 = Nothing
Dim tag_0702PV35
tag_0702PV35 = HMIRuntime.Tags("0702PV35").Read
strSQL = "INSERT INTO z_tag_0702PV35 (tag_value, created) values(" & tag_0702PV35 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV35 = Nothing
Dim tag_0702PV36
tag_0702PV36 = HMIRuntime.Tags("0702PV36").Read
strSQL = "INSERT INTO z_tag_0702PV36 (tag_value, created) values(" & tag_0702PV36 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV36 = Nothing
Dim tag_0702PV37
tag_0702PV37 = HMIRuntime.Tags("0702PV37").Read
strSQL = "INSERT INTO z_tag_0702PV37 (tag_value, created) values(" & tag_0702PV37 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV37 = Nothing
Dim tag_0702PV38
tag_0702PV38 = HMIRuntime.Tags("0702PV38").Read
strSQL = "INSERT INTO z_tag_0702PV38 (tag_value, created) values(" & tag_0702PV38 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV38 = Nothing
Dim tag_0702PV39
tag_0702PV39 = HMIRuntime.Tags("0702PV39").Read
strSQL = "INSERT INTO z_tag_0702PV39 (tag_value, created) values(" & tag_0702PV39 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV39 = Nothing
Dim tag_0702PV40
tag_0702PV40 = HMIRuntime.Tags("0702PV40").Read
strSQL = "INSERT INTO z_tag_0702PV40 (tag_value, created) values(" & tag_0702PV40 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV40 = Nothing
Dim tag_0702PV41
tag_0702PV41 = HMIRuntime.Tags("0702PV41").Read
strSQL = "INSERT INTO z_tag_0702PV41 (tag_value, created) values(" & tag_0702PV41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV41 = Nothing
Dim tag_0702PV42
tag_0702PV42 = HMIRuntime.Tags("0702PV42").Read
strSQL = "INSERT INTO z_tag_0702PV42 (tag_value, created) values(" & tag_0702PV42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV42 = Nothing
Dim tag_0702PV43
tag_0702PV43 = HMIRuntime.Tags("0702PV43").Read
strSQL = "INSERT INTO z_tag_0702PV43 (tag_value, created) values(" & tag_0702PV43 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0702PV43 = Nothing
Dim tag_0703_FQET_01
tag_0703_FQET_01 = HMIRuntime.Tags("0703_FQET_01").Read
strSQL = "INSERT INTO z_tag_0703_FQET_01 (tag_value, created) values(" & tag_0703_FQET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_FQET_01 = Nothing
Dim tag_0703_FQET_02
tag_0703_FQET_02 = HMIRuntime.Tags("0703_FQET_02").Read
strSQL = "INSERT INTO z_tag_0703_FQET_02 (tag_value, created) values(" & tag_0703_FQET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_FQET_02 = Nothing
Dim tag_0703_QET_01
tag_0703_QET_01 = HMIRuntime.Tags("0703_QET_01").Read
strSQL = "INSERT INTO z_tag_0703_QET_01 (tag_value, created) values(" & tag_0703_QET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_QET_01 = Nothing
Dim tag_0703_QET_02
tag_0703_QET_02 = HMIRuntime.Tags("0703_QET_02").Read
strSQL = "INSERT INTO z_tag_0703_QET_02 (tag_value, created) values(" & tag_0703_QET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_QET_02 = Nothing
Dim tag_0703_TET_01
tag_0703_TET_01 = HMIRuntime.Tags("0703_TET_01").Read
strSQL = "INSERT INTO z_tag_0703_TET_01 (tag_value, created) values(" & tag_0703_TET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_TET_01 = Nothing
Dim tag_0703_TET_02
tag_0703_TET_02 = HMIRuntime.Tags("0703_TET_02").Read
strSQL = "INSERT INTO z_tag_0703_TET_02 (tag_value, created) values(" & tag_0703_TET_02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703_TET_02 = Nothing
Dim tag_0703LET01
tag_0703LET01 = HMIRuntime.Tags("0703LET01").Read
strSQL = "INSERT INTO z_tag_0703LET01 (tag_value, created) values(" & tag_0703LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703LET01 = Nothing
Dim tag_0703LET02
tag_0703LET02 = HMIRuntime.Tags("0703LET02").Read
strSQL = "INSERT INTO z_tag_0703LET02 (tag_value, created) values(" & tag_0703LET02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703LET02 = Nothing
Dim tag_0703LET03
tag_0703LET03 = HMIRuntime.Tags("0703LET03").Read
strSQL = "INSERT INTO z_tag_0703LET03 (tag_value, created) values(" & tag_0703LET03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703LET03 = Nothing
Dim tag_0703LET04
tag_0703LET04 = HMIRuntime.Tags("0703LET04").Read
strSQL = "INSERT INTO z_tag_0703LET04 (tag_value, created) values(" & tag_0703LET04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703LET04 = Nothing
Dim tag_0703M01
tag_0703M01 = HMIRuntime.Tags("0703M01").Read
strSQL = "INSERT INTO z_tag_0703M01 (tag_value, created) values(" & tag_0703M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703M01 = Nothing
Dim tag_0703M02
tag_0703M02 = HMIRuntime.Tags("0703M02").Read
strSQL = "INSERT INTO z_tag_0703M02 (tag_value, created) values(" & tag_0703M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703M02 = Nothing
Dim tag_0703M03
tag_0703M03 = HMIRuntime.Tags("0703M03").Read
strSQL = "INSERT INTO z_tag_0703M03 (tag_value, created) values(" & tag_0703M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703M03 = Nothing
Dim tag_0703PV01
tag_0703PV01 = HMIRuntime.Tags("0703PV01").Read
strSQL = "INSERT INTO z_tag_0703PV01 (tag_value, created) values(" & tag_0703PV01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV01 = Nothing
Dim tag_0703PV02
tag_0703PV02 = HMIRuntime.Tags("0703PV02").Read
strSQL = "INSERT INTO z_tag_0703PV02 (tag_value, created) values(" & tag_0703PV02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV02 = Nothing
Dim tag_0703PV03
tag_0703PV03 = HMIRuntime.Tags("0703PV03").Read
strSQL = "INSERT INTO z_tag_0703PV03 (tag_value, created) values(" & tag_0703PV03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV03 = Nothing
Dim tag_0703PV04
tag_0703PV04 = HMIRuntime.Tags("0703PV04").Read
strSQL = "INSERT INTO z_tag_0703PV04 (tag_value, created) values(" & tag_0703PV04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV04 = Nothing
Dim tag_0703PV05
tag_0703PV05 = HMIRuntime.Tags("0703PV05").Read
strSQL = "INSERT INTO z_tag_0703PV05 (tag_value, created) values(" & tag_0703PV05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV05 = Nothing
Dim tag_0703PV06
tag_0703PV06 = HMIRuntime.Tags("0703PV06").Read
strSQL = "INSERT INTO z_tag_0703PV06 (tag_value, created) values(" & tag_0703PV06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV06 = Nothing
Dim tag_0703PV07
tag_0703PV07 = HMIRuntime.Tags("0703PV07").Read
strSQL = "INSERT INTO z_tag_0703PV07 (tag_value, created) values(" & tag_0703PV07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV07 = Nothing
Dim tag_0703PV08
tag_0703PV08 = HMIRuntime.Tags("0703PV08").Read
strSQL = "INSERT INTO z_tag_0703PV08 (tag_value, created) values(" & tag_0703PV08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV08 = Nothing
Dim tag_0703PV09
tag_0703PV09 = HMIRuntime.Tags("0703PV09").Read
strSQL = "INSERT INTO z_tag_0703PV09 (tag_value, created) values(" & tag_0703PV09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV09 = Nothing
Dim tag_0703PV10
tag_0703PV10 = HMIRuntime.Tags("0703PV10").Read
strSQL = "INSERT INTO z_tag_0703PV10 (tag_value, created) values(" & tag_0703PV10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV10 = Nothing
Dim tag_0703PV11
tag_0703PV11 = HMIRuntime.Tags("0703PV11").Read
strSQL = "INSERT INTO z_tag_0703PV11 (tag_value, created) values(" & tag_0703PV11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV11 = Nothing
Dim tag_0703PV12
tag_0703PV12 = HMIRuntime.Tags("0703PV12").Read
strSQL = "INSERT INTO z_tag_0703PV12 (tag_value, created) values(" & tag_0703PV12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV12 = Nothing
Dim tag_0703PV13
tag_0703PV13 = HMIRuntime.Tags("0703PV13").Read
strSQL = "INSERT INTO z_tag_0703PV13 (tag_value, created) values(" & tag_0703PV13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV13 = Nothing
Dim tag_0703PV14
tag_0703PV14 = HMIRuntime.Tags("0703PV14").Read
strSQL = "INSERT INTO z_tag_0703PV14 (tag_value, created) values(" & tag_0703PV14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV14 = Nothing
Dim tag_0703PV15
tag_0703PV15 = HMIRuntime.Tags("0703PV15").Read
strSQL = "INSERT INTO z_tag_0703PV15 (tag_value, created) values(" & tag_0703PV15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV15 = Nothing
Dim tag_0703PV16
tag_0703PV16 = HMIRuntime.Tags("0703PV16").Read
strSQL = "INSERT INTO z_tag_0703PV16 (tag_value, created) values(" & tag_0703PV16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV16 = Nothing
Dim tag_0703PV17
tag_0703PV17 = HMIRuntime.Tags("0703PV17").Read
strSQL = "INSERT INTO z_tag_0703PV17 (tag_value, created) values(" & tag_0703PV17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV17 = Nothing
Dim tag_0703PV18
tag_0703PV18 = HMIRuntime.Tags("0703PV18").Read
strSQL = "INSERT INTO z_tag_0703PV18 (tag_value, created) values(" & tag_0703PV18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV18 = Nothing
Dim tag_0703PV19
tag_0703PV19 = HMIRuntime.Tags("0703PV19").Read
strSQL = "INSERT INTO z_tag_0703PV19 (tag_value, created) values(" & tag_0703PV19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV19 = Nothing
Dim tag_0703PV20
tag_0703PV20 = HMIRuntime.Tags("0703PV20").Read
strSQL = "INSERT INTO z_tag_0703PV20 (tag_value, created) values(" & tag_0703PV20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV20 = Nothing
Dim tag_0703PV21
tag_0703PV21 = HMIRuntime.Tags("0703PV21").Read
strSQL = "INSERT INTO z_tag_0703PV21 (tag_value, created) values(" & tag_0703PV21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV21 = Nothing
Dim tag_0703PV22
tag_0703PV22 = HMIRuntime.Tags("0703PV22").Read
strSQL = "INSERT INTO z_tag_0703PV22 (tag_value, created) values(" & tag_0703PV22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV22 = Nothing
Dim tag_0703PV23
tag_0703PV23 = HMIRuntime.Tags("0703PV23").Read
strSQL = "INSERT INTO z_tag_0703PV23 (tag_value, created) values(" & tag_0703PV23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV23 = Nothing
Dim tag_0703PV24
tag_0703PV24 = HMIRuntime.Tags("0703PV24").Read
strSQL = "INSERT INTO z_tag_0703PV24 (tag_value, created) values(" & tag_0703PV24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV24 = Nothing
Dim tag_0703PV25
tag_0703PV25 = HMIRuntime.Tags("0703PV25").Read
strSQL = "INSERT INTO z_tag_0703PV25 (tag_value, created) values(" & tag_0703PV25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV25 = Nothing
Dim tag_0703PV26
tag_0703PV26 = HMIRuntime.Tags("0703PV26").Read
strSQL = "INSERT INTO z_tag_0703PV26 (tag_value, created) values(" & tag_0703PV26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV26 = Nothing
Dim tag_0703PV27
tag_0703PV27 = HMIRuntime.Tags("0703PV27").Read
strSQL = "INSERT INTO z_tag_0703PV27 (tag_value, created) values(" & tag_0703PV27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV27 = Nothing
Dim tag_0703PV28
tag_0703PV28 = HMIRuntime.Tags("0703PV28").Read
strSQL = "INSERT INTO z_tag_0703PV28 (tag_value, created) values(" & tag_0703PV28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV28 = Nothing
Dim tag_0703PV29
tag_0703PV29 = HMIRuntime.Tags("0703PV29").Read
strSQL = "INSERT INTO z_tag_0703PV29 (tag_value, created) values(" & tag_0703PV29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV29 = Nothing
Dim tag_0703PV30
tag_0703PV30 = HMIRuntime.Tags("0703PV30").Read
strSQL = "INSERT INTO z_tag_0703PV30 (tag_value, created) values(" & tag_0703PV30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV30 = Nothing
Dim tag_0703PV31
tag_0703PV31 = HMIRuntime.Tags("0703PV31").Read
strSQL = "INSERT INTO z_tag_0703PV31 (tag_value, created) values(" & tag_0703PV31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV31 = Nothing
Dim tag_0703PV32
tag_0703PV32 = HMIRuntime.Tags("0703PV32").Read
strSQL = "INSERT INTO z_tag_0703PV32 (tag_value, created) values(" & tag_0703PV32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV32 = Nothing
Dim tag_0703PV33
tag_0703PV33 = HMIRuntime.Tags("0703PV33").Read
strSQL = "INSERT INTO z_tag_0703PV33 (tag_value, created) values(" & tag_0703PV33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV33 = Nothing
Dim tag_0703PV34
tag_0703PV34 = HMIRuntime.Tags("0703PV34").Read
strSQL = "INSERT INTO z_tag_0703PV34 (tag_value, created) values(" & tag_0703PV34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV34 = Nothing
Dim tag_0703PV35
tag_0703PV35 = HMIRuntime.Tags("0703PV35").Read
strSQL = "INSERT INTO z_tag_0703PV35 (tag_value, created) values(" & tag_0703PV35 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV35 = Nothing
Dim tag_0703PV36
tag_0703PV36 = HMIRuntime.Tags("0703PV36").Read
strSQL = "INSERT INTO z_tag_0703PV36 (tag_value, created) values(" & tag_0703PV36 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV36 = Nothing
Dim tag_0703PV37
tag_0703PV37 = HMIRuntime.Tags("0703PV37").Read
strSQL = "INSERT INTO z_tag_0703PV37 (tag_value, created) values(" & tag_0703PV37 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV37 = Nothing
Dim tag_0703PV38
tag_0703PV38 = HMIRuntime.Tags("0703PV38").Read
strSQL = "INSERT INTO z_tag_0703PV38 (tag_value, created) values(" & tag_0703PV38 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV38 = Nothing
Dim tag_0703PV39
tag_0703PV39 = HMIRuntime.Tags("0703PV39").Read
strSQL = "INSERT INTO z_tag_0703PV39 (tag_value, created) values(" & tag_0703PV39 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV39 = Nothing
Dim tag_0703PV40
tag_0703PV40 = HMIRuntime.Tags("0703PV40").Read
strSQL = "INSERT INTO z_tag_0703PV40 (tag_value, created) values(" & tag_0703PV40 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV40 = Nothing
Dim tag_0703PV41
tag_0703PV41 = HMIRuntime.Tags("0703PV41").Read
strSQL = "INSERT INTO z_tag_0703PV41 (tag_value, created) values(" & tag_0703PV41 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV41 = Nothing
Dim tag_0703PV42
tag_0703PV42 = HMIRuntime.Tags("0703PV42").Read
strSQL = "INSERT INTO z_tag_0703PV42 (tag_value, created) values(" & tag_0703PV42 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV42 = Nothing
Dim tag_0703PV43
tag_0703PV43 = HMIRuntime.Tags("0703PV43").Read
strSQL = "INSERT INTO z_tag_0703PV43 (tag_value, created) values(" & tag_0703PV43 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV43 = Nothing
Dim tag_0703PV44
tag_0703PV44 = HMIRuntime.Tags("0703PV44").Read
strSQL = "INSERT INTO z_tag_0703PV44 (tag_value, created) values(" & tag_0703PV44 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV44 = Nothing
Dim tag_0703PV45
tag_0703PV45 = HMIRuntime.Tags("0703PV45").Read
strSQL = "INSERT INTO z_tag_0703PV45 (tag_value, created) values(" & tag_0703PV45 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV45 = Nothing
Dim tag_0703PV46
tag_0703PV46 = HMIRuntime.Tags("0703PV46").Read
strSQL = "INSERT INTO z_tag_0703PV46 (tag_value, created) values(" & tag_0703PV46 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV46 = Nothing
Dim tag_0703PV47
tag_0703PV47 = HMIRuntime.Tags("0703PV47").Read
strSQL = "INSERT INTO z_tag_0703PV47 (tag_value, created) values(" & tag_0703PV47 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV47 = Nothing
Dim tag_0703PV48
tag_0703PV48 = HMIRuntime.Tags("0703PV48").Read
strSQL = "INSERT INTO z_tag_0703PV48 (tag_value, created) values(" & tag_0703PV48 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV48 = Nothing
Dim tag_0703PV49
tag_0703PV49 = HMIRuntime.Tags("0703PV49").Read
strSQL = "INSERT INTO z_tag_0703PV49 (tag_value, created) values(" & tag_0703PV49 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV49 = Nothing
Dim tag_0703PV50
tag_0703PV50 = HMIRuntime.Tags("0703PV50").Read
strSQL = "INSERT INTO z_tag_0703PV50 (tag_value, created) values(" & tag_0703PV50 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV50 = Nothing
Dim tag_0703PV51
tag_0703PV51 = HMIRuntime.Tags("0703PV51").Read
strSQL = "INSERT INTO z_tag_0703PV51 (tag_value, created) values(" & tag_0703PV51 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0703PV51 = Nothing
Dim tag_0704_PET_01
tag_0704_PET_01 = HMIRuntime.Tags("0704_PET_01").Read
strSQL = "INSERT INTO z_tag_0704_PET_01 (tag_value, created) values(" & tag_0704_PET_01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0704_PET_01 = Nothing
Dim tag_0704M01
tag_0704M01 = HMIRuntime.Tags("0704M01").Read
strSQL = "INSERT INTO z_tag_0704M01 (tag_value, created) values(" & tag_0704M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0704M01 = Nothing
Dim tag_0704M02
tag_0704M02 = HMIRuntime.Tags("0704M02").Read
strSQL = "INSERT INTO z_tag_0704M02 (tag_value, created) values(" & tag_0704M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0704M02 = Nothing
Dim tag_0704M03
tag_0704M03 = HMIRuntime.Tags("0704M03").Read
strSQL = "INSERT INTO z_tag_0704M03 (tag_value, created) values(" & tag_0704M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_0704M03 = Nothing
Dim tag_1001LET01
tag_1001LET01 = HMIRuntime.Tags("1001LET01").Read
strSQL = "INSERT INTO z_tag_1001LET01 (tag_value, created) values(" & tag_1001LET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_1001LET01 = Nothing
Dim tag_1001TET01
tag_1001TET01 = HMIRuntime.Tags("1001TET01").Read
strSQL = "INSERT INTO z_tag_1001TET01 (tag_value, created) values(" & tag_1001TET01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_1001TET01 = Nothing
Dim tag_AS_tank01
tag_AS_tank01 = HMIRuntime.Tags("AS_tank01").Read
strSQL = "INSERT INTO z_tag_AS_tank01 (tag_value, created) values(" & tag_AS_tank01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_AS_tank01 = Nothing
Dim tag_AS_tank02
tag_AS_tank02 = HMIRuntime.Tags("AS_tank02").Read
strSQL = "INSERT INTO z_tag_AS_tank02 (tag_value, created) values(" & tag_AS_tank02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_AS_tank02 = Nothing
Dim tag_AS_tank03
tag_AS_tank03 = HMIRuntime.Tags("AS_tank03").Read
strSQL = "INSERT INTO z_tag_AS_tank03 (tag_value, created) values(" & tag_AS_tank03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_AS_tank03 = Nothing
Dim tag_AS_tank04
tag_AS_tank04 = HMIRuntime.Tags("AS_tank04").Read
strSQL = "INSERT INTO z_tag_AS_tank04 (tag_value, created) values(" & tag_AS_tank04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_AS_tank04 = Nothing
Dim tag_baorongduongcipmen_hoi
tag_baorongduongcipmen_hoi = HMIRuntime.Tags("baorongduongcipmen_hoi").Read
strSQL = "INSERT INTO z_tag_baorongduongcipmen_hoi (tag_value, created) values(" & tag_baorongduongcipmen_hoi & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baorongduongcipmen_hoi = Nothing
Dim tag_baorongduongcipmen_lanh
tag_baorongduongcipmen_lanh = HMIRuntime.Tags("baorongduongcipmen_lanh").Read
strSQL = "INSERT INTO z_tag_baorongduongcipmen_lanh (tag_value, created) values(" & tag_baorongduongcipmen_lanh & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baorongduongcipmen_lanh = Nothing
Dim tag_baorongduongcipmen_lanhhoi
tag_baorongduongcipmen_lanhhoi = HMIRuntime.Tags("baorongduongcipmen_lanhhoi").Read
strSQL = "INSERT INTO z_tag_baorongduongcipmen_lanhhoi (tag_value, created) values(" & tag_baorongduongcipmen_lanhhoi & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baorongduongcipmen_lanhhoi = Nothing
Dim tag_baorongduongcipmen_nong
tag_baorongduongcipmen_nong = HMIRuntime.Tags("baorongduongcipmen_nong").Read
strSQL = "INSERT INTO z_tag_baorongduongcipmen_nong (tag_value, created) values(" & tag_baorongduongcipmen_nong & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baorongduongcipmen_nong = Nothing
Dim tag_baorongduongcipmen_nonghoi
tag_baorongduongcipmen_nonghoi = HMIRuntime.Tags("baorongduongcipmen_nonghoi").Read
strSQL = "INSERT INTO z_tag_baorongduongcipmen_nonghoi (tag_value, created) values(" & tag_baorongduongcipmen_nonghoi & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_baorongduongcipmen_nonghoi = Nothing
Dim tag_Beer2Filter_run
tag_Beer2Filter_run = HMIRuntime.Tags("Beer2Filter_run").Read
strSQL = "INSERT INTO z_tag_Beer2Filter_run (tag_value, created) values(" & tag_Beer2Filter_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Beer2Filter_run = Nothing
Dim tag_Beer2Filter_seq
tag_Beer2Filter_seq = HMIRuntime.Tags("Beer2Filter_seq").Read
strSQL = "INSERT INTO z_tag_Beer2Filter_seq (tag_value, created) values(" & tag_Beer2Filter_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Beer2Filter_seq = Nothing
Dim tag_Beercominghere
tag_Beercominghere = HMIRuntime.Tags("Beercominghere").Read
strSQL = "INSERT INTO z_tag_Beercominghere (tag_value, created) values(" & tag_Beercominghere & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Beercominghere = Nothing
Dim tag_bieukien_T1
tag_bieukien_T1 = HMIRuntime.Tags("bieukien_T1").Read
strSQL = "INSERT INTO z_tag_bieukien_T1 (tag_value, created) values(" & tag_bieukien_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T1 = Nothing
Dim tag_bieukien_T2
tag_bieukien_T2 = HMIRuntime.Tags("bieukien_T2").Read
strSQL = "INSERT INTO z_tag_bieukien_T2 (tag_value, created) values(" & tag_bieukien_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T2 = Nothing
Dim tag_bieukien_T3
tag_bieukien_T3 = HMIRuntime.Tags("bieukien_T3").Read
strSQL = "INSERT INTO z_tag_bieukien_T3 (tag_value, created) values(" & tag_bieukien_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T3 = Nothing
Dim tag_bieukien_T4
tag_bieukien_T4 = HMIRuntime.Tags("bieukien_T4").Read
strSQL = "INSERT INTO z_tag_bieukien_T4 (tag_value, created) values(" & tag_bieukien_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T4 = Nothing
Dim tag_bieukien_T5
tag_bieukien_T5 = HMIRuntime.Tags("bieukien_T5").Read
strSQL = "INSERT INTO z_tag_bieukien_T5 (tag_value, created) values(" & tag_bieukien_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T5 = Nothing
Dim tag_bieukien_T6
tag_bieukien_T6 = HMIRuntime.Tags("bieukien_T6").Read
strSQL = "INSERT INTO z_tag_bieukien_T6 (tag_value, created) values(" & tag_bieukien_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T6 = Nothing
Dim tag_bieukien_T7
tag_bieukien_T7 = HMIRuntime.Tags("bieukien_T7").Read
strSQL = "INSERT INTO z_tag_bieukien_T7 (tag_value, created) values(" & tag_bieukien_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T7 = Nothing
Dim tag_bieukien_T8
tag_bieukien_T8 = HMIRuntime.Tags("bieukien_T8").Read
strSQL = "INSERT INTO z_tag_bieukien_T8 (tag_value, created) values(" & tag_bieukien_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T8 = Nothing
Dim tag_bieukien_T9
tag_bieukien_T9 = HMIRuntime.Tags("bieukien_T9").Read
strSQL = "INSERT INTO z_tag_bieukien_T9 (tag_value, created) values(" & tag_bieukien_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T9 = Nothing
Dim tag_bieukien_T10
tag_bieukien_T10 = HMIRuntime.Tags("bieukien_T10").Read
strSQL = "INSERT INTO z_tag_bieukien_T10 (tag_value, created) values(" & tag_bieukien_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T10 = Nothing
Dim tag_bieukien_T11
tag_bieukien_T11 = HMIRuntime.Tags("bieukien_T11").Read
strSQL = "INSERT INTO z_tag_bieukien_T11 (tag_value, created) values(" & tag_bieukien_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T11 = Nothing
Dim tag_bieukien_T12
tag_bieukien_T12 = HMIRuntime.Tags("bieukien_T12").Read
strSQL = "INSERT INTO z_tag_bieukien_T12 (tag_value, created) values(" & tag_bieukien_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T12 = Nothing
Dim tag_bieukien_T13
tag_bieukien_T13 = HMIRuntime.Tags("bieukien_T13").Read
strSQL = "INSERT INTO z_tag_bieukien_T13 (tag_value, created) values(" & tag_bieukien_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T13 = Nothing
Dim tag_bieukien_T14
tag_bieukien_T14 = HMIRuntime.Tags("bieukien_T14").Read
strSQL = "INSERT INTO z_tag_bieukien_T14 (tag_value, created) values(" & tag_bieukien_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T14 = Nothing
Dim tag_bieukien_T15
tag_bieukien_T15 = HMIRuntime.Tags("bieukien_T15").Read
strSQL = "INSERT INTO z_tag_bieukien_T15 (tag_value, created) values(" & tag_bieukien_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T15 = Nothing
Dim tag_bieukien_T16
tag_bieukien_T16 = HMIRuntime.Tags("bieukien_T16").Read
strSQL = "INSERT INTO z_tag_bieukien_T16 (tag_value, created) values(" & tag_bieukien_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T16 = Nothing
Dim tag_bieukien_T17
tag_bieukien_T17 = HMIRuntime.Tags("bieukien_T17").Read
strSQL = "INSERT INTO z_tag_bieukien_T17 (tag_value, created) values(" & tag_bieukien_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T17 = Nothing
Dim tag_bieukien_T18
tag_bieukien_T18 = HMIRuntime.Tags("bieukien_T18").Read
strSQL = "INSERT INTO z_tag_bieukien_T18 (tag_value, created) values(" & tag_bieukien_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T18 = Nothing
Dim tag_bieukien_T19
tag_bieukien_T19 = HMIRuntime.Tags("bieukien_T19").Read
strSQL = "INSERT INTO z_tag_bieukien_T19 (tag_value, created) values(" & tag_bieukien_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T19 = Nothing
Dim tag_bieukien_T20
tag_bieukien_T20 = HMIRuntime.Tags("bieukien_T20").Read
strSQL = "INSERT INTO z_tag_bieukien_T20 (tag_value, created) values(" & tag_bieukien_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T20 = Nothing
Dim tag_bieukien_T21
tag_bieukien_T21 = HMIRuntime.Tags("bieukien_T21").Read
strSQL = "INSERT INTO z_tag_bieukien_T21 (tag_value, created) values(" & tag_bieukien_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T21 = Nothing
Dim tag_bieukien_T22
tag_bieukien_T22 = HMIRuntime.Tags("bieukien_T22").Read
strSQL = "INSERT INTO z_tag_bieukien_T22 (tag_value, created) values(" & tag_bieukien_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T22 = Nothing
Dim tag_bieukien_T23
tag_bieukien_T23 = HMIRuntime.Tags("bieukien_T23").Read
strSQL = "INSERT INTO z_tag_bieukien_T23 (tag_value, created) values(" & tag_bieukien_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T23 = Nothing
Dim tag_bieukien_T24
tag_bieukien_T24 = HMIRuntime.Tags("bieukien_T24").Read
strSQL = "INSERT INTO z_tag_bieukien_T24 (tag_value, created) values(" & tag_bieukien_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T24 = Nothing
Dim tag_bieukien_T25
tag_bieukien_T25 = HMIRuntime.Tags("bieukien_T25").Read
strSQL = "INSERT INTO z_tag_bieukien_T25 (tag_value, created) values(" & tag_bieukien_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T25 = Nothing
Dim tag_bieukien_T26
tag_bieukien_T26 = HMIRuntime.Tags("bieukien_T26").Read
strSQL = "INSERT INTO z_tag_bieukien_T26 (tag_value, created) values(" & tag_bieukien_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T26 = Nothing
Dim tag_bieukien_T27
tag_bieukien_T27 = HMIRuntime.Tags("bieukien_T27").Read
strSQL = "INSERT INTO z_tag_bieukien_T27 (tag_value, created) values(" & tag_bieukien_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T27 = Nothing
Dim tag_bieukien_T28
tag_bieukien_T28 = HMIRuntime.Tags("bieukien_T28").Read
strSQL = "INSERT INTO z_tag_bieukien_T28 (tag_value, created) values(" & tag_bieukien_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T28 = Nothing
Dim tag_bieukien_T29
tag_bieukien_T29 = HMIRuntime.Tags("bieukien_T29").Read
strSQL = "INSERT INTO z_tag_bieukien_T29 (tag_value, created) values(" & tag_bieukien_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T29 = Nothing
Dim tag_bieukien_T30
tag_bieukien_T30 = HMIRuntime.Tags("bieukien_T30").Read
strSQL = "INSERT INTO z_tag_bieukien_T30 (tag_value, created) values(" & tag_bieukien_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_bieukien_T30 = Nothing
Dim tag_Canhkhuaytrumen_1
tag_Canhkhuaytrumen_1 = HMIRuntime.Tags("Canhkhuaytrumen_1").Read
strSQL = "INSERT INTO z_tag_Canhkhuaytrumen_1 (tag_value, created) values(" & tag_Canhkhuaytrumen_1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Canhkhuaytrumen_1 = Nothing
Dim tag_Canhkhuaytrumen_2
tag_Canhkhuaytrumen_2 = HMIRuntime.Tags("Canhkhuaytrumen_2").Read
strSQL = "INSERT INTO z_tag_Canhkhuaytrumen_2 (tag_value, created) values(" & tag_Canhkhuaytrumen_2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Canhkhuaytrumen_2 = Nothing
Dim tag_CIP_TACHMEN70hl
tag_CIP_TACHMEN70hl = HMIRuntime.Tags("CIP_TACHMEN70hl").Read
strSQL = "INSERT INTO z_tag_CIP_TACHMEN70hl (tag_value, created) values(" & tag_CIP_TACHMEN70hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIP_TACHMEN70hl = Nothing
Dim tag_CIP_TACHMEN100hl
tag_CIP_TACHMEN100hl = HMIRuntime.Tags("CIP_TACHMEN100hl").Read
strSQL = "INSERT INTO z_tag_CIP_TACHMEN100hl (tag_value, created) values(" & tag_CIP_TACHMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIP_TACHMEN100hl = Nothing
Dim tag_CIP_TANKMEN100hl
tag_CIP_TANKMEN100hl = HMIRuntime.Tags("CIP_TANKMEN100hl").Read
strSQL = "INSERT INTO z_tag_CIP_TANKMEN100hl (tag_value, created) values(" & tag_CIP_TANKMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIP_TANKMEN100hl = Nothing
Dim tag_CIP_TBF08
tag_CIP_TBF08 = HMIRuntime.Tags("CIP_TBF08").Read
strSQL = "INSERT INTO z_tag_CIP_TBF08 (tag_value, created) values(" & tag_CIP_TBF08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIP_TBF08 = Nothing
Dim tag_CIP_TBF09
tag_CIP_TBF09 = HMIRuntime.Tags("CIP_TBF09").Read
strSQL = "INSERT INTO z_tag_CIP_TBF09 (tag_value, created) values(" & tag_CIP_TBF09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CIP_TBF09 = Nothing
Dim tag_CipBTF_Filler_run
tag_CipBTF_Filler_run = HMIRuntime.Tags("CipBTF_Filler_run").Read
strSQL = "INSERT INTO z_tag_CipBTF_Filler_run (tag_value, created) values(" & tag_CipBTF_Filler_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipBTF_Filler_run = Nothing
Dim tag_CipBTF_Filler_seq
tag_CipBTF_Filler_seq = HMIRuntime.Tags("CipBTF_Filler_seq").Read
strSQL = "INSERT INTO z_tag_CipBTF_Filler_seq (tag_value, created) values(" & tag_CipBTF_Filler_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipBTF_Filler_seq = Nothing
Dim tag_CipCCTYT_run
tag_CipCCTYT_run = HMIRuntime.Tags("CipCCTYT_run").Read
strSQL = "INSERT INTO z_tag_CipCCTYT_run (tag_value, created) values(" & tag_CipCCTYT_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipCCTYT_run = Nothing
Dim tag_CipCCTYT_seq
tag_CipCCTYT_seq = HMIRuntime.Tags("CipCCTYT_seq").Read
strSQL = "INSERT INTO z_tag_CipCCTYT_seq (tag_value, created) values(" & tag_CipCCTYT_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipCCTYT_seq = Nothing
Dim tag_CipKG_DATank_run
tag_CipKG_DATank_run = HMIRuntime.Tags("CipKG_DATank_run").Read
strSQL = "INSERT INTO z_tag_CipKG_DATank_run (tag_value, created) values(" & tag_CipKG_DATank_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipKG_DATank_run = Nothing
Dim tag_CipKG_DATank_seq
tag_CipKG_DATank_seq = HMIRuntime.Tags("CipKG_DATank_seq").Read
strSQL = "INSERT INTO z_tag_CipKG_DATank_seq (tag_value, created) values(" & tag_CipKG_DATank_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipKG_DATank_seq = Nothing
Dim tag_CipKronesFill_run
tag_CipKronesFill_run = HMIRuntime.Tags("CipKronesFill_run").Read
strSQL = "INSERT INTO z_tag_CipKronesFill_run (tag_value, created) values(" & tag_CipKronesFill_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipKronesFill_run = Nothing
Dim tag_CipKronesFill_seq
tag_CipKronesFill_seq = HMIRuntime.Tags("CipKronesFill_seq").Read
strSQL = "INSERT INTO z_tag_CipKronesFill_seq (tag_value, created) values(" & tag_CipKronesFill_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipKronesFill_seq = Nothing
Dim tag_CipPipePre_F_run
tag_CipPipePre_F_run = HMIRuntime.Tags("CipPipePre_F_run").Read
strSQL = "INSERT INTO z_tag_CipPipePre_F_run (tag_value, created) values(" & tag_CipPipePre_F_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipPipePre_F_run = Nothing
Dim tag_CipPipePre_F_seq
tag_CipPipePre_F_seq = HMIRuntime.Tags("CipPipePre_F_seq").Read
strSQL = "INSERT INTO z_tag_CipPipePre_F_seq (tag_value, created) values(" & tag_CipPipePre_F_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipPipePre_F_seq = Nothing
Dim tag_CipPVPP_DASys_run
tag_CipPVPP_DASys_run = HMIRuntime.Tags("CipPVPP_DASys_run").Read
strSQL = "INSERT INTO z_tag_CipPVPP_DASys_run (tag_value, created) values(" & tag_CipPVPP_DASys_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipPVPP_DASys_run = Nothing
Dim tag_CipPVPP_DASys_seq
tag_CipPVPP_DASys_seq = HMIRuntime.Tags("CipPVPP_DASys_seq").Read
strSQL = "INSERT INTO z_tag_CipPVPP_DASys_seq (tag_value, created) values(" & tag_CipPVPP_DASys_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CipPVPP_DASys_seq = Nothing
Dim tag_CO2_T1
tag_CO2_T1 = HMIRuntime.Tags("CO2_T1").Read
strSQL = "INSERT INTO z_tag_CO2_T1 (tag_value, created) values(" & tag_CO2_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T1 = Nothing
Dim tag_CO2_T2
tag_CO2_T2 = HMIRuntime.Tags("CO2_T2").Read
strSQL = "INSERT INTO z_tag_CO2_T2 (tag_value, created) values(" & tag_CO2_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T2 = Nothing
Dim tag_CO2_T3
tag_CO2_T3 = HMIRuntime.Tags("CO2_T3").Read
strSQL = "INSERT INTO z_tag_CO2_T3 (tag_value, created) values(" & tag_CO2_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T3 = Nothing
Dim tag_CO2_T4
tag_CO2_T4 = HMIRuntime.Tags("CO2_T4").Read
strSQL = "INSERT INTO z_tag_CO2_T4 (tag_value, created) values(" & tag_CO2_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T4 = Nothing
Dim tag_CO2_T5
tag_CO2_T5 = HMIRuntime.Tags("CO2_T5").Read
strSQL = "INSERT INTO z_tag_CO2_T5 (tag_value, created) values(" & tag_CO2_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T5 = Nothing
Dim tag_CO2_T6
tag_CO2_T6 = HMIRuntime.Tags("CO2_T6").Read
strSQL = "INSERT INTO z_tag_CO2_T6 (tag_value, created) values(" & tag_CO2_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T6 = Nothing
Dim tag_CO2_T7
tag_CO2_T7 = HMIRuntime.Tags("CO2_T7").Read
strSQL = "INSERT INTO z_tag_CO2_T7 (tag_value, created) values(" & tag_CO2_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T7 = Nothing
Dim tag_CO2_T8
tag_CO2_T8 = HMIRuntime.Tags("CO2_T8").Read
strSQL = "INSERT INTO z_tag_CO2_T8 (tag_value, created) values(" & tag_CO2_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T8 = Nothing
Dim tag_CO2_T9
tag_CO2_T9 = HMIRuntime.Tags("CO2_T9").Read
strSQL = "INSERT INTO z_tag_CO2_T9 (tag_value, created) values(" & tag_CO2_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T9 = Nothing
Dim tag_CO2_T10
tag_CO2_T10 = HMIRuntime.Tags("CO2_T10").Read
strSQL = "INSERT INTO z_tag_CO2_T10 (tag_value, created) values(" & tag_CO2_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T10 = Nothing
Dim tag_CO2_T11
tag_CO2_T11 = HMIRuntime.Tags("CO2_T11").Read
strSQL = "INSERT INTO z_tag_CO2_T11 (tag_value, created) values(" & tag_CO2_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T11 = Nothing
Dim tag_CO2_T12
tag_CO2_T12 = HMIRuntime.Tags("CO2_T12").Read
strSQL = "INSERT INTO z_tag_CO2_T12 (tag_value, created) values(" & tag_CO2_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T12 = Nothing
Dim tag_CO2_T13
tag_CO2_T13 = HMIRuntime.Tags("CO2_T13").Read
strSQL = "INSERT INTO z_tag_CO2_T13 (tag_value, created) values(" & tag_CO2_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T13 = Nothing
Dim tag_CO2_T14
tag_CO2_T14 = HMIRuntime.Tags("CO2_T14").Read
strSQL = "INSERT INTO z_tag_CO2_T14 (tag_value, created) values(" & tag_CO2_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T14 = Nothing
Dim tag_CO2_T15
tag_CO2_T15 = HMIRuntime.Tags("CO2_T15").Read
strSQL = "INSERT INTO z_tag_CO2_T15 (tag_value, created) values(" & tag_CO2_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T15 = Nothing
Dim tag_CO2_T16
tag_CO2_T16 = HMIRuntime.Tags("CO2_T16").Read
strSQL = "INSERT INTO z_tag_CO2_T16 (tag_value, created) values(" & tag_CO2_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T16 = Nothing
Dim tag_CO2_T17
tag_CO2_T17 = HMIRuntime.Tags("CO2_T17").Read
strSQL = "INSERT INTO z_tag_CO2_T17 (tag_value, created) values(" & tag_CO2_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T17 = Nothing
Dim tag_CO2_T18
tag_CO2_T18 = HMIRuntime.Tags("CO2_T18").Read
strSQL = "INSERT INTO z_tag_CO2_T18 (tag_value, created) values(" & tag_CO2_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T18 = Nothing
Dim tag_CO2_T19
tag_CO2_T19 = HMIRuntime.Tags("CO2_T19").Read
strSQL = "INSERT INTO z_tag_CO2_T19 (tag_value, created) values(" & tag_CO2_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T19 = Nothing
Dim tag_CO2_T20
tag_CO2_T20 = HMIRuntime.Tags("CO2_T20").Read
strSQL = "INSERT INTO z_tag_CO2_T20 (tag_value, created) values(" & tag_CO2_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T20 = Nothing
Dim tag_CO2_T21
tag_CO2_T21 = HMIRuntime.Tags("CO2_T21").Read
strSQL = "INSERT INTO z_tag_CO2_T21 (tag_value, created) values(" & tag_CO2_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T21 = Nothing
Dim tag_CO2_T22
tag_CO2_T22 = HMIRuntime.Tags("CO2_T22").Read
strSQL = "INSERT INTO z_tag_CO2_T22 (tag_value, created) values(" & tag_CO2_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T22 = Nothing
Dim tag_CO2_T23
tag_CO2_T23 = HMIRuntime.Tags("CO2_T23").Read
strSQL = "INSERT INTO z_tag_CO2_T23 (tag_value, created) values(" & tag_CO2_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T23 = Nothing
Dim tag_CO2_T24
tag_CO2_T24 = HMIRuntime.Tags("CO2_T24").Read
strSQL = "INSERT INTO z_tag_CO2_T24 (tag_value, created) values(" & tag_CO2_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T24 = Nothing
Dim tag_CO2_T25
tag_CO2_T25 = HMIRuntime.Tags("CO2_T25").Read
strSQL = "INSERT INTO z_tag_CO2_T25 (tag_value, created) values(" & tag_CO2_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T25 = Nothing
Dim tag_CO2_T26
tag_CO2_T26 = HMIRuntime.Tags("CO2_T26").Read
strSQL = "INSERT INTO z_tag_CO2_T26 (tag_value, created) values(" & tag_CO2_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T26 = Nothing
Dim tag_CO2_T27
tag_CO2_T27 = HMIRuntime.Tags("CO2_T27").Read
strSQL = "INSERT INTO z_tag_CO2_T27 (tag_value, created) values(" & tag_CO2_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T27 = Nothing
Dim tag_CO2_T28
tag_CO2_T28 = HMIRuntime.Tags("CO2_T28").Read
strSQL = "INSERT INTO z_tag_CO2_T28 (tag_value, created) values(" & tag_CO2_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T28 = Nothing
Dim tag_CO2_T29
tag_CO2_T29 = HMIRuntime.Tags("CO2_T29").Read
strSQL = "INSERT INTO z_tag_CO2_T29 (tag_value, created) values(" & tag_CO2_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T29 = Nothing
Dim tag_CO2_T30
tag_CO2_T30 = HMIRuntime.Tags("CO2_T30").Read
strSQL = "INSERT INTO z_tag_CO2_T30 (tag_value, created) values(" & tag_CO2_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_CO2_T30 = Nothing
Dim tag_date_T1
tag_date_T1 = HMIRuntime.Tags("date_T1").Read
strSQL = "INSERT INTO z_tag_date_T1 (tag_value, created) values(" & tag_date_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T1 = Nothing
Dim tag_date_T2
tag_date_T2 = HMIRuntime.Tags("date_T2").Read
strSQL = "INSERT INTO z_tag_date_T2 (tag_value, created) values(" & tag_date_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T2 = Nothing
Dim tag_date_T3
tag_date_T3 = HMIRuntime.Tags("date_T3").Read
strSQL = "INSERT INTO z_tag_date_T3 (tag_value, created) values(" & tag_date_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T3 = Nothing
Dim tag_date_T4
tag_date_T4 = HMIRuntime.Tags("date_T4").Read
strSQL = "INSERT INTO z_tag_date_T4 (tag_value, created) values(" & tag_date_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T4 = Nothing
Dim tag_date_T5
tag_date_T5 = HMIRuntime.Tags("date_T5").Read
strSQL = "INSERT INTO z_tag_date_T5 (tag_value, created) values(" & tag_date_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T5 = Nothing
Dim tag_date_T6
tag_date_T6 = HMIRuntime.Tags("date_T6").Read
strSQL = "INSERT INTO z_tag_date_T6 (tag_value, created) values(" & tag_date_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T6 = Nothing
Dim tag_date_T7
tag_date_T7 = HMIRuntime.Tags("date_T7").Read
strSQL = "INSERT INTO z_tag_date_T7 (tag_value, created) values(" & tag_date_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T7 = Nothing
Dim tag_date_T8
tag_date_T8 = HMIRuntime.Tags("date_T8").Read
strSQL = "INSERT INTO z_tag_date_T8 (tag_value, created) values(" & tag_date_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T8 = Nothing
Dim tag_date_T9
tag_date_T9 = HMIRuntime.Tags("date_T9").Read
strSQL = "INSERT INTO z_tag_date_T9 (tag_value, created) values(" & tag_date_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T9 = Nothing
Dim tag_date_T10
tag_date_T10 = HMIRuntime.Tags("date_T10").Read
strSQL = "INSERT INTO z_tag_date_T10 (tag_value, created) values(" & tag_date_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T10 = Nothing
Dim tag_date_T11
tag_date_T11 = HMIRuntime.Tags("date_T11").Read
strSQL = "INSERT INTO z_tag_date_T11 (tag_value, created) values(" & tag_date_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T11 = Nothing
Dim tag_date_T12
tag_date_T12 = HMIRuntime.Tags("date_T12").Read
strSQL = "INSERT INTO z_tag_date_T12 (tag_value, created) values(" & tag_date_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T12 = Nothing
Dim tag_date_T13
tag_date_T13 = HMIRuntime.Tags("date_T13").Read
strSQL = "INSERT INTO z_tag_date_T13 (tag_value, created) values(" & tag_date_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T13 = Nothing
Dim tag_date_T14
tag_date_T14 = HMIRuntime.Tags("date_T14").Read
strSQL = "INSERT INTO z_tag_date_T14 (tag_value, created) values(" & tag_date_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T14 = Nothing
Dim tag_date_T15
tag_date_T15 = HMIRuntime.Tags("date_T15").Read
strSQL = "INSERT INTO z_tag_date_T15 (tag_value, created) values(" & tag_date_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T15 = Nothing
Dim tag_date_T16
tag_date_T16 = HMIRuntime.Tags("date_T16").Read
strSQL = "INSERT INTO z_tag_date_T16 (tag_value, created) values(" & tag_date_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T16 = Nothing
Dim tag_date_T17
tag_date_T17 = HMIRuntime.Tags("date_T17").Read
strSQL = "INSERT INTO z_tag_date_T17 (tag_value, created) values(" & tag_date_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T17 = Nothing
Dim tag_date_T18
tag_date_T18 = HMIRuntime.Tags("date_T18").Read
strSQL = "INSERT INTO z_tag_date_T18 (tag_value, created) values(" & tag_date_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T18 = Nothing
Dim tag_date_T19
tag_date_T19 = HMIRuntime.Tags("date_T19").Read
strSQL = "INSERT INTO z_tag_date_T19 (tag_value, created) values(" & tag_date_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T19 = Nothing
Dim tag_date_T20
tag_date_T20 = HMIRuntime.Tags("date_T20").Read
strSQL = "INSERT INTO z_tag_date_T20 (tag_value, created) values(" & tag_date_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T20 = Nothing
Dim tag_date_T21
tag_date_T21 = HMIRuntime.Tags("date_T21").Read
strSQL = "INSERT INTO z_tag_date_T21 (tag_value, created) values(" & tag_date_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T21 = Nothing
Dim tag_date_T22
tag_date_T22 = HMIRuntime.Tags("date_T22").Read
strSQL = "INSERT INTO z_tag_date_T22 (tag_value, created) values(" & tag_date_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T22 = Nothing
Dim tag_date_T23
tag_date_T23 = HMIRuntime.Tags("date_T23").Read
strSQL = "INSERT INTO z_tag_date_T23 (tag_value, created) values(" & tag_date_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T23 = Nothing
Dim tag_date_T24
tag_date_T24 = HMIRuntime.Tags("date_T24").Read
strSQL = "INSERT INTO z_tag_date_T24 (tag_value, created) values(" & tag_date_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T24 = Nothing
Dim tag_date_T25
tag_date_T25 = HMIRuntime.Tags("date_T25").Read
strSQL = "INSERT INTO z_tag_date_T25 (tag_value, created) values(" & tag_date_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T25 = Nothing
Dim tag_date_T26
tag_date_T26 = HMIRuntime.Tags("date_T26").Read
strSQL = "INSERT INTO z_tag_date_T26 (tag_value, created) values(" & tag_date_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T26 = Nothing
Dim tag_date_T27
tag_date_T27 = HMIRuntime.Tags("date_T27").Read
strSQL = "INSERT INTO z_tag_date_T27 (tag_value, created) values(" & tag_date_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T27 = Nothing
Dim tag_date_T28
tag_date_T28 = HMIRuntime.Tags("date_T28").Read
strSQL = "INSERT INTO z_tag_date_T28 (tag_value, created) values(" & tag_date_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T28 = Nothing
Dim tag_date_T29
tag_date_T29 = HMIRuntime.Tags("date_T29").Read
strSQL = "INSERT INTO z_tag_date_T29 (tag_value, created) values(" & tag_date_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T29 = Nothing
Dim tag_date_T30
tag_date_T30 = HMIRuntime.Tags("date_T30").Read
strSQL = "INSERT INTO z_tag_date_T30 (tag_value, created) values(" & tag_date_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date_T30 = Nothing
Dim tag_date2_T1
tag_date2_T1 = HMIRuntime.Tags("date2_T1").Read
strSQL = "INSERT INTO z_tag_date2_T1 (tag_value, created) values(" & tag_date2_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T1 = Nothing
Dim tag_date2_T2
tag_date2_T2 = HMIRuntime.Tags("date2_T2").Read
strSQL = "INSERT INTO z_tag_date2_T2 (tag_value, created) values(" & tag_date2_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T2 = Nothing
Dim tag_date2_T3
tag_date2_T3 = HMIRuntime.Tags("date2_T3").Read
strSQL = "INSERT INTO z_tag_date2_T3 (tag_value, created) values(" & tag_date2_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T3 = Nothing
Dim tag_date2_T4
tag_date2_T4 = HMIRuntime.Tags("date2_T4").Read
strSQL = "INSERT INTO z_tag_date2_T4 (tag_value, created) values(" & tag_date2_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T4 = Nothing
Dim tag_date2_T5
tag_date2_T5 = HMIRuntime.Tags("date2_T5").Read
strSQL = "INSERT INTO z_tag_date2_T5 (tag_value, created) values(" & tag_date2_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T5 = Nothing
Dim tag_date2_T6
tag_date2_T6 = HMIRuntime.Tags("date2_T6").Read
strSQL = "INSERT INTO z_tag_date2_T6 (tag_value, created) values(" & tag_date2_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T6 = Nothing
Dim tag_date2_T7
tag_date2_T7 = HMIRuntime.Tags("date2_T7").Read
strSQL = "INSERT INTO z_tag_date2_T7 (tag_value, created) values(" & tag_date2_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T7 = Nothing
Dim tag_date2_T8
tag_date2_T8 = HMIRuntime.Tags("date2_T8").Read
strSQL = "INSERT INTO z_tag_date2_T8 (tag_value, created) values(" & tag_date2_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T8 = Nothing
Dim tag_date2_T9
tag_date2_T9 = HMIRuntime.Tags("date2_T9").Read
strSQL = "INSERT INTO z_tag_date2_T9 (tag_value, created) values(" & tag_date2_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T9 = Nothing
Dim tag_date2_T10
tag_date2_T10 = HMIRuntime.Tags("date2_T10").Read
strSQL = "INSERT INTO z_tag_date2_T10 (tag_value, created) values(" & tag_date2_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T10 = Nothing
Dim tag_date2_T11
tag_date2_T11 = HMIRuntime.Tags("date2_T11").Read
strSQL = "INSERT INTO z_tag_date2_T11 (tag_value, created) values(" & tag_date2_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T11 = Nothing
Dim tag_date2_T12
tag_date2_T12 = HMIRuntime.Tags("date2_T12").Read
strSQL = "INSERT INTO z_tag_date2_T12 (tag_value, created) values(" & tag_date2_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T12 = Nothing
Dim tag_date2_T13
tag_date2_T13 = HMIRuntime.Tags("date2_T13").Read
strSQL = "INSERT INTO z_tag_date2_T13 (tag_value, created) values(" & tag_date2_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T13 = Nothing
Dim tag_date2_T14
tag_date2_T14 = HMIRuntime.Tags("date2_T14").Read
strSQL = "INSERT INTO z_tag_date2_T14 (tag_value, created) values(" & tag_date2_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T14 = Nothing
Dim tag_date2_T15
tag_date2_T15 = HMIRuntime.Tags("date2_T15").Read
strSQL = "INSERT INTO z_tag_date2_T15 (tag_value, created) values(" & tag_date2_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T15 = Nothing
Dim tag_date2_T16
tag_date2_T16 = HMIRuntime.Tags("date2_T16").Read
strSQL = "INSERT INTO z_tag_date2_T16 (tag_value, created) values(" & tag_date2_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T16 = Nothing
Dim tag_date2_T17
tag_date2_T17 = HMIRuntime.Tags("date2_T17").Read
strSQL = "INSERT INTO z_tag_date2_T17 (tag_value, created) values(" & tag_date2_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T17 = Nothing
Dim tag_date2_T18
tag_date2_T18 = HMIRuntime.Tags("date2_T18").Read
strSQL = "INSERT INTO z_tag_date2_T18 (tag_value, created) values(" & tag_date2_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T18 = Nothing
Dim tag_date2_T19
tag_date2_T19 = HMIRuntime.Tags("date2_T19").Read
strSQL = "INSERT INTO z_tag_date2_T19 (tag_value, created) values(" & tag_date2_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T19 = Nothing
Dim tag_date2_T20
tag_date2_T20 = HMIRuntime.Tags("date2_T20").Read
strSQL = "INSERT INTO z_tag_date2_T20 (tag_value, created) values(" & tag_date2_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T20 = Nothing
Dim tag_date2_T21
tag_date2_T21 = HMIRuntime.Tags("date2_T21").Read
strSQL = "INSERT INTO z_tag_date2_T21 (tag_value, created) values(" & tag_date2_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T21 = Nothing
Dim tag_date2_T22
tag_date2_T22 = HMIRuntime.Tags("date2_T22").Read
strSQL = "INSERT INTO z_tag_date2_T22 (tag_value, created) values(" & tag_date2_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T22 = Nothing
Dim tag_date2_T23
tag_date2_T23 = HMIRuntime.Tags("date2_T23").Read
strSQL = "INSERT INTO z_tag_date2_T23 (tag_value, created) values(" & tag_date2_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T23 = Nothing
Dim tag_date2_T24
tag_date2_T24 = HMIRuntime.Tags("date2_T24").Read
strSQL = "INSERT INTO z_tag_date2_T24 (tag_value, created) values(" & tag_date2_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T24 = Nothing
Dim tag_date2_T25
tag_date2_T25 = HMIRuntime.Tags("date2_T25").Read
strSQL = "INSERT INTO z_tag_date2_T25 (tag_value, created) values(" & tag_date2_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T25 = Nothing
Dim tag_date2_T26
tag_date2_T26 = HMIRuntime.Tags("date2_T26").Read
strSQL = "INSERT INTO z_tag_date2_T26 (tag_value, created) values(" & tag_date2_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T26 = Nothing
Dim tag_date2_T27
tag_date2_T27 = HMIRuntime.Tags("date2_T27").Read
strSQL = "INSERT INTO z_tag_date2_T27 (tag_value, created) values(" & tag_date2_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T27 = Nothing
Dim tag_date2_T28
tag_date2_T28 = HMIRuntime.Tags("date2_T28").Read
strSQL = "INSERT INTO z_tag_date2_T28 (tag_value, created) values(" & tag_date2_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T28 = Nothing
Dim tag_date2_T29
tag_date2_T29 = HMIRuntime.Tags("date2_T29").Read
strSQL = "INSERT INTO z_tag_date2_T29 (tag_value, created) values(" & tag_date2_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T29 = Nothing
Dim tag_date2_T30
tag_date2_T30 = HMIRuntime.Tags("date2_T30").Read
strSQL = "INSERT INTO z_tag_date2_T30 (tag_value, created) values(" & tag_date2_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_date2_T30 = Nothing
Dim tag_flowmeternew
tag_flowmeternew = HMIRuntime.Tags("flowmeternew").Read
strSQL = "INSERT INTO z_tag_flowmeternew (tag_value, created) values(" & tag_flowmeternew & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_flowmeternew = Nothing
Dim tag_GIOCAODIEM
tag_GIOCAODIEM = HMIRuntime.Tags("GIOCAODIEM").Read
strSQL = "INSERT INTO z_tag_GIOCAODIEM (tag_value, created) values(" & tag_GIOCAODIEM & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_GIOCAODIEM = Nothing
Dim tag_ICM_DA01
tag_ICM_DA01 = HMIRuntime.Tags("ICM_DA01").Read
strSQL = "INSERT INTO z_tag_ICM_DA01 (tag_value, created) values(" & tag_ICM_DA01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ICM_DA01 = Nothing
Dim tag_ICM_DA02
tag_ICM_DA02 = HMIRuntime.Tags("ICM_DA02").Read
strSQL = "INSERT INTO z_tag_ICM_DA02 (tag_value, created) values(" & tag_ICM_DA02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ICM_DA02 = Nothing
Dim tag_luuluongthuhoimen
tag_luuluongthuhoimen = HMIRuntime.Tags("luuluongthuhoimen").Read
strSQL = "INSERT INTO z_tag_luuluongthuhoimen (tag_value, created) values(" & tag_luuluongthuhoimen & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_luuluongthuhoimen = Nothing
Dim tag_MakeupCipPre_F_run
tag_MakeupCipPre_F_run = HMIRuntime.Tags("MakeupCipPre_F_run").Read
strSQL = "INSERT INTO z_tag_MakeupCipPre_F_run (tag_value, created) values(" & tag_MakeupCipPre_F_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MakeupCipPre_F_run = Nothing
Dim tag_MakeupCipPre_F_seq
tag_MakeupCipPre_F_seq = HMIRuntime.Tags("MakeupCipPre_F_seq").Read
strSQL = "INSERT INTO z_tag_MakeupCipPre_F_seq (tag_value, created) values(" & tag_MakeupCipPre_F_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MakeupCipPre_F_seq = Nothing
Dim tag_matdotbm_T1
tag_matdotbm_T1 = HMIRuntime.Tags("matdotbm_T1").Read
strSQL = "INSERT INTO z_tag_matdotbm_T1 (tag_value, created) values(" & tag_matdotbm_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T1 = Nothing
Dim tag_matdotbm_T2
tag_matdotbm_T2 = HMIRuntime.Tags("matdotbm_T2").Read
strSQL = "INSERT INTO z_tag_matdotbm_T2 (tag_value, created) values(" & tag_matdotbm_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T2 = Nothing
Dim tag_matdotbm_T3
tag_matdotbm_T3 = HMIRuntime.Tags("matdotbm_T3").Read
strSQL = "INSERT INTO z_tag_matdotbm_T3 (tag_value, created) values(" & tag_matdotbm_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T3 = Nothing
Dim tag_matdotbm_T4
tag_matdotbm_T4 = HMIRuntime.Tags("matdotbm_T4").Read
strSQL = "INSERT INTO z_tag_matdotbm_T4 (tag_value, created) values(" & tag_matdotbm_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T4 = Nothing
Dim tag_matdotbm_T5
tag_matdotbm_T5 = HMIRuntime.Tags("matdotbm_T5").Read
strSQL = "INSERT INTO z_tag_matdotbm_T5 (tag_value, created) values(" & tag_matdotbm_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T5 = Nothing
Dim tag_matdotbm_T6
tag_matdotbm_T6 = HMIRuntime.Tags("matdotbm_T6").Read
strSQL = "INSERT INTO z_tag_matdotbm_T6 (tag_value, created) values(" & tag_matdotbm_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T6 = Nothing
Dim tag_matdotbm_T7
tag_matdotbm_T7 = HMIRuntime.Tags("matdotbm_T7").Read
strSQL = "INSERT INTO z_tag_matdotbm_T7 (tag_value, created) values(" & tag_matdotbm_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T7 = Nothing
Dim tag_matdotbm_T8
tag_matdotbm_T8 = HMIRuntime.Tags("matdotbm_T8").Read
strSQL = "INSERT INTO z_tag_matdotbm_T8 (tag_value, created) values(" & tag_matdotbm_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T8 = Nothing
Dim tag_matdotbm_T9
tag_matdotbm_T9 = HMIRuntime.Tags("matdotbm_T9").Read
strSQL = "INSERT INTO z_tag_matdotbm_T9 (tag_value, created) values(" & tag_matdotbm_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T9 = Nothing
Dim tag_matdotbm_T10
tag_matdotbm_T10 = HMIRuntime.Tags("matdotbm_T10").Read
strSQL = "INSERT INTO z_tag_matdotbm_T10 (tag_value, created) values(" & tag_matdotbm_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T10 = Nothing
Dim tag_matdotbm_T11
tag_matdotbm_T11 = HMIRuntime.Tags("matdotbm_T11").Read
strSQL = "INSERT INTO z_tag_matdotbm_T11 (tag_value, created) values(" & tag_matdotbm_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T11 = Nothing
Dim tag_matdotbm_T12
tag_matdotbm_T12 = HMIRuntime.Tags("matdotbm_T12").Read
strSQL = "INSERT INTO z_tag_matdotbm_T12 (tag_value, created) values(" & tag_matdotbm_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T12 = Nothing
Dim tag_matdotbm_T13
tag_matdotbm_T13 = HMIRuntime.Tags("matdotbm_T13").Read
strSQL = "INSERT INTO z_tag_matdotbm_T13 (tag_value, created) values(" & tag_matdotbm_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T13 = Nothing
Dim tag_matdotbm_T14
tag_matdotbm_T14 = HMIRuntime.Tags("matdotbm_T14").Read
strSQL = "INSERT INTO z_tag_matdotbm_T14 (tag_value, created) values(" & tag_matdotbm_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T14 = Nothing
Dim tag_matdotbm_T15
tag_matdotbm_T15 = HMIRuntime.Tags("matdotbm_T15").Read
strSQL = "INSERT INTO z_tag_matdotbm_T15 (tag_value, created) values(" & tag_matdotbm_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T15 = Nothing
Dim tag_matdotbm_T16
tag_matdotbm_T16 = HMIRuntime.Tags("matdotbm_T16").Read
strSQL = "INSERT INTO z_tag_matdotbm_T16 (tag_value, created) values(" & tag_matdotbm_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T16 = Nothing
Dim tag_matdotbm_T17
tag_matdotbm_T17 = HMIRuntime.Tags("matdotbm_T17").Read
strSQL = "INSERT INTO z_tag_matdotbm_T17 (tag_value, created) values(" & tag_matdotbm_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T17 = Nothing
Dim tag_matdotbm_T18
tag_matdotbm_T18 = HMIRuntime.Tags("matdotbm_T18").Read
strSQL = "INSERT INTO z_tag_matdotbm_T18 (tag_value, created) values(" & tag_matdotbm_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T18 = Nothing
Dim tag_matdotbm_T19
tag_matdotbm_T19 = HMIRuntime.Tags("matdotbm_T19").Read
strSQL = "INSERT INTO z_tag_matdotbm_T19 (tag_value, created) values(" & tag_matdotbm_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T19 = Nothing
Dim tag_matdotbm_T20
tag_matdotbm_T20 = HMIRuntime.Tags("matdotbm_T20").Read
strSQL = "INSERT INTO z_tag_matdotbm_T20 (tag_value, created) values(" & tag_matdotbm_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T20 = Nothing
Dim tag_matdotbm_T21
tag_matdotbm_T21 = HMIRuntime.Tags("matdotbm_T21").Read
strSQL = "INSERT INTO z_tag_matdotbm_T21 (tag_value, created) values(" & tag_matdotbm_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T21 = Nothing
Dim tag_matdotbm_T22
tag_matdotbm_T22 = HMIRuntime.Tags("matdotbm_T22").Read
strSQL = "INSERT INTO z_tag_matdotbm_T22 (tag_value, created) values(" & tag_matdotbm_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T22 = Nothing
Dim tag_matdotbm_T23
tag_matdotbm_T23 = HMIRuntime.Tags("matdotbm_T23").Read
strSQL = "INSERT INTO z_tag_matdotbm_T23 (tag_value, created) values(" & tag_matdotbm_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T23 = Nothing
Dim tag_matdotbm_T24
tag_matdotbm_T24 = HMIRuntime.Tags("matdotbm_T24").Read
strSQL = "INSERT INTO z_tag_matdotbm_T24 (tag_value, created) values(" & tag_matdotbm_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T24 = Nothing
Dim tag_matdotbm_T25
tag_matdotbm_T25 = HMIRuntime.Tags("matdotbm_T25").Read
strSQL = "INSERT INTO z_tag_matdotbm_T25 (tag_value, created) values(" & tag_matdotbm_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T25 = Nothing
Dim tag_matdotbm_T26
tag_matdotbm_T26 = HMIRuntime.Tags("matdotbm_T26").Read
strSQL = "INSERT INTO z_tag_matdotbm_T26 (tag_value, created) values(" & tag_matdotbm_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T26 = Nothing
Dim tag_matdotbm_T27
tag_matdotbm_T27 = HMIRuntime.Tags("matdotbm_T27").Read
strSQL = "INSERT INTO z_tag_matdotbm_T27 (tag_value, created) values(" & tag_matdotbm_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T27 = Nothing
Dim tag_matdotbm_T28
tag_matdotbm_T28 = HMIRuntime.Tags("matdotbm_T28").Read
strSQL = "INSERT INTO z_tag_matdotbm_T28 (tag_value, created) values(" & tag_matdotbm_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T28 = Nothing
Dim tag_matdotbm_T29
tag_matdotbm_T29 = HMIRuntime.Tags("matdotbm_T29").Read
strSQL = "INSERT INTO z_tag_matdotbm_T29 (tag_value, created) values(" & tag_matdotbm_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T29 = Nothing
Dim tag_matdotbm_T30
tag_matdotbm_T30 = HMIRuntime.Tags("matdotbm_T30").Read
strSQL = "INSERT INTO z_tag_matdotbm_T30 (tag_value, created) values(" & tag_matdotbm_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_matdotbm_T30 = Nothing
Dim tag_month_T1
tag_month_T1 = HMIRuntime.Tags("month_T1").Read
strSQL = "INSERT INTO z_tag_month_T1 (tag_value, created) values(" & tag_month_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T1 = Nothing
Dim tag_month_T2
tag_month_T2 = HMIRuntime.Tags("month_T2").Read
strSQL = "INSERT INTO z_tag_month_T2 (tag_value, created) values(" & tag_month_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T2 = Nothing
Dim tag_month_T3
tag_month_T3 = HMIRuntime.Tags("month_T3").Read
strSQL = "INSERT INTO z_tag_month_T3 (tag_value, created) values(" & tag_month_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T3 = Nothing
Dim tag_month_T4
tag_month_T4 = HMIRuntime.Tags("month_T4").Read
strSQL = "INSERT INTO z_tag_month_T4 (tag_value, created) values(" & tag_month_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T4 = Nothing
Dim tag_month_T5
tag_month_T5 = HMIRuntime.Tags("month_T5").Read
strSQL = "INSERT INTO z_tag_month_T5 (tag_value, created) values(" & tag_month_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T5 = Nothing
Dim tag_month_T6
tag_month_T6 = HMIRuntime.Tags("month_T6").Read
strSQL = "INSERT INTO z_tag_month_T6 (tag_value, created) values(" & tag_month_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T6 = Nothing
Dim tag_month_T7
tag_month_T7 = HMIRuntime.Tags("month_T7").Read
strSQL = "INSERT INTO z_tag_month_T7 (tag_value, created) values(" & tag_month_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T7 = Nothing
Dim tag_month_T8
tag_month_T8 = HMIRuntime.Tags("month_T8").Read
strSQL = "INSERT INTO z_tag_month_T8 (tag_value, created) values(" & tag_month_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T8 = Nothing
Dim tag_month_T9
tag_month_T9 = HMIRuntime.Tags("month_T9").Read
strSQL = "INSERT INTO z_tag_month_T9 (tag_value, created) values(" & tag_month_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T9 = Nothing
Dim tag_month_T10
tag_month_T10 = HMIRuntime.Tags("month_T10").Read
strSQL = "INSERT INTO z_tag_month_T10 (tag_value, created) values(" & tag_month_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T10 = Nothing
Dim tag_month_T11
tag_month_T11 = HMIRuntime.Tags("month_T11").Read
strSQL = "INSERT INTO z_tag_month_T11 (tag_value, created) values(" & tag_month_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T11 = Nothing
Dim tag_month_T12
tag_month_T12 = HMIRuntime.Tags("month_T12").Read
strSQL = "INSERT INTO z_tag_month_T12 (tag_value, created) values(" & tag_month_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T12 = Nothing
Dim tag_month_T13
tag_month_T13 = HMIRuntime.Tags("month_T13").Read
strSQL = "INSERT INTO z_tag_month_T13 (tag_value, created) values(" & tag_month_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T13 = Nothing
Dim tag_month_T14
tag_month_T14 = HMIRuntime.Tags("month_T14").Read
strSQL = "INSERT INTO z_tag_month_T14 (tag_value, created) values(" & tag_month_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T14 = Nothing
Dim tag_month_T15
tag_month_T15 = HMIRuntime.Tags("month_T15").Read
strSQL = "INSERT INTO z_tag_month_T15 (tag_value, created) values(" & tag_month_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T15 = Nothing
Dim tag_month_T16
tag_month_T16 = HMIRuntime.Tags("month_T16").Read
strSQL = "INSERT INTO z_tag_month_T16 (tag_value, created) values(" & tag_month_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T16 = Nothing
Dim tag_month_T17
tag_month_T17 = HMIRuntime.Tags("month_T17").Read
strSQL = "INSERT INTO z_tag_month_T17 (tag_value, created) values(" & tag_month_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T17 = Nothing
Dim tag_month_T18
tag_month_T18 = HMIRuntime.Tags("month_T18").Read
strSQL = "INSERT INTO z_tag_month_T18 (tag_value, created) values(" & tag_month_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T18 = Nothing
Dim tag_month_T19
tag_month_T19 = HMIRuntime.Tags("month_T19").Read
strSQL = "INSERT INTO z_tag_month_T19 (tag_value, created) values(" & tag_month_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T19 = Nothing
Dim tag_month_T20
tag_month_T20 = HMIRuntime.Tags("month_T20").Read
strSQL = "INSERT INTO z_tag_month_T20 (tag_value, created) values(" & tag_month_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T20 = Nothing
Dim tag_month_T21
tag_month_T21 = HMIRuntime.Tags("month_T21").Read
strSQL = "INSERT INTO z_tag_month_T21 (tag_value, created) values(" & tag_month_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T21 = Nothing
Dim tag_month_T22
tag_month_T22 = HMIRuntime.Tags("month_T22").Read
strSQL = "INSERT INTO z_tag_month_T22 (tag_value, created) values(" & tag_month_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T22 = Nothing
Dim tag_month_T23
tag_month_T23 = HMIRuntime.Tags("month_T23").Read
strSQL = "INSERT INTO z_tag_month_T23 (tag_value, created) values(" & tag_month_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T23 = Nothing
Dim tag_month_T24
tag_month_T24 = HMIRuntime.Tags("month_T24").Read
strSQL = "INSERT INTO z_tag_month_T24 (tag_value, created) values(" & tag_month_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T24 = Nothing
Dim tag_month_T25
tag_month_T25 = HMIRuntime.Tags("month_T25").Read
strSQL = "INSERT INTO z_tag_month_T25 (tag_value, created) values(" & tag_month_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T25 = Nothing
Dim tag_month_T26
tag_month_T26 = HMIRuntime.Tags("month_T26").Read
strSQL = "INSERT INTO z_tag_month_T26 (tag_value, created) values(" & tag_month_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T26 = Nothing
Dim tag_month_T27
tag_month_T27 = HMIRuntime.Tags("month_T27").Read
strSQL = "INSERT INTO z_tag_month_T27 (tag_value, created) values(" & tag_month_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T27 = Nothing
Dim tag_month_T28
tag_month_T28 = HMIRuntime.Tags("month_T28").Read
strSQL = "INSERT INTO z_tag_month_T28 (tag_value, created) values(" & tag_month_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T28 = Nothing
Dim tag_month_T29
tag_month_T29 = HMIRuntime.Tags("month_T29").Read
strSQL = "INSERT INTO z_tag_month_T29 (tag_value, created) values(" & tag_month_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T29 = Nothing
Dim tag_month_T30
tag_month_T30 = HMIRuntime.Tags("month_T30").Read
strSQL = "INSERT INTO z_tag_month_T30 (tag_value, created) values(" & tag_month_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month_T30 = Nothing
Dim tag_month2_T1
tag_month2_T1 = HMIRuntime.Tags("month2_T1").Read
strSQL = "INSERT INTO z_tag_month2_T1 (tag_value, created) values(" & tag_month2_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T1 = Nothing
Dim tag_month2_T2
tag_month2_T2 = HMIRuntime.Tags("month2_T2").Read
strSQL = "INSERT INTO z_tag_month2_T2 (tag_value, created) values(" & tag_month2_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T2 = Nothing
Dim tag_month2_T3
tag_month2_T3 = HMIRuntime.Tags("month2_T3").Read
strSQL = "INSERT INTO z_tag_month2_T3 (tag_value, created) values(" & tag_month2_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T3 = Nothing
Dim tag_month2_T4
tag_month2_T4 = HMIRuntime.Tags("month2_T4").Read
strSQL = "INSERT INTO z_tag_month2_T4 (tag_value, created) values(" & tag_month2_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T4 = Nothing
Dim tag_month2_T5
tag_month2_T5 = HMIRuntime.Tags("month2_T5").Read
strSQL = "INSERT INTO z_tag_month2_T5 (tag_value, created) values(" & tag_month2_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T5 = Nothing
Dim tag_month2_T6
tag_month2_T6 = HMIRuntime.Tags("month2_T6").Read
strSQL = "INSERT INTO z_tag_month2_T6 (tag_value, created) values(" & tag_month2_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T6 = Nothing
Dim tag_month2_T7
tag_month2_T7 = HMIRuntime.Tags("month2_T7").Read
strSQL = "INSERT INTO z_tag_month2_T7 (tag_value, created) values(" & tag_month2_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T7 = Nothing
Dim tag_month2_T8
tag_month2_T8 = HMIRuntime.Tags("month2_T8").Read
strSQL = "INSERT INTO z_tag_month2_T8 (tag_value, created) values(" & tag_month2_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T8 = Nothing
Dim tag_month2_T9
tag_month2_T9 = HMIRuntime.Tags("month2_T9").Read
strSQL = "INSERT INTO z_tag_month2_T9 (tag_value, created) values(" & tag_month2_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T9 = Nothing
Dim tag_month2_T10
tag_month2_T10 = HMIRuntime.Tags("month2_T10").Read
strSQL = "INSERT INTO z_tag_month2_T10 (tag_value, created) values(" & tag_month2_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T10 = Nothing
Dim tag_month2_T11
tag_month2_T11 = HMIRuntime.Tags("month2_T11").Read
strSQL = "INSERT INTO z_tag_month2_T11 (tag_value, created) values(" & tag_month2_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T11 = Nothing
Dim tag_month2_T12
tag_month2_T12 = HMIRuntime.Tags("month2_T12").Read
strSQL = "INSERT INTO z_tag_month2_T12 (tag_value, created) values(" & tag_month2_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T12 = Nothing
Dim tag_month2_T13
tag_month2_T13 = HMIRuntime.Tags("month2_T13").Read
strSQL = "INSERT INTO z_tag_month2_T13 (tag_value, created) values(" & tag_month2_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T13 = Nothing
Dim tag_month2_T14
tag_month2_T14 = HMIRuntime.Tags("month2_T14").Read
strSQL = "INSERT INTO z_tag_month2_T14 (tag_value, created) values(" & tag_month2_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T14 = Nothing
Dim tag_month2_T15
tag_month2_T15 = HMIRuntime.Tags("month2_T15").Read
strSQL = "INSERT INTO z_tag_month2_T15 (tag_value, created) values(" & tag_month2_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T15 = Nothing
Dim tag_month2_T16
tag_month2_T16 = HMIRuntime.Tags("month2_T16").Read
strSQL = "INSERT INTO z_tag_month2_T16 (tag_value, created) values(" & tag_month2_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T16 = Nothing
Dim tag_month2_T17
tag_month2_T17 = HMIRuntime.Tags("month2_T17").Read
strSQL = "INSERT INTO z_tag_month2_T17 (tag_value, created) values(" & tag_month2_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T17 = Nothing
Dim tag_month2_T18
tag_month2_T18 = HMIRuntime.Tags("month2_T18").Read
strSQL = "INSERT INTO z_tag_month2_T18 (tag_value, created) values(" & tag_month2_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T18 = Nothing
Dim tag_month2_T19
tag_month2_T19 = HMIRuntime.Tags("month2_T19").Read
strSQL = "INSERT INTO z_tag_month2_T19 (tag_value, created) values(" & tag_month2_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T19 = Nothing
Dim tag_month2_T20
tag_month2_T20 = HMIRuntime.Tags("month2_T20").Read
strSQL = "INSERT INTO z_tag_month2_T20 (tag_value, created) values(" & tag_month2_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T20 = Nothing
Dim tag_month2_T21
tag_month2_T21 = HMIRuntime.Tags("month2_T21").Read
strSQL = "INSERT INTO z_tag_month2_T21 (tag_value, created) values(" & tag_month2_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T21 = Nothing
Dim tag_month2_T22
tag_month2_T22 = HMIRuntime.Tags("month2_T22").Read
strSQL = "INSERT INTO z_tag_month2_T22 (tag_value, created) values(" & tag_month2_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T22 = Nothing
Dim tag_month2_T23
tag_month2_T23 = HMIRuntime.Tags("month2_T23").Read
strSQL = "INSERT INTO z_tag_month2_T23 (tag_value, created) values(" & tag_month2_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T23 = Nothing
Dim tag_month2_T24
tag_month2_T24 = HMIRuntime.Tags("month2_T24").Read
strSQL = "INSERT INTO z_tag_month2_T24 (tag_value, created) values(" & tag_month2_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T24 = Nothing
Dim tag_month2_T25
tag_month2_T25 = HMIRuntime.Tags("month2_T25").Read
strSQL = "INSERT INTO z_tag_month2_T25 (tag_value, created) values(" & tag_month2_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T25 = Nothing
Dim tag_month2_T26
tag_month2_T26 = HMIRuntime.Tags("month2_T26").Read
strSQL = "INSERT INTO z_tag_month2_T26 (tag_value, created) values(" & tag_month2_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T26 = Nothing
Dim tag_month2_T27
tag_month2_T27 = HMIRuntime.Tags("month2_T27").Read
strSQL = "INSERT INTO z_tag_month2_T27 (tag_value, created) values(" & tag_month2_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T27 = Nothing
Dim tag_month2_T28
tag_month2_T28 = HMIRuntime.Tags("month2_T28").Read
strSQL = "INSERT INTO z_tag_month2_T28 (tag_value, created) values(" & tag_month2_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T28 = Nothing
Dim tag_month2_T29
tag_month2_T29 = HMIRuntime.Tags("month2_T29").Read
strSQL = "INSERT INTO z_tag_month2_T29 (tag_value, created) values(" & tag_month2_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T29 = Nothing
Dim tag_month2_T30
tag_month2_T30 = HMIRuntime.Tags("month2_T30").Read
strSQL = "INSERT INTO z_tag_month2_T30 (tag_value, created) values(" & tag_month2_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_month2_T30 = Nothing
Dim tag_Mucthap_TBF01
tag_Mucthap_TBF01 = HMIRuntime.Tags("Mucthap_TBF01").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF01 (tag_value, created) values(" & tag_Mucthap_TBF01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF01 = Nothing
Dim tag_Mucthap_TBF02
tag_Mucthap_TBF02 = HMIRuntime.Tags("Mucthap_TBF02").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF02 (tag_value, created) values(" & tag_Mucthap_TBF02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF02 = Nothing
Dim tag_Mucthap_TBF03
tag_Mucthap_TBF03 = HMIRuntime.Tags("Mucthap_TBF03").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF03 (tag_value, created) values(" & tag_Mucthap_TBF03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF03 = Nothing
Dim tag_Mucthap_TBF04
tag_Mucthap_TBF04 = HMIRuntime.Tags("Mucthap_TBF04").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF04 (tag_value, created) values(" & tag_Mucthap_TBF04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF04 = Nothing
Dim tag_Mucthap_TBF07
tag_Mucthap_TBF07 = HMIRuntime.Tags("Mucthap_TBF07").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF07 (tag_value, created) values(" & tag_Mucthap_TBF07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF07 = Nothing
Dim tag_Mucthap_TBF08
tag_Mucthap_TBF08 = HMIRuntime.Tags("Mucthap_TBF08").Read
strSQL = "INSERT INTO z_tag_Mucthap_TBF08 (tag_value, created) values(" & tag_Mucthap_TBF08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Mucthap_TBF08 = Nothing
Dim tag_MUpCipFiller_run
tag_MUpCipFiller_run = HMIRuntime.Tags("MUpCipFiller_run").Read
strSQL = "INSERT INTO z_tag_MUpCipFiller_run (tag_value, created) values(" & tag_MUpCipFiller_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpCipFiller_run = Nothing
Dim tag_MUpCipFiller_seq
tag_MUpCipFiller_seq = HMIRuntime.Tags("MUpCipFiller_seq").Read
strSQL = "INSERT INTO z_tag_MUpCipFiller_seq (tag_value, created) values(" & tag_MUpCipFiller_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpCipFiller_seq = Nothing
Dim tag_MUpCipFilter_run
tag_MUpCipFilter_run = HMIRuntime.Tags("MUpCipFilter_run").Read
strSQL = "INSERT INTO z_tag_MUpCipFilter_run (tag_value, created) values(" & tag_MUpCipFilter_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpCipFilter_run = Nothing
Dim tag_MUpCipFilter_seq
tag_MUpCipFilter_seq = HMIRuntime.Tags("MUpCipFilter_seq").Read
strSQL = "INSERT INTO z_tag_MUpCipFilter_seq (tag_value, created) values(" & tag_MUpCipFilter_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpCipFilter_seq = Nothing
Dim tag_MUpFilter_run
tag_MUpFilter_run = HMIRuntime.Tags("MUpFilter_run").Read
strSQL = "INSERT INTO z_tag_MUpFilter_run (tag_value, created) values(" & tag_MUpFilter_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpFilter_run = Nothing
Dim tag_MUpFilter_seq
tag_MUpFilter_seq = HMIRuntime.Tags("MUpFilter_seq").Read
strSQL = "INSERT INTO z_tag_MUpFilter_seq (tag_value, created) values(" & tag_MUpFilter_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_MUpFilter_seq = Nothing
Dim tag_nhietdolmc_T1
tag_nhietdolmc_T1 = HMIRuntime.Tags("nhietdolmc_T1").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T1 (tag_value, created) values(" & tag_nhietdolmc_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T1 = Nothing
Dim tag_nhietdolmc_T2
tag_nhietdolmc_T2 = HMIRuntime.Tags("nhietdolmc_T2").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T2 (tag_value, created) values(" & tag_nhietdolmc_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T2 = Nothing
Dim tag_nhietdolmc_T3
tag_nhietdolmc_T3 = HMIRuntime.Tags("nhietdolmc_T3").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T3 (tag_value, created) values(" & tag_nhietdolmc_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T3 = Nothing
Dim tag_nhietdolmc_T4
tag_nhietdolmc_T4 = HMIRuntime.Tags("nhietdolmc_T4").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T4 (tag_value, created) values(" & tag_nhietdolmc_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T4 = Nothing
Dim tag_nhietdolmc_T5
tag_nhietdolmc_T5 = HMIRuntime.Tags("nhietdolmc_T5").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T5 (tag_value, created) values(" & tag_nhietdolmc_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T5 = Nothing
Dim tag_nhietdolmc_T6
tag_nhietdolmc_T6 = HMIRuntime.Tags("nhietdolmc_T6").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T6 (tag_value, created) values(" & tag_nhietdolmc_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T6 = Nothing
Dim tag_nhietdolmc_T7
tag_nhietdolmc_T7 = HMIRuntime.Tags("nhietdolmc_T7").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T7 (tag_value, created) values(" & tag_nhietdolmc_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T7 = Nothing
Dim tag_nhietdolmc_T8
tag_nhietdolmc_T8 = HMIRuntime.Tags("nhietdolmc_T8").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T8 (tag_value, created) values(" & tag_nhietdolmc_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T8 = Nothing
Dim tag_nhietdolmc_T9
tag_nhietdolmc_T9 = HMIRuntime.Tags("nhietdolmc_T9").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T9 (tag_value, created) values(" & tag_nhietdolmc_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T9 = Nothing
Dim tag_nhietdolmc_T10
tag_nhietdolmc_T10 = HMIRuntime.Tags("nhietdolmc_T10").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T10 (tag_value, created) values(" & tag_nhietdolmc_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T10 = Nothing
Dim tag_nhietdolmc_T11
tag_nhietdolmc_T11 = HMIRuntime.Tags("nhietdolmc_T11").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T11 (tag_value, created) values(" & tag_nhietdolmc_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T11 = Nothing
Dim tag_nhietdolmc_T12
tag_nhietdolmc_T12 = HMIRuntime.Tags("nhietdolmc_T12").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T12 (tag_value, created) values(" & tag_nhietdolmc_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T12 = Nothing
Dim tag_nhietdolmc_T13
tag_nhietdolmc_T13 = HMIRuntime.Tags("nhietdolmc_T13").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T13 (tag_value, created) values(" & tag_nhietdolmc_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T13 = Nothing
Dim tag_nhietdolmc_T14
tag_nhietdolmc_T14 = HMIRuntime.Tags("nhietdolmc_T14").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T14 (tag_value, created) values(" & tag_nhietdolmc_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T14 = Nothing
Dim tag_nhietdolmc_T15
tag_nhietdolmc_T15 = HMIRuntime.Tags("nhietdolmc_T15").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T15 (tag_value, created) values(" & tag_nhietdolmc_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T15 = Nothing
Dim tag_nhietdolmc_T16
tag_nhietdolmc_T16 = HMIRuntime.Tags("nhietdolmc_T16").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T16 (tag_value, created) values(" & tag_nhietdolmc_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T16 = Nothing
Dim tag_nhietdolmc_T17
tag_nhietdolmc_T17 = HMIRuntime.Tags("nhietdolmc_T17").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T17 (tag_value, created) values(" & tag_nhietdolmc_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T17 = Nothing
Dim tag_nhietdolmc_T18
tag_nhietdolmc_T18 = HMIRuntime.Tags("nhietdolmc_T18").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T18 (tag_value, created) values(" & tag_nhietdolmc_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T18 = Nothing
Dim tag_nhietdolmc_T19
tag_nhietdolmc_T19 = HMIRuntime.Tags("nhietdolmc_T19").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T19 (tag_value, created) values(" & tag_nhietdolmc_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T19 = Nothing
Dim tag_nhietdolmc_T20
tag_nhietdolmc_T20 = HMIRuntime.Tags("nhietdolmc_T20").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T20 (tag_value, created) values(" & tag_nhietdolmc_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T20 = Nothing
Dim tag_nhietdolmc_T21
tag_nhietdolmc_T21 = HMIRuntime.Tags("nhietdolmc_T21").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T21 (tag_value, created) values(" & tag_nhietdolmc_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T21 = Nothing
Dim tag_nhietdolmc_T22
tag_nhietdolmc_T22 = HMIRuntime.Tags("nhietdolmc_T22").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T22 (tag_value, created) values(" & tag_nhietdolmc_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T22 = Nothing
Dim tag_nhietdolmc_T23
tag_nhietdolmc_T23 = HMIRuntime.Tags("nhietdolmc_T23").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T23 (tag_value, created) values(" & tag_nhietdolmc_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T23 = Nothing
Dim tag_nhietdolmc_T24
tag_nhietdolmc_T24 = HMIRuntime.Tags("nhietdolmc_T24").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T24 (tag_value, created) values(" & tag_nhietdolmc_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T24 = Nothing
Dim tag_nhietdolmc_T25
tag_nhietdolmc_T25 = HMIRuntime.Tags("nhietdolmc_T25").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T25 (tag_value, created) values(" & tag_nhietdolmc_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T25 = Nothing
Dim tag_nhietdolmc_T26
tag_nhietdolmc_T26 = HMIRuntime.Tags("nhietdolmc_T26").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T26 (tag_value, created) values(" & tag_nhietdolmc_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T26 = Nothing
Dim tag_nhietdolmc_T27
tag_nhietdolmc_T27 = HMIRuntime.Tags("nhietdolmc_T27").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T27 (tag_value, created) values(" & tag_nhietdolmc_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T27 = Nothing
Dim tag_nhietdolmc_T28
tag_nhietdolmc_T28 = HMIRuntime.Tags("nhietdolmc_T28").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T28 (tag_value, created) values(" & tag_nhietdolmc_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T28 = Nothing
Dim tag_nhietdolmc_T29
tag_nhietdolmc_T29 = HMIRuntime.Tags("nhietdolmc_T29").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T29 (tag_value, created) values(" & tag_nhietdolmc_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T29 = Nothing
Dim tag_nhietdolmc_T30
tag_nhietdolmc_T30 = HMIRuntime.Tags("nhietdolmc_T30").Read
strSQL = "INSERT INTO z_tag_nhietdolmc_T30 (tag_value, created) values(" & tag_nhietdolmc_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_nhietdolmc_T30 = Nothing
Dim tag_Nhietdolocbia
tag_Nhietdolocbia = HMIRuntime.Tags("Nhietdolocbia").Read
strSQL = "INSERT INTO z_tag_Nhietdolocbia (tag_value, created) values(" & tag_Nhietdolocbia & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Nhietdolocbia = Nothing
Dim tag_Oxybialoc
tag_Oxybialoc = HMIRuntime.Tags("Oxybialoc").Read
strSQL = "INSERT INTO z_tag_Oxybialoc (tag_value, created) values(" & tag_Oxybialoc & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Oxybialoc = Nothing
Dim tag_OxydauvaoBBT
tag_OxydauvaoBBT = HMIRuntime.Tags("OxydauvaoBBT").Read
strSQL = "INSERT INTO z_tag_OxydauvaoBBT (tag_value, created) values(" & tag_OxydauvaoBBT & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_OxydauvaoBBT = Nothing
Dim tag_PET_Hoimua
tag_PET_Hoimua = HMIRuntime.Tags("PET_Hoimua").Read
strSQL = "INSERT INTO z_tag_PET_Hoimua (tag_value, created) values(" & tag_PET_Hoimua & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PET_Hoimua = Nothing
Dim tag_PH_return_water
tag_PH_return_water = HMIRuntime.Tags("PH_return_water").Read
strSQL = "INSERT INTO z_tag_PH_return_water (tag_value, created) values(" & tag_PH_return_water & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PH_return_water = Nothing
Dim tag_ph_T1
tag_ph_T1 = HMIRuntime.Tags("ph_T1").Read
strSQL = "INSERT INTO z_tag_ph_T1 (tag_value, created) values(" & tag_ph_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T1 = Nothing
Dim tag_ph_T2
tag_ph_T2 = HMIRuntime.Tags("ph_T2").Read
strSQL = "INSERT INTO z_tag_ph_T2 (tag_value, created) values(" & tag_ph_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T2 = Nothing
Dim tag_ph_T3
tag_ph_T3 = HMIRuntime.Tags("ph_T3").Read
strSQL = "INSERT INTO z_tag_ph_T3 (tag_value, created) values(" & tag_ph_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T3 = Nothing
Dim tag_ph_T4
tag_ph_T4 = HMIRuntime.Tags("ph_T4").Read
strSQL = "INSERT INTO z_tag_ph_T4 (tag_value, created) values(" & tag_ph_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T4 = Nothing
Dim tag_ph_T5
tag_ph_T5 = HMIRuntime.Tags("ph_T5").Read
strSQL = "INSERT INTO z_tag_ph_T5 (tag_value, created) values(" & tag_ph_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T5 = Nothing
Dim tag_ph_T6
tag_ph_T6 = HMIRuntime.Tags("ph_T6").Read
strSQL = "INSERT INTO z_tag_ph_T6 (tag_value, created) values(" & tag_ph_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T6 = Nothing
Dim tag_ph_T7
tag_ph_T7 = HMIRuntime.Tags("ph_T7").Read
strSQL = "INSERT INTO z_tag_ph_T7 (tag_value, created) values(" & tag_ph_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T7 = Nothing
Dim tag_ph_T8
tag_ph_T8 = HMIRuntime.Tags("ph_T8").Read
strSQL = "INSERT INTO z_tag_ph_T8 (tag_value, created) values(" & tag_ph_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T8 = Nothing
Dim tag_ph_T9
tag_ph_T9 = HMIRuntime.Tags("ph_T9").Read
strSQL = "INSERT INTO z_tag_ph_T9 (tag_value, created) values(" & tag_ph_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T9 = Nothing
Dim tag_ph_T10
tag_ph_T10 = HMIRuntime.Tags("ph_T10").Read
strSQL = "INSERT INTO z_tag_ph_T10 (tag_value, created) values(" & tag_ph_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T10 = Nothing
Dim tag_ph_T11
tag_ph_T11 = HMIRuntime.Tags("ph_T11").Read
strSQL = "INSERT INTO z_tag_ph_T11 (tag_value, created) values(" & tag_ph_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T11 = Nothing
Dim tag_ph_T12
tag_ph_T12 = HMIRuntime.Tags("ph_T12").Read
strSQL = "INSERT INTO z_tag_ph_T12 (tag_value, created) values(" & tag_ph_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T12 = Nothing
Dim tag_ph_T13
tag_ph_T13 = HMIRuntime.Tags("ph_T13").Read
strSQL = "INSERT INTO z_tag_ph_T13 (tag_value, created) values(" & tag_ph_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T13 = Nothing
Dim tag_ph_T14
tag_ph_T14 = HMIRuntime.Tags("ph_T14").Read
strSQL = "INSERT INTO z_tag_ph_T14 (tag_value, created) values(" & tag_ph_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T14 = Nothing
Dim tag_ph_T15
tag_ph_T15 = HMIRuntime.Tags("ph_T15").Read
strSQL = "INSERT INTO z_tag_ph_T15 (tag_value, created) values(" & tag_ph_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T15 = Nothing
Dim tag_ph_T16
tag_ph_T16 = HMIRuntime.Tags("ph_T16").Read
strSQL = "INSERT INTO z_tag_ph_T16 (tag_value, created) values(" & tag_ph_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T16 = Nothing
Dim tag_ph_T17
tag_ph_T17 = HMIRuntime.Tags("ph_T17").Read
strSQL = "INSERT INTO z_tag_ph_T17 (tag_value, created) values(" & tag_ph_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T17 = Nothing
Dim tag_ph_T18
tag_ph_T18 = HMIRuntime.Tags("ph_T18").Read
strSQL = "INSERT INTO z_tag_ph_T18 (tag_value, created) values(" & tag_ph_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T18 = Nothing
Dim tag_ph_T19
tag_ph_T19 = HMIRuntime.Tags("ph_T19").Read
strSQL = "INSERT INTO z_tag_ph_T19 (tag_value, created) values(" & tag_ph_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T19 = Nothing
Dim tag_ph_T20
tag_ph_T20 = HMIRuntime.Tags("ph_T20").Read
strSQL = "INSERT INTO z_tag_ph_T20 (tag_value, created) values(" & tag_ph_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T20 = Nothing
Dim tag_ph_T21
tag_ph_T21 = HMIRuntime.Tags("ph_T21").Read
strSQL = "INSERT INTO z_tag_ph_T21 (tag_value, created) values(" & tag_ph_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T21 = Nothing
Dim tag_ph_T22
tag_ph_T22 = HMIRuntime.Tags("ph_T22").Read
strSQL = "INSERT INTO z_tag_ph_T22 (tag_value, created) values(" & tag_ph_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T22 = Nothing
Dim tag_ph_T23
tag_ph_T23 = HMIRuntime.Tags("ph_T23").Read
strSQL = "INSERT INTO z_tag_ph_T23 (tag_value, created) values(" & tag_ph_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T23 = Nothing
Dim tag_ph_T24
tag_ph_T24 = HMIRuntime.Tags("ph_T24").Read
strSQL = "INSERT INTO z_tag_ph_T24 (tag_value, created) values(" & tag_ph_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T24 = Nothing
Dim tag_ph_T25
tag_ph_T25 = HMIRuntime.Tags("ph_T25").Read
strSQL = "INSERT INTO z_tag_ph_T25 (tag_value, created) values(" & tag_ph_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T25 = Nothing
Dim tag_ph_T26
tag_ph_T26 = HMIRuntime.Tags("ph_T26").Read
strSQL = "INSERT INTO z_tag_ph_T26 (tag_value, created) values(" & tag_ph_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T26 = Nothing
Dim tag_ph_T27
tag_ph_T27 = HMIRuntime.Tags("ph_T27").Read
strSQL = "INSERT INTO z_tag_ph_T27 (tag_value, created) values(" & tag_ph_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T27 = Nothing
Dim tag_ph_T28
tag_ph_T28 = HMIRuntime.Tags("ph_T28").Read
strSQL = "INSERT INTO z_tag_ph_T28 (tag_value, created) values(" & tag_ph_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T28 = Nothing
Dim tag_ph_T29
tag_ph_T29 = HMIRuntime.Tags("ph_T29").Read
strSQL = "INSERT INTO z_tag_ph_T29 (tag_value, created) values(" & tag_ph_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T29 = Nothing
Dim tag_ph_T30
tag_ph_T30 = HMIRuntime.Tags("ph_T30").Read
strSQL = "INSERT INTO z_tag_ph_T30 (tag_value, created) values(" & tag_ph_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph_T30 = Nothing
Dim tag_ph2_T1
tag_ph2_T1 = HMIRuntime.Tags("ph2_T1").Read
strSQL = "INSERT INTO z_tag_ph2_T1 (tag_value, created) values(" & tag_ph2_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T1 = Nothing
Dim tag_ph2_T2
tag_ph2_T2 = HMIRuntime.Tags("ph2_T2").Read
strSQL = "INSERT INTO z_tag_ph2_T2 (tag_value, created) values(" & tag_ph2_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T2 = Nothing
Dim tag_ph2_T3
tag_ph2_T3 = HMIRuntime.Tags("ph2_T3").Read
strSQL = "INSERT INTO z_tag_ph2_T3 (tag_value, created) values(" & tag_ph2_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T3 = Nothing
Dim tag_ph2_T4
tag_ph2_T4 = HMIRuntime.Tags("ph2_T4").Read
strSQL = "INSERT INTO z_tag_ph2_T4 (tag_value, created) values(" & tag_ph2_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T4 = Nothing
Dim tag_ph2_T5
tag_ph2_T5 = HMIRuntime.Tags("ph2_T5").Read
strSQL = "INSERT INTO z_tag_ph2_T5 (tag_value, created) values(" & tag_ph2_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T5 = Nothing
Dim tag_ph2_T6
tag_ph2_T6 = HMIRuntime.Tags("ph2_T6").Read
strSQL = "INSERT INTO z_tag_ph2_T6 (tag_value, created) values(" & tag_ph2_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T6 = Nothing
Dim tag_ph2_T7
tag_ph2_T7 = HMIRuntime.Tags("ph2_T7").Read
strSQL = "INSERT INTO z_tag_ph2_T7 (tag_value, created) values(" & tag_ph2_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T7 = Nothing
Dim tag_ph2_T8
tag_ph2_T8 = HMIRuntime.Tags("ph2_T8").Read
strSQL = "INSERT INTO z_tag_ph2_T8 (tag_value, created) values(" & tag_ph2_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T8 = Nothing
Dim tag_ph2_T9
tag_ph2_T9 = HMIRuntime.Tags("ph2_T9").Read
strSQL = "INSERT INTO z_tag_ph2_T9 (tag_value, created) values(" & tag_ph2_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T9 = Nothing
Dim tag_ph2_T10
tag_ph2_T10 = HMIRuntime.Tags("ph2_T10").Read
strSQL = "INSERT INTO z_tag_ph2_T10 (tag_value, created) values(" & tag_ph2_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T10 = Nothing
Dim tag_ph2_T11
tag_ph2_T11 = HMIRuntime.Tags("ph2_T11").Read
strSQL = "INSERT INTO z_tag_ph2_T11 (tag_value, created) values(" & tag_ph2_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T11 = Nothing
Dim tag_ph2_T12
tag_ph2_T12 = HMIRuntime.Tags("ph2_T12").Read
strSQL = "INSERT INTO z_tag_ph2_T12 (tag_value, created) values(" & tag_ph2_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T12 = Nothing
Dim tag_ph2_T13
tag_ph2_T13 = HMIRuntime.Tags("ph2_T13").Read
strSQL = "INSERT INTO z_tag_ph2_T13 (tag_value, created) values(" & tag_ph2_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T13 = Nothing
Dim tag_ph2_T14
tag_ph2_T14 = HMIRuntime.Tags("ph2_T14").Read
strSQL = "INSERT INTO z_tag_ph2_T14 (tag_value, created) values(" & tag_ph2_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T14 = Nothing
Dim tag_ph2_T15
tag_ph2_T15 = HMIRuntime.Tags("ph2_T15").Read
strSQL = "INSERT INTO z_tag_ph2_T15 (tag_value, created) values(" & tag_ph2_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T15 = Nothing
Dim tag_ph2_T16
tag_ph2_T16 = HMIRuntime.Tags("ph2_T16").Read
strSQL = "INSERT INTO z_tag_ph2_T16 (tag_value, created) values(" & tag_ph2_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T16 = Nothing
Dim tag_ph2_T17
tag_ph2_T17 = HMIRuntime.Tags("ph2_T17").Read
strSQL = "INSERT INTO z_tag_ph2_T17 (tag_value, created) values(" & tag_ph2_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T17 = Nothing
Dim tag_ph2_T18
tag_ph2_T18 = HMIRuntime.Tags("ph2_T18").Read
strSQL = "INSERT INTO z_tag_ph2_T18 (tag_value, created) values(" & tag_ph2_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T18 = Nothing
Dim tag_ph2_T19
tag_ph2_T19 = HMIRuntime.Tags("ph2_T19").Read
strSQL = "INSERT INTO z_tag_ph2_T19 (tag_value, created) values(" & tag_ph2_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T19 = Nothing
Dim tag_ph2_T20
tag_ph2_T20 = HMIRuntime.Tags("ph2_T20").Read
strSQL = "INSERT INTO z_tag_ph2_T20 (tag_value, created) values(" & tag_ph2_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T20 = Nothing
Dim tag_ph2_T21
tag_ph2_T21 = HMIRuntime.Tags("ph2_T21").Read
strSQL = "INSERT INTO z_tag_ph2_T21 (tag_value, created) values(" & tag_ph2_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T21 = Nothing
Dim tag_ph2_T22
tag_ph2_T22 = HMIRuntime.Tags("ph2_T22").Read
strSQL = "INSERT INTO z_tag_ph2_T22 (tag_value, created) values(" & tag_ph2_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T22 = Nothing
Dim tag_ph2_T23
tag_ph2_T23 = HMIRuntime.Tags("ph2_T23").Read
strSQL = "INSERT INTO z_tag_ph2_T23 (tag_value, created) values(" & tag_ph2_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T23 = Nothing
Dim tag_ph2_T24
tag_ph2_T24 = HMIRuntime.Tags("ph2_T24").Read
strSQL = "INSERT INTO z_tag_ph2_T24 (tag_value, created) values(" & tag_ph2_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T24 = Nothing
Dim tag_ph2_T25
tag_ph2_T25 = HMIRuntime.Tags("ph2_T25").Read
strSQL = "INSERT INTO z_tag_ph2_T25 (tag_value, created) values(" & tag_ph2_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T25 = Nothing
Dim tag_ph2_T26
tag_ph2_T26 = HMIRuntime.Tags("ph2_T26").Read
strSQL = "INSERT INTO z_tag_ph2_T26 (tag_value, created) values(" & tag_ph2_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T26 = Nothing
Dim tag_ph2_T27
tag_ph2_T27 = HMIRuntime.Tags("ph2_T27").Read
strSQL = "INSERT INTO z_tag_ph2_T27 (tag_value, created) values(" & tag_ph2_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T27 = Nothing
Dim tag_ph2_T28
tag_ph2_T28 = HMIRuntime.Tags("ph2_T28").Read
strSQL = "INSERT INTO z_tag_ph2_T28 (tag_value, created) values(" & tag_ph2_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T28 = Nothing
Dim tag_ph2_T29
tag_ph2_T29 = HMIRuntime.Tags("ph2_T29").Read
strSQL = "INSERT INTO z_tag_ph2_T29 (tag_value, created) values(" & tag_ph2_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T29 = Nothing
Dim tag_ph2_T30
tag_ph2_T30 = HMIRuntime.Tags("ph2_T30").Read
strSQL = "INSERT INTO z_tag_ph2_T30 (tag_value, created) values(" & tag_ph2_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ph2_T30 = Nothing
Dim tag_PID_0401M01
tag_PID_0401M01 = HMIRuntime.Tags("PID_0401M01").Read
strSQL = "INSERT INTO z_tag_PID_0401M01 (tag_value, created) values(" & tag_PID_0401M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID_0401M01 = Nothing
Dim tag_PID27_0401M05
tag_PID27_0401M05 = HMIRuntime.Tags("PID27_0401M05").Read
strSQL = "INSERT INTO z_tag_PID27_0401M05 (tag_value, created) values(" & tag_PID27_0401M05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_PID27_0401M05 = Nothing
Dim tag_platoda_T1
tag_platoda_T1 = HMIRuntime.Tags("platoda_T1").Read
strSQL = "INSERT INTO z_tag_platoda_T1 (tag_value, created) values(" & tag_platoda_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T1 = Nothing
Dim tag_platoda_T2
tag_platoda_T2 = HMIRuntime.Tags("platoda_T2").Read
strSQL = "INSERT INTO z_tag_platoda_T2 (tag_value, created) values(" & tag_platoda_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T2 = Nothing
Dim tag_platoda_T3
tag_platoda_T3 = HMIRuntime.Tags("platoda_T3").Read
strSQL = "INSERT INTO z_tag_platoda_T3 (tag_value, created) values(" & tag_platoda_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T3 = Nothing
Dim tag_platoda_T4
tag_platoda_T4 = HMIRuntime.Tags("platoda_T4").Read
strSQL = "INSERT INTO z_tag_platoda_T4 (tag_value, created) values(" & tag_platoda_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T4 = Nothing
Dim tag_platoda_T5
tag_platoda_T5 = HMIRuntime.Tags("platoda_T5").Read
strSQL = "INSERT INTO z_tag_platoda_T5 (tag_value, created) values(" & tag_platoda_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T5 = Nothing
Dim tag_platoda_T6
tag_platoda_T6 = HMIRuntime.Tags("platoda_T6").Read
strSQL = "INSERT INTO z_tag_platoda_T6 (tag_value, created) values(" & tag_platoda_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T6 = Nothing
Dim tag_platoda_T7
tag_platoda_T7 = HMIRuntime.Tags("platoda_T7").Read
strSQL = "INSERT INTO z_tag_platoda_T7 (tag_value, created) values(" & tag_platoda_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T7 = Nothing
Dim tag_platoda_T8
tag_platoda_T8 = HMIRuntime.Tags("platoda_T8").Read
strSQL = "INSERT INTO z_tag_platoda_T8 (tag_value, created) values(" & tag_platoda_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T8 = Nothing
Dim tag_platoda_T9
tag_platoda_T9 = HMIRuntime.Tags("platoda_T9").Read
strSQL = "INSERT INTO z_tag_platoda_T9 (tag_value, created) values(" & tag_platoda_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T9 = Nothing
Dim tag_platoda_T10
tag_platoda_T10 = HMIRuntime.Tags("platoda_T10").Read
strSQL = "INSERT INTO z_tag_platoda_T10 (tag_value, created) values(" & tag_platoda_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T10 = Nothing
Dim tag_platoda_T11
tag_platoda_T11 = HMIRuntime.Tags("platoda_T11").Read
strSQL = "INSERT INTO z_tag_platoda_T11 (tag_value, created) values(" & tag_platoda_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T11 = Nothing
Dim tag_platoda_T12
tag_platoda_T12 = HMIRuntime.Tags("platoda_T12").Read
strSQL = "INSERT INTO z_tag_platoda_T12 (tag_value, created) values(" & tag_platoda_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T12 = Nothing
Dim tag_platoda_T13
tag_platoda_T13 = HMIRuntime.Tags("platoda_T13").Read
strSQL = "INSERT INTO z_tag_platoda_T13 (tag_value, created) values(" & tag_platoda_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T13 = Nothing
Dim tag_platoda_T14
tag_platoda_T14 = HMIRuntime.Tags("platoda_T14").Read
strSQL = "INSERT INTO z_tag_platoda_T14 (tag_value, created) values(" & tag_platoda_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T14 = Nothing
Dim tag_platoda_T15
tag_platoda_T15 = HMIRuntime.Tags("platoda_T15").Read
strSQL = "INSERT INTO z_tag_platoda_T15 (tag_value, created) values(" & tag_platoda_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T15 = Nothing
Dim tag_platoda_T16
tag_platoda_T16 = HMIRuntime.Tags("platoda_T16").Read
strSQL = "INSERT INTO z_tag_platoda_T16 (tag_value, created) values(" & tag_platoda_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T16 = Nothing
Dim tag_platoda_T17
tag_platoda_T17 = HMIRuntime.Tags("platoda_T17").Read
strSQL = "INSERT INTO z_tag_platoda_T17 (tag_value, created) values(" & tag_platoda_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T17 = Nothing
Dim tag_platoda_T18
tag_platoda_T18 = HMIRuntime.Tags("platoda_T18").Read
strSQL = "INSERT INTO z_tag_platoda_T18 (tag_value, created) values(" & tag_platoda_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T18 = Nothing
Dim tag_platoda_T19
tag_platoda_T19 = HMIRuntime.Tags("platoda_T19").Read
strSQL = "INSERT INTO z_tag_platoda_T19 (tag_value, created) values(" & tag_platoda_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T19 = Nothing
Dim tag_platoda_T20
tag_platoda_T20 = HMIRuntime.Tags("platoda_T20").Read
strSQL = "INSERT INTO z_tag_platoda_T20 (tag_value, created) values(" & tag_platoda_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T20 = Nothing
Dim tag_platoda_T21
tag_platoda_T21 = HMIRuntime.Tags("platoda_T21").Read
strSQL = "INSERT INTO z_tag_platoda_T21 (tag_value, created) values(" & tag_platoda_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T21 = Nothing
Dim tag_platoda_T22
tag_platoda_T22 = HMIRuntime.Tags("platoda_T22").Read
strSQL = "INSERT INTO z_tag_platoda_T22 (tag_value, created) values(" & tag_platoda_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T22 = Nothing
Dim tag_platoda_T23
tag_platoda_T23 = HMIRuntime.Tags("platoda_T23").Read
strSQL = "INSERT INTO z_tag_platoda_T23 (tag_value, created) values(" & tag_platoda_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T23 = Nothing
Dim tag_platoda_T24
tag_platoda_T24 = HMIRuntime.Tags("platoda_T24").Read
strSQL = "INSERT INTO z_tag_platoda_T24 (tag_value, created) values(" & tag_platoda_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T24 = Nothing
Dim tag_platoda_T25
tag_platoda_T25 = HMIRuntime.Tags("platoda_T25").Read
strSQL = "INSERT INTO z_tag_platoda_T25 (tag_value, created) values(" & tag_platoda_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T25 = Nothing
Dim tag_platoda_T26
tag_platoda_T26 = HMIRuntime.Tags("platoda_T26").Read
strSQL = "INSERT INTO z_tag_platoda_T26 (tag_value, created) values(" & tag_platoda_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T26 = Nothing
Dim tag_platoda_T27
tag_platoda_T27 = HMIRuntime.Tags("platoda_T27").Read
strSQL = "INSERT INTO z_tag_platoda_T27 (tag_value, created) values(" & tag_platoda_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T27 = Nothing
Dim tag_platoda_T28
tag_platoda_T28 = HMIRuntime.Tags("platoda_T28").Read
strSQL = "INSERT INTO z_tag_platoda_T28 (tag_value, created) values(" & tag_platoda_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T28 = Nothing
Dim tag_platoda_T29
tag_platoda_T29 = HMIRuntime.Tags("platoda_T29").Read
strSQL = "INSERT INTO z_tag_platoda_T29 (tag_value, created) values(" & tag_platoda_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T29 = Nothing
Dim tag_platoda_T30
tag_platoda_T30 = HMIRuntime.Tags("platoda_T30").Read
strSQL = "INSERT INTO z_tag_platoda_T30 (tag_value, created) values(" & tag_platoda_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platoda_T30 = Nothing
Dim tag_platohnd_T1
tag_platohnd_T1 = HMIRuntime.Tags("platohnd_T1").Read
strSQL = "INSERT INTO z_tag_platohnd_T1 (tag_value, created) values(" & tag_platohnd_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T1 = Nothing
Dim tag_platohnd_T2
tag_platohnd_T2 = HMIRuntime.Tags("platohnd_T2").Read
strSQL = "INSERT INTO z_tag_platohnd_T2 (tag_value, created) values(" & tag_platohnd_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T2 = Nothing
Dim tag_platohnd_T3
tag_platohnd_T3 = HMIRuntime.Tags("platohnd_T3").Read
strSQL = "INSERT INTO z_tag_platohnd_T3 (tag_value, created) values(" & tag_platohnd_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T3 = Nothing
Dim tag_platohnd_T4
tag_platohnd_T4 = HMIRuntime.Tags("platohnd_T4").Read
strSQL = "INSERT INTO z_tag_platohnd_T4 (tag_value, created) values(" & tag_platohnd_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T4 = Nothing
Dim tag_platohnd_T5
tag_platohnd_T5 = HMIRuntime.Tags("platohnd_T5").Read
strSQL = "INSERT INTO z_tag_platohnd_T5 (tag_value, created) values(" & tag_platohnd_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T5 = Nothing
Dim tag_platohnd_T6
tag_platohnd_T6 = HMIRuntime.Tags("platohnd_T6").Read
strSQL = "INSERT INTO z_tag_platohnd_T6 (tag_value, created) values(" & tag_platohnd_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T6 = Nothing
Dim tag_platohnd_T7
tag_platohnd_T7 = HMIRuntime.Tags("platohnd_T7").Read
strSQL = "INSERT INTO z_tag_platohnd_T7 (tag_value, created) values(" & tag_platohnd_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T7 = Nothing
Dim tag_platohnd_T8
tag_platohnd_T8 = HMIRuntime.Tags("platohnd_T8").Read
strSQL = "INSERT INTO z_tag_platohnd_T8 (tag_value, created) values(" & tag_platohnd_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T8 = Nothing
Dim tag_platohnd_T9
tag_platohnd_T9 = HMIRuntime.Tags("platohnd_T9").Read
strSQL = "INSERT INTO z_tag_platohnd_T9 (tag_value, created) values(" & tag_platohnd_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T9 = Nothing
Dim tag_platohnd_T10
tag_platohnd_T10 = HMIRuntime.Tags("platohnd_T10").Read
strSQL = "INSERT INTO z_tag_platohnd_T10 (tag_value, created) values(" & tag_platohnd_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T10 = Nothing
Dim tag_platohnd_T11
tag_platohnd_T11 = HMIRuntime.Tags("platohnd_T11").Read
strSQL = "INSERT INTO z_tag_platohnd_T11 (tag_value, created) values(" & tag_platohnd_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T11 = Nothing
Dim tag_platohnd_T12
tag_platohnd_T12 = HMIRuntime.Tags("platohnd_T12").Read
strSQL = "INSERT INTO z_tag_platohnd_T12 (tag_value, created) values(" & tag_platohnd_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T12 = Nothing
Dim tag_platohnd_T13
tag_platohnd_T13 = HMIRuntime.Tags("platohnd_T13").Read
strSQL = "INSERT INTO z_tag_platohnd_T13 (tag_value, created) values(" & tag_platohnd_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T13 = Nothing
Dim tag_platohnd_T14
tag_platohnd_T14 = HMIRuntime.Tags("platohnd_T14").Read
strSQL = "INSERT INTO z_tag_platohnd_T14 (tag_value, created) values(" & tag_platohnd_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T14 = Nothing
Dim tag_platohnd_T15
tag_platohnd_T15 = HMIRuntime.Tags("platohnd_T15").Read
strSQL = "INSERT INTO z_tag_platohnd_T15 (tag_value, created) values(" & tag_platohnd_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T15 = Nothing
Dim tag_platohnd_T16
tag_platohnd_T16 = HMIRuntime.Tags("platohnd_T16").Read
strSQL = "INSERT INTO z_tag_platohnd_T16 (tag_value, created) values(" & tag_platohnd_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T16 = Nothing
Dim tag_platohnd_T17
tag_platohnd_T17 = HMIRuntime.Tags("platohnd_T17").Read
strSQL = "INSERT INTO z_tag_platohnd_T17 (tag_value, created) values(" & tag_platohnd_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T17 = Nothing
Dim tag_platohnd_T18
tag_platohnd_T18 = HMIRuntime.Tags("platohnd_T18").Read
strSQL = "INSERT INTO z_tag_platohnd_T18 (tag_value, created) values(" & tag_platohnd_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T18 = Nothing
Dim tag_platohnd_T19
tag_platohnd_T19 = HMIRuntime.Tags("platohnd_T19").Read
strSQL = "INSERT INTO z_tag_platohnd_T19 (tag_value, created) values(" & tag_platohnd_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T19 = Nothing
Dim tag_platohnd_T20
tag_platohnd_T20 = HMIRuntime.Tags("platohnd_T20").Read
strSQL = "INSERT INTO z_tag_platohnd_T20 (tag_value, created) values(" & tag_platohnd_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T20 = Nothing
Dim tag_platohnd_T21
tag_platohnd_T21 = HMIRuntime.Tags("platohnd_T21").Read
strSQL = "INSERT INTO z_tag_platohnd_T21 (tag_value, created) values(" & tag_platohnd_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T21 = Nothing
Dim tag_platohnd_T22
tag_platohnd_T22 = HMIRuntime.Tags("platohnd_T22").Read
strSQL = "INSERT INTO z_tag_platohnd_T22 (tag_value, created) values(" & tag_platohnd_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T22 = Nothing
Dim tag_platohnd_T23
tag_platohnd_T23 = HMIRuntime.Tags("platohnd_T23").Read
strSQL = "INSERT INTO z_tag_platohnd_T23 (tag_value, created) values(" & tag_platohnd_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T23 = Nothing
Dim tag_platohnd_T24
tag_platohnd_T24 = HMIRuntime.Tags("platohnd_T24").Read
strSQL = "INSERT INTO z_tag_platohnd_T24 (tag_value, created) values(" & tag_platohnd_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T24 = Nothing
Dim tag_platohnd_T25
tag_platohnd_T25 = HMIRuntime.Tags("platohnd_T25").Read
strSQL = "INSERT INTO z_tag_platohnd_T25 (tag_value, created) values(" & tag_platohnd_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T25 = Nothing
Dim tag_platohnd_T26
tag_platohnd_T26 = HMIRuntime.Tags("platohnd_T26").Read
strSQL = "INSERT INTO z_tag_platohnd_T26 (tag_value, created) values(" & tag_platohnd_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T26 = Nothing
Dim tag_platohnd_T27
tag_platohnd_T27 = HMIRuntime.Tags("platohnd_T27").Read
strSQL = "INSERT INTO z_tag_platohnd_T27 (tag_value, created) values(" & tag_platohnd_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T27 = Nothing
Dim tag_platohnd_T28
tag_platohnd_T28 = HMIRuntime.Tags("platohnd_T28").Read
strSQL = "INSERT INTO z_tag_platohnd_T28 (tag_value, created) values(" & tag_platohnd_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T28 = Nothing
Dim tag_platohnd_T29
tag_platohnd_T29 = HMIRuntime.Tags("platohnd_T29").Read
strSQL = "INSERT INTO z_tag_platohnd_T29 (tag_value, created) values(" & tag_platohnd_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T29 = Nothing
Dim tag_platohnd_T30
tag_platohnd_T30 = HMIRuntime.Tags("platohnd_T30").Read
strSQL = "INSERT INTO z_tag_platohnd_T30 (tag_value, created) values(" & tag_platohnd_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_platohnd_T30 = Nothing
Dim tag_SAPDENGIOCAODIEM
tag_SAPDENGIOCAODIEM = HMIRuntime.Tags("SAPDENGIOCAODIEM").Read
strSQL = "INSERT INTO z_tag_SAPDENGIOCAODIEM (tag_value, created) values(" & tag_SAPDENGIOCAODIEM & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_SAPDENGIOCAODIEM = Nothing
Dim tag_sl_tankthuCO2_men
tag_sl_tankthuCO2_men = HMIRuntime.Tags("sl_tankthuCO2_men").Read
strSQL = "INSERT INTO z_tag_sl_tankthuCO2_men (tag_value, created) values(" & tag_sl_tankthuCO2_men & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_sl_tankthuCO2_men = Nothing
Dim tag_Tank01_auto
tag_Tank01_auto = HMIRuntime.Tags("Tank01_auto").Read
strSQL = "INSERT INTO z_tag_Tank01_auto (tag_value, created) values(" & tag_Tank01_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank01_auto = Nothing
Dim tag_Tank01_CIP
tag_Tank01_CIP = HMIRuntime.Tags("Tank01_CIP").Read
strSQL = "INSERT INTO z_tag_Tank01_CIP (tag_value, created) values(" & tag_Tank01_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank01_CIP = Nothing
Dim tag_Tank01_run
tag_Tank01_run = HMIRuntime.Tags("Tank01_run").Read
strSQL = "INSERT INTO z_tag_Tank01_run (tag_value, created) values(" & tag_Tank01_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank01_run = Nothing
Dim tag_Tank01_seq
tag_Tank01_seq = HMIRuntime.Tags("Tank01_seq").Read
strSQL = "INSERT INTO z_tag_Tank01_seq (tag_value, created) values(" & tag_Tank01_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank01_seq = Nothing
Dim tag_Tank02_auto
tag_Tank02_auto = HMIRuntime.Tags("Tank02_auto").Read
strSQL = "INSERT INTO z_tag_Tank02_auto (tag_value, created) values(" & tag_Tank02_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank02_auto = Nothing
Dim tag_Tank02_CIP
tag_Tank02_CIP = HMIRuntime.Tags("Tank02_CIP").Read
strSQL = "INSERT INTO z_tag_Tank02_CIP (tag_value, created) values(" & tag_Tank02_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank02_CIP = Nothing
Dim tag_Tank02_run
tag_Tank02_run = HMIRuntime.Tags("Tank02_run").Read
strSQL = "INSERT INTO z_tag_Tank02_run (tag_value, created) values(" & tag_Tank02_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank02_run = Nothing
Dim tag_Tank02_seq
tag_Tank02_seq = HMIRuntime.Tags("Tank02_seq").Read
strSQL = "INSERT INTO z_tag_Tank02_seq (tag_value, created) values(" & tag_Tank02_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank02_seq = Nothing
Dim tag_Tank03_auto
tag_Tank03_auto = HMIRuntime.Tags("Tank03_auto").Read
strSQL = "INSERT INTO z_tag_Tank03_auto (tag_value, created) values(" & tag_Tank03_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank03_auto = Nothing
Dim tag_Tank03_CIP
tag_Tank03_CIP = HMIRuntime.Tags("Tank03_CIP").Read
strSQL = "INSERT INTO z_tag_Tank03_CIP (tag_value, created) values(" & tag_Tank03_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank03_CIP = Nothing
Dim tag_Tank03_run
tag_Tank03_run = HMIRuntime.Tags("Tank03_run").Read
strSQL = "INSERT INTO z_tag_Tank03_run (tag_value, created) values(" & tag_Tank03_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank03_run = Nothing
Dim tag_Tank03_seq
tag_Tank03_seq = HMIRuntime.Tags("Tank03_seq").Read
strSQL = "INSERT INTO z_tag_Tank03_seq (tag_value, created) values(" & tag_Tank03_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank03_seq = Nothing
Dim tag_Tank04_auto
tag_Tank04_auto = HMIRuntime.Tags("Tank04_auto").Read
strSQL = "INSERT INTO z_tag_Tank04_auto (tag_value, created) values(" & tag_Tank04_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank04_auto = Nothing
Dim tag_Tank04_CIP
tag_Tank04_CIP = HMIRuntime.Tags("Tank04_CIP").Read
strSQL = "INSERT INTO z_tag_Tank04_CIP (tag_value, created) values(" & tag_Tank04_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank04_CIP = Nothing
Dim tag_Tank04_run
tag_Tank04_run = HMIRuntime.Tags("Tank04_run").Read
strSQL = "INSERT INTO z_tag_Tank04_run (tag_value, created) values(" & tag_Tank04_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank04_run = Nothing
Dim tag_Tank04_seq
tag_Tank04_seq = HMIRuntime.Tags("Tank04_seq").Read
strSQL = "INSERT INTO z_tag_Tank04_seq (tag_value, created) values(" & tag_Tank04_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank04_seq = Nothing
Dim tag_Tank05_auto
tag_Tank05_auto = HMIRuntime.Tags("Tank05_auto").Read
strSQL = "INSERT INTO z_tag_Tank05_auto (tag_value, created) values(" & tag_Tank05_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank05_auto = Nothing
Dim tag_Tank05_CIP
tag_Tank05_CIP = HMIRuntime.Tags("Tank05_CIP").Read
strSQL = "INSERT INTO z_tag_Tank05_CIP (tag_value, created) values(" & tag_Tank05_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank05_CIP = Nothing
Dim tag_Tank05_run
tag_Tank05_run = HMIRuntime.Tags("Tank05_run").Read
strSQL = "INSERT INTO z_tag_Tank05_run (tag_value, created) values(" & tag_Tank05_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank05_run = Nothing
Dim tag_Tank05_seq
tag_Tank05_seq = HMIRuntime.Tags("Tank05_seq").Read
strSQL = "INSERT INTO z_tag_Tank05_seq (tag_value, created) values(" & tag_Tank05_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank05_seq = Nothing
Dim tag_Tank06_auto
tag_Tank06_auto = HMIRuntime.Tags("Tank06_auto").Read
strSQL = "INSERT INTO z_tag_Tank06_auto (tag_value, created) values(" & tag_Tank06_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank06_auto = Nothing
Dim tag_Tank06_CIP
tag_Tank06_CIP = HMIRuntime.Tags("Tank06_CIP").Read
strSQL = "INSERT INTO z_tag_Tank06_CIP (tag_value, created) values(" & tag_Tank06_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank06_CIP = Nothing
Dim tag_Tank06_run
tag_Tank06_run = HMIRuntime.Tags("Tank06_run").Read
strSQL = "INSERT INTO z_tag_Tank06_run (tag_value, created) values(" & tag_Tank06_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank06_run = Nothing
Dim tag_Tank06_seq
tag_Tank06_seq = HMIRuntime.Tags("Tank06_seq").Read
strSQL = "INSERT INTO z_tag_Tank06_seq (tag_value, created) values(" & tag_Tank06_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank06_seq = Nothing
Dim tag_Tank07_auto
tag_Tank07_auto = HMIRuntime.Tags("Tank07_auto").Read
strSQL = "INSERT INTO z_tag_Tank07_auto (tag_value, created) values(" & tag_Tank07_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank07_auto = Nothing
Dim tag_Tank07_CIP
tag_Tank07_CIP = HMIRuntime.Tags("Tank07_CIP").Read
strSQL = "INSERT INTO z_tag_Tank07_CIP (tag_value, created) values(" & tag_Tank07_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank07_CIP = Nothing
Dim tag_Tank07_run
tag_Tank07_run = HMIRuntime.Tags("Tank07_run").Read
strSQL = "INSERT INTO z_tag_Tank07_run (tag_value, created) values(" & tag_Tank07_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank07_run = Nothing
Dim tag_Tank07_seq
tag_Tank07_seq = HMIRuntime.Tags("Tank07_seq").Read
strSQL = "INSERT INTO z_tag_Tank07_seq (tag_value, created) values(" & tag_Tank07_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank07_seq = Nothing
Dim tag_Tank08_auto
tag_Tank08_auto = HMIRuntime.Tags("Tank08_auto").Read
strSQL = "INSERT INTO z_tag_Tank08_auto (tag_value, created) values(" & tag_Tank08_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank08_auto = Nothing
Dim tag_Tank08_CIP
tag_Tank08_CIP = HMIRuntime.Tags("Tank08_CIP").Read
strSQL = "INSERT INTO z_tag_Tank08_CIP (tag_value, created) values(" & tag_Tank08_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank08_CIP = Nothing
Dim tag_Tank08_run
tag_Tank08_run = HMIRuntime.Tags("Tank08_run").Read
strSQL = "INSERT INTO z_tag_Tank08_run (tag_value, created) values(" & tag_Tank08_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank08_run = Nothing
Dim tag_Tank08_seq
tag_Tank08_seq = HMIRuntime.Tags("Tank08_seq").Read
strSQL = "INSERT INTO z_tag_Tank08_seq (tag_value, created) values(" & tag_Tank08_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank08_seq = Nothing
Dim tag_Tank09_auto
tag_Tank09_auto = HMIRuntime.Tags("Tank09_auto").Read
strSQL = "INSERT INTO z_tag_Tank09_auto (tag_value, created) values(" & tag_Tank09_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank09_auto = Nothing
Dim tag_Tank09_CIP
tag_Tank09_CIP = HMIRuntime.Tags("Tank09_CIP").Read
strSQL = "INSERT INTO z_tag_Tank09_CIP (tag_value, created) values(" & tag_Tank09_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank09_CIP = Nothing
Dim tag_Tank09_run
tag_Tank09_run = HMIRuntime.Tags("Tank09_run").Read
strSQL = "INSERT INTO z_tag_Tank09_run (tag_value, created) values(" & tag_Tank09_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank09_run = Nothing
Dim tag_Tank09_seq
tag_Tank09_seq = HMIRuntime.Tags("Tank09_seq").Read
strSQL = "INSERT INTO z_tag_Tank09_seq (tag_value, created) values(" & tag_Tank09_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank09_seq = Nothing
Dim tag_Tank10_auto
tag_Tank10_auto = HMIRuntime.Tags("Tank10_auto").Read
strSQL = "INSERT INTO z_tag_Tank10_auto (tag_value, created) values(" & tag_Tank10_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank10_auto = Nothing
Dim tag_Tank10_CIP
tag_Tank10_CIP = HMIRuntime.Tags("Tank10_CIP").Read
strSQL = "INSERT INTO z_tag_Tank10_CIP (tag_value, created) values(" & tag_Tank10_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank10_CIP = Nothing
Dim tag_Tank10_run
tag_Tank10_run = HMIRuntime.Tags("Tank10_run").Read
strSQL = "INSERT INTO z_tag_Tank10_run (tag_value, created) values(" & tag_Tank10_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank10_run = Nothing
Dim tag_Tank10_seq
tag_Tank10_seq = HMIRuntime.Tags("Tank10_seq").Read
strSQL = "INSERT INTO z_tag_Tank10_seq (tag_value, created) values(" & tag_Tank10_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank10_seq = Nothing
Dim tag_Tank11_auto
tag_Tank11_auto = HMIRuntime.Tags("Tank11_auto").Read
strSQL = "INSERT INTO z_tag_Tank11_auto (tag_value, created) values(" & tag_Tank11_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank11_auto = Nothing
Dim tag_Tank11_CIP
tag_Tank11_CIP = HMIRuntime.Tags("Tank11_CIP").Read
strSQL = "INSERT INTO z_tag_Tank11_CIP (tag_value, created) values(" & tag_Tank11_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank11_CIP = Nothing
Dim tag_Tank11_run
tag_Tank11_run = HMIRuntime.Tags("Tank11_run").Read
strSQL = "INSERT INTO z_tag_Tank11_run (tag_value, created) values(" & tag_Tank11_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank11_run = Nothing
Dim tag_Tank11_seq
tag_Tank11_seq = HMIRuntime.Tags("Tank11_seq").Read
strSQL = "INSERT INTO z_tag_Tank11_seq (tag_value, created) values(" & tag_Tank11_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank11_seq = Nothing
Dim tag_Tank12_auto
tag_Tank12_auto = HMIRuntime.Tags("Tank12_auto").Read
strSQL = "INSERT INTO z_tag_Tank12_auto (tag_value, created) values(" & tag_Tank12_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank12_auto = Nothing
Dim tag_Tank12_CIP
tag_Tank12_CIP = HMIRuntime.Tags("Tank12_CIP").Read
strSQL = "INSERT INTO z_tag_Tank12_CIP (tag_value, created) values(" & tag_Tank12_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank12_CIP = Nothing
Dim tag_Tank12_run
tag_Tank12_run = HMIRuntime.Tags("Tank12_run").Read
strSQL = "INSERT INTO z_tag_Tank12_run (tag_value, created) values(" & tag_Tank12_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank12_run = Nothing
Dim tag_Tank12_seq
tag_Tank12_seq = HMIRuntime.Tags("Tank12_seq").Read
strSQL = "INSERT INTO z_tag_Tank12_seq (tag_value, created) values(" & tag_Tank12_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank12_seq = Nothing
Dim tag_Tank13_auto
tag_Tank13_auto = HMIRuntime.Tags("Tank13_auto").Read
strSQL = "INSERT INTO z_tag_Tank13_auto (tag_value, created) values(" & tag_Tank13_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank13_auto = Nothing
Dim tag_Tank13_CIP
tag_Tank13_CIP = HMIRuntime.Tags("Tank13_CIP").Read
strSQL = "INSERT INTO z_tag_Tank13_CIP (tag_value, created) values(" & tag_Tank13_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank13_CIP = Nothing
Dim tag_Tank13_run
tag_Tank13_run = HMIRuntime.Tags("Tank13_run").Read
strSQL = "INSERT INTO z_tag_Tank13_run (tag_value, created) values(" & tag_Tank13_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank13_run = Nothing
Dim tag_Tank13_seq
tag_Tank13_seq = HMIRuntime.Tags("Tank13_seq").Read
strSQL = "INSERT INTO z_tag_Tank13_seq (tag_value, created) values(" & tag_Tank13_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank13_seq = Nothing
Dim tag_Tank14_auto
tag_Tank14_auto = HMIRuntime.Tags("Tank14_auto").Read
strSQL = "INSERT INTO z_tag_Tank14_auto (tag_value, created) values(" & tag_Tank14_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank14_auto = Nothing
Dim tag_Tank14_CIP
tag_Tank14_CIP = HMIRuntime.Tags("Tank14_CIP").Read
strSQL = "INSERT INTO z_tag_Tank14_CIP (tag_value, created) values(" & tag_Tank14_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank14_CIP = Nothing
Dim tag_Tank14_run
tag_Tank14_run = HMIRuntime.Tags("Tank14_run").Read
strSQL = "INSERT INTO z_tag_Tank14_run (tag_value, created) values(" & tag_Tank14_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank14_run = Nothing
Dim tag_Tank14_seq
tag_Tank14_seq = HMIRuntime.Tags("Tank14_seq").Read
strSQL = "INSERT INTO z_tag_Tank14_seq (tag_value, created) values(" & tag_Tank14_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank14_seq = Nothing
Dim tag_Tank15_auto
tag_Tank15_auto = HMIRuntime.Tags("Tank15_auto").Read
strSQL = "INSERT INTO z_tag_Tank15_auto (tag_value, created) values(" & tag_Tank15_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank15_auto = Nothing
Dim tag_Tank15_CIP
tag_Tank15_CIP = HMIRuntime.Tags("Tank15_CIP").Read
strSQL = "INSERT INTO z_tag_Tank15_CIP (tag_value, created) values(" & tag_Tank15_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank15_CIP = Nothing
Dim tag_Tank15_run
tag_Tank15_run = HMIRuntime.Tags("Tank15_run").Read
strSQL = "INSERT INTO z_tag_Tank15_run (tag_value, created) values(" & tag_Tank15_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank15_run = Nothing
Dim tag_Tank15_seq
tag_Tank15_seq = HMIRuntime.Tags("Tank15_seq").Read
strSQL = "INSERT INTO z_tag_Tank15_seq (tag_value, created) values(" & tag_Tank15_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank15_seq = Nothing
Dim tag_Tank16_auto
tag_Tank16_auto = HMIRuntime.Tags("Tank16_auto").Read
strSQL = "INSERT INTO z_tag_Tank16_auto (tag_value, created) values(" & tag_Tank16_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank16_auto = Nothing
Dim tag_Tank16_CIP
tag_Tank16_CIP = HMIRuntime.Tags("Tank16_CIP").Read
strSQL = "INSERT INTO z_tag_Tank16_CIP (tag_value, created) values(" & tag_Tank16_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank16_CIP = Nothing
Dim tag_Tank16_run
tag_Tank16_run = HMIRuntime.Tags("Tank16_run").Read
strSQL = "INSERT INTO z_tag_Tank16_run (tag_value, created) values(" & tag_Tank16_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank16_run = Nothing
Dim tag_Tank16_seq
tag_Tank16_seq = HMIRuntime.Tags("Tank16_seq").Read
strSQL = "INSERT INTO z_tag_Tank16_seq (tag_value, created) values(" & tag_Tank16_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank16_seq = Nothing
Dim tag_Tank17_auto
tag_Tank17_auto = HMIRuntime.Tags("Tank17_auto").Read
strSQL = "INSERT INTO z_tag_Tank17_auto (tag_value, created) values(" & tag_Tank17_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank17_auto = Nothing
Dim tag_Tank17_CIP
tag_Tank17_CIP = HMIRuntime.Tags("Tank17_CIP").Read
strSQL = "INSERT INTO z_tag_Tank17_CIP (tag_value, created) values(" & tag_Tank17_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank17_CIP = Nothing
Dim tag_Tank17_run
tag_Tank17_run = HMIRuntime.Tags("Tank17_run").Read
strSQL = "INSERT INTO z_tag_Tank17_run (tag_value, created) values(" & tag_Tank17_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank17_run = Nothing
Dim tag_Tank17_seq
tag_Tank17_seq = HMIRuntime.Tags("Tank17_seq").Read
strSQL = "INSERT INTO z_tag_Tank17_seq (tag_value, created) values(" & tag_Tank17_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank17_seq = Nothing
Dim tag_Tank18_auto
tag_Tank18_auto = HMIRuntime.Tags("Tank18_auto").Read
strSQL = "INSERT INTO z_tag_Tank18_auto (tag_value, created) values(" & tag_Tank18_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank18_auto = Nothing
Dim tag_Tank18_CIP
tag_Tank18_CIP = HMIRuntime.Tags("Tank18_CIP").Read
strSQL = "INSERT INTO z_tag_Tank18_CIP (tag_value, created) values(" & tag_Tank18_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank18_CIP = Nothing
Dim tag_Tank18_run
tag_Tank18_run = HMIRuntime.Tags("Tank18_run").Read
strSQL = "INSERT INTO z_tag_Tank18_run (tag_value, created) values(" & tag_Tank18_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank18_run = Nothing
Dim tag_Tank18_seq
tag_Tank18_seq = HMIRuntime.Tags("Tank18_seq").Read
strSQL = "INSERT INTO z_tag_Tank18_seq (tag_value, created) values(" & tag_Tank18_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank18_seq = Nothing
Dim tag_Tank19_auto
tag_Tank19_auto = HMIRuntime.Tags("Tank19_auto").Read
strSQL = "INSERT INTO z_tag_Tank19_auto (tag_value, created) values(" & tag_Tank19_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank19_auto = Nothing
Dim tag_Tank19_CIP
tag_Tank19_CIP = HMIRuntime.Tags("Tank19_CIP").Read
strSQL = "INSERT INTO z_tag_Tank19_CIP (tag_value, created) values(" & tag_Tank19_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank19_CIP = Nothing
Dim tag_Tank19_run
tag_Tank19_run = HMIRuntime.Tags("Tank19_run").Read
strSQL = "INSERT INTO z_tag_Tank19_run (tag_value, created) values(" & tag_Tank19_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank19_run = Nothing
Dim tag_Tank19_seq
tag_Tank19_seq = HMIRuntime.Tags("Tank19_seq").Read
strSQL = "INSERT INTO z_tag_Tank19_seq (tag_value, created) values(" & tag_Tank19_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank19_seq = Nothing
Dim tag_Tank20_auto
tag_Tank20_auto = HMIRuntime.Tags("Tank20_auto").Read
strSQL = "INSERT INTO z_tag_Tank20_auto (tag_value, created) values(" & tag_Tank20_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank20_auto = Nothing
Dim tag_Tank20_CIP
tag_Tank20_CIP = HMIRuntime.Tags("Tank20_CIP").Read
strSQL = "INSERT INTO z_tag_Tank20_CIP (tag_value, created) values(" & tag_Tank20_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank20_CIP = Nothing
Dim tag_Tank20_run
tag_Tank20_run = HMIRuntime.Tags("Tank20_run").Read
strSQL = "INSERT INTO z_tag_Tank20_run (tag_value, created) values(" & tag_Tank20_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank20_run = Nothing
Dim tag_Tank20_seq
tag_Tank20_seq = HMIRuntime.Tags("Tank20_seq").Read
strSQL = "INSERT INTO z_tag_Tank20_seq (tag_value, created) values(" & tag_Tank20_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank20_seq = Nothing
Dim tag_Tank21_auto
tag_Tank21_auto = HMIRuntime.Tags("Tank21_auto").Read
strSQL = "INSERT INTO z_tag_Tank21_auto (tag_value, created) values(" & tag_Tank21_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank21_auto = Nothing
Dim tag_Tank21_CIP
tag_Tank21_CIP = HMIRuntime.Tags("Tank21_CIP").Read
strSQL = "INSERT INTO z_tag_Tank21_CIP (tag_value, created) values(" & tag_Tank21_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank21_CIP = Nothing
Dim tag_Tank21_run
tag_Tank21_run = HMIRuntime.Tags("Tank21_run").Read
strSQL = "INSERT INTO z_tag_Tank21_run (tag_value, created) values(" & tag_Tank21_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank21_run = Nothing
Dim tag_Tank21_seq
tag_Tank21_seq = HMIRuntime.Tags("Tank21_seq").Read
strSQL = "INSERT INTO z_tag_Tank21_seq (tag_value, created) values(" & tag_Tank21_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank21_seq = Nothing
Dim tag_Tank22_auto
tag_Tank22_auto = HMIRuntime.Tags("Tank22_auto").Read
strSQL = "INSERT INTO z_tag_Tank22_auto (tag_value, created) values(" & tag_Tank22_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank22_auto = Nothing
Dim tag_Tank22_CIP
tag_Tank22_CIP = HMIRuntime.Tags("Tank22_CIP").Read
strSQL = "INSERT INTO z_tag_Tank22_CIP (tag_value, created) values(" & tag_Tank22_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank22_CIP = Nothing
Dim tag_Tank22_run
tag_Tank22_run = HMIRuntime.Tags("Tank22_run").Read
strSQL = "INSERT INTO z_tag_Tank22_run (tag_value, created) values(" & tag_Tank22_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank22_run = Nothing
Dim tag_Tank22_seq
tag_Tank22_seq = HMIRuntime.Tags("Tank22_seq").Read
strSQL = "INSERT INTO z_tag_Tank22_seq (tag_value, created) values(" & tag_Tank22_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank22_seq = Nothing
Dim tag_Tank23_auto
tag_Tank23_auto = HMIRuntime.Tags("Tank23_auto").Read
strSQL = "INSERT INTO z_tag_Tank23_auto (tag_value, created) values(" & tag_Tank23_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank23_auto = Nothing
Dim tag_Tank23_CIP
tag_Tank23_CIP = HMIRuntime.Tags("Tank23_CIP").Read
strSQL = "INSERT INTO z_tag_Tank23_CIP (tag_value, created) values(" & tag_Tank23_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank23_CIP = Nothing
Dim tag_Tank23_run
tag_Tank23_run = HMIRuntime.Tags("Tank23_run").Read
strSQL = "INSERT INTO z_tag_Tank23_run (tag_value, created) values(" & tag_Tank23_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank23_run = Nothing
Dim tag_Tank23_seq
tag_Tank23_seq = HMIRuntime.Tags("Tank23_seq").Read
strSQL = "INSERT INTO z_tag_Tank23_seq (tag_value, created) values(" & tag_Tank23_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank23_seq = Nothing
Dim tag_Tank24_auto
tag_Tank24_auto = HMIRuntime.Tags("Tank24_auto").Read
strSQL = "INSERT INTO z_tag_Tank24_auto (tag_value, created) values(" & tag_Tank24_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank24_auto = Nothing
Dim tag_Tank24_CIP
tag_Tank24_CIP = HMIRuntime.Tags("Tank24_CIP").Read
strSQL = "INSERT INTO z_tag_Tank24_CIP (tag_value, created) values(" & tag_Tank24_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank24_CIP = Nothing
Dim tag_Tank24_run
tag_Tank24_run = HMIRuntime.Tags("Tank24_run").Read
strSQL = "INSERT INTO z_tag_Tank24_run (tag_value, created) values(" & tag_Tank24_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank24_run = Nothing
Dim tag_Tank24_seq
tag_Tank24_seq = HMIRuntime.Tags("Tank24_seq").Read
strSQL = "INSERT INTO z_tag_Tank24_seq (tag_value, created) values(" & tag_Tank24_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank24_seq = Nothing
Dim tag_Tank25_auto
tag_Tank25_auto = HMIRuntime.Tags("Tank25_auto").Read
strSQL = "INSERT INTO z_tag_Tank25_auto (tag_value, created) values(" & tag_Tank25_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank25_auto = Nothing
Dim tag_Tank25_CIP
tag_Tank25_CIP = HMIRuntime.Tags("Tank25_CIP").Read
strSQL = "INSERT INTO z_tag_Tank25_CIP (tag_value, created) values(" & tag_Tank25_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank25_CIP = Nothing
Dim tag_Tank25_run
tag_Tank25_run = HMIRuntime.Tags("Tank25_run").Read
strSQL = "INSERT INTO z_tag_Tank25_run (tag_value, created) values(" & tag_Tank25_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank25_run = Nothing
Dim tag_Tank25_seq
tag_Tank25_seq = HMIRuntime.Tags("Tank25_seq").Read
strSQL = "INSERT INTO z_tag_Tank25_seq (tag_value, created) values(" & tag_Tank25_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank25_seq = Nothing
Dim tag_Tank26_auto
tag_Tank26_auto = HMIRuntime.Tags("Tank26_auto").Read
strSQL = "INSERT INTO z_tag_Tank26_auto (tag_value, created) values(" & tag_Tank26_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank26_auto = Nothing
Dim tag_Tank26_CIP
tag_Tank26_CIP = HMIRuntime.Tags("Tank26_CIP").Read
strSQL = "INSERT INTO z_tag_Tank26_CIP (tag_value, created) values(" & tag_Tank26_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank26_CIP = Nothing
Dim tag_Tank26_run
tag_Tank26_run = HMIRuntime.Tags("Tank26_run").Read
strSQL = "INSERT INTO z_tag_Tank26_run (tag_value, created) values(" & tag_Tank26_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank26_run = Nothing
Dim tag_Tank26_seq
tag_Tank26_seq = HMIRuntime.Tags("Tank26_seq").Read
strSQL = "INSERT INTO z_tag_Tank26_seq (tag_value, created) values(" & tag_Tank26_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank26_seq = Nothing
Dim tag_Tank27_auto
tag_Tank27_auto = HMIRuntime.Tags("Tank27_auto").Read
strSQL = "INSERT INTO z_tag_Tank27_auto (tag_value, created) values(" & tag_Tank27_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank27_auto = Nothing
Dim tag_Tank27_CIP
tag_Tank27_CIP = HMIRuntime.Tags("Tank27_CIP").Read
strSQL = "INSERT INTO z_tag_Tank27_CIP (tag_value, created) values(" & tag_Tank27_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank27_CIP = Nothing
Dim tag_Tank27_run
tag_Tank27_run = HMIRuntime.Tags("Tank27_run").Read
strSQL = "INSERT INTO z_tag_Tank27_run (tag_value, created) values(" & tag_Tank27_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank27_run = Nothing
Dim tag_Tank27_seq
tag_Tank27_seq = HMIRuntime.Tags("Tank27_seq").Read
strSQL = "INSERT INTO z_tag_Tank27_seq (tag_value, created) values(" & tag_Tank27_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank27_seq = Nothing
Dim tag_Tank28_auto
tag_Tank28_auto = HMIRuntime.Tags("Tank28_auto").Read
strSQL = "INSERT INTO z_tag_Tank28_auto (tag_value, created) values(" & tag_Tank28_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank28_auto = Nothing
Dim tag_Tank28_CIP
tag_Tank28_CIP = HMIRuntime.Tags("Tank28_CIP").Read
strSQL = "INSERT INTO z_tag_Tank28_CIP (tag_value, created) values(" & tag_Tank28_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank28_CIP = Nothing
Dim tag_Tank28_run
tag_Tank28_run = HMIRuntime.Tags("Tank28_run").Read
strSQL = "INSERT INTO z_tag_Tank28_run (tag_value, created) values(" & tag_Tank28_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank28_run = Nothing
Dim tag_Tank28_seq
tag_Tank28_seq = HMIRuntime.Tags("Tank28_seq").Read
strSQL = "INSERT INTO z_tag_Tank28_seq (tag_value, created) values(" & tag_Tank28_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank28_seq = Nothing
Dim tag_Tank29_auto
tag_Tank29_auto = HMIRuntime.Tags("Tank29_auto").Read
strSQL = "INSERT INTO z_tag_Tank29_auto (tag_value, created) values(" & tag_Tank29_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank29_auto = Nothing
Dim tag_Tank29_CIP
tag_Tank29_CIP = HMIRuntime.Tags("Tank29_CIP").Read
strSQL = "INSERT INTO z_tag_Tank29_CIP (tag_value, created) values(" & tag_Tank29_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank29_CIP = Nothing
Dim tag_Tank29_run
tag_Tank29_run = HMIRuntime.Tags("Tank29_run").Read
strSQL = "INSERT INTO z_tag_Tank29_run (tag_value, created) values(" & tag_Tank29_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank29_run = Nothing
Dim tag_Tank29_seq
tag_Tank29_seq = HMIRuntime.Tags("Tank29_seq").Read
strSQL = "INSERT INTO z_tag_Tank29_seq (tag_value, created) values(" & tag_Tank29_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank29_seq = Nothing
Dim tag_Tank30_auto
tag_Tank30_auto = HMIRuntime.Tags("Tank30_auto").Read
strSQL = "INSERT INTO z_tag_Tank30_auto (tag_value, created) values(" & tag_Tank30_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank30_auto = Nothing
Dim tag_Tank30_CIP
tag_Tank30_CIP = HMIRuntime.Tags("Tank30_CIP").Read
strSQL = "INSERT INTO z_tag_Tank30_CIP (tag_value, created) values(" & tag_Tank30_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank30_CIP = Nothing
Dim tag_Tank30_run
tag_Tank30_run = HMIRuntime.Tags("Tank30_run").Read
strSQL = "INSERT INTO z_tag_Tank30_run (tag_value, created) values(" & tag_Tank30_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank30_run = Nothing
Dim tag_Tank30_seq
tag_Tank30_seq = HMIRuntime.Tags("Tank30_seq").Read
strSQL = "INSERT INTO z_tag_Tank30_seq (tag_value, created) values(" & tag_Tank30_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank30_seq = Nothing
Dim tag_Tank31_CIP
tag_Tank31_CIP = HMIRuntime.Tags("Tank31_CIP").Read
strSQL = "INSERT INTO z_tag_Tank31_CIP (tag_value, created) values(" & tag_Tank31_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank31_CIP = Nothing
Dim tag_Tank31_seq
tag_Tank31_seq = HMIRuntime.Tags("Tank31_seq").Read
strSQL = "INSERT INTO z_tag_Tank31_seq (tag_value, created) values(" & tag_Tank31_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank31_seq = Nothing
Dim tag_Tank32_CIP
tag_Tank32_CIP = HMIRuntime.Tags("Tank32_CIP").Read
strSQL = "INSERT INTO z_tag_Tank32_CIP (tag_value, created) values(" & tag_Tank32_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank32_CIP = Nothing
Dim tag_Tank32_seq
tag_Tank32_seq = HMIRuntime.Tags("Tank32_seq").Read
strSQL = "INSERT INTO z_tag_Tank32_seq (tag_value, created) values(" & tag_Tank32_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank32_seq = Nothing
Dim tag_Tank33_CIP
tag_Tank33_CIP = HMIRuntime.Tags("Tank33_CIP").Read
strSQL = "INSERT INTO z_tag_Tank33_CIP (tag_value, created) values(" & tag_Tank33_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank33_CIP = Nothing
Dim tag_Tank33_seq
tag_Tank33_seq = HMIRuntime.Tags("Tank33_seq").Read
strSQL = "INSERT INTO z_tag_Tank33_seq (tag_value, created) values(" & tag_Tank33_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank33_seq = Nothing
Dim tag_Tank34_CIP
tag_Tank34_CIP = HMIRuntime.Tags("Tank34_CIP").Read
strSQL = "INSERT INTO z_tag_Tank34_CIP (tag_value, created) values(" & tag_Tank34_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank34_CIP = Nothing
Dim tag_Tank34_seq
tag_Tank34_seq = HMIRuntime.Tags("Tank34_seq").Read
strSQL = "INSERT INTO z_tag_Tank34_seq (tag_value, created) values(" & tag_Tank34_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Tank34_seq = Nothing
Dim tag_TBF_Volum_Flow
tag_TBF_Volum_Flow = HMIRuntime.Tags("TBF_Volum_Flow").Read
strSQL = "INSERT INTO z_tag_TBF_Volum_Flow (tag_value, created) values(" & tag_TBF_Volum_Flow & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF_Volum_Flow = Nothing
Dim tag_TBF01_auto
tag_TBF01_auto = HMIRuntime.Tags("TBF01_auto").Read
strSQL = "INSERT INTO z_tag_TBF01_auto (tag_value, created) values(" & tag_TBF01_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_auto = Nothing
Dim tag_TBF01_CIP
tag_TBF01_CIP = HMIRuntime.Tags("TBF01_CIP").Read
strSQL = "INSERT INTO z_tag_TBF01_CIP (tag_value, created) values(" & tag_TBF01_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_CIP = Nothing
Dim tag_TBF01_run
tag_TBF01_run = HMIRuntime.Tags("TBF01_run").Read
strSQL = "INSERT INTO z_tag_TBF01_run (tag_value, created) values(" & tag_TBF01_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_run = Nothing
Dim tag_TBF01_seq
tag_TBF01_seq = HMIRuntime.Tags("TBF01_seq").Read
strSQL = "INSERT INTO z_tag_TBF01_seq (tag_value, created) values(" & tag_TBF01_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_seq = Nothing
Dim tag_TBF01_Sequence
tag_TBF01_Sequence = HMIRuntime.Tags("TBF01_Sequence").Read
strSQL = "INSERT INTO z_tag_TBF01_Sequence (tag_value, created) values(" & tag_TBF01_Sequence & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_Sequence = Nothing
Dim tag_TBF01_Volum
tag_TBF01_Volum = HMIRuntime.Tags("TBF01_Volum").Read
strSQL = "INSERT INTO z_tag_TBF01_Volum (tag_value, created) values(" & tag_TBF01_Volum & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF01_Volum = Nothing
Dim tag_TBF02_auto
tag_TBF02_auto = HMIRuntime.Tags("TBF02_auto").Read
strSQL = "INSERT INTO z_tag_TBF02_auto (tag_value, created) values(" & tag_TBF02_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_auto = Nothing
Dim tag_TBF02_CIP
tag_TBF02_CIP = HMIRuntime.Tags("TBF02_CIP").Read
strSQL = "INSERT INTO z_tag_TBF02_CIP (tag_value, created) values(" & tag_TBF02_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_CIP = Nothing
Dim tag_TBF02_run
tag_TBF02_run = HMIRuntime.Tags("TBF02_run").Read
strSQL = "INSERT INTO z_tag_TBF02_run (tag_value, created) values(" & tag_TBF02_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_run = Nothing
Dim tag_TBF02_seq
tag_TBF02_seq = HMIRuntime.Tags("TBF02_seq").Read
strSQL = "INSERT INTO z_tag_TBF02_seq (tag_value, created) values(" & tag_TBF02_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_seq = Nothing
Dim tag_TBF02_Sequence
tag_TBF02_Sequence = HMIRuntime.Tags("TBF02_Sequence").Read
strSQL = "INSERT INTO z_tag_TBF02_Sequence (tag_value, created) values(" & tag_TBF02_Sequence & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_Sequence = Nothing
Dim tag_TBF02_Volum
tag_TBF02_Volum = HMIRuntime.Tags("TBF02_Volum").Read
strSQL = "INSERT INTO z_tag_TBF02_Volum (tag_value, created) values(" & tag_TBF02_Volum & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF02_Volum = Nothing
Dim tag_TBF03_auto
tag_TBF03_auto = HMIRuntime.Tags("TBF03_auto").Read
strSQL = "INSERT INTO z_tag_TBF03_auto (tag_value, created) values(" & tag_TBF03_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_auto = Nothing
Dim tag_TBF03_CIP
tag_TBF03_CIP = HMIRuntime.Tags("TBF03_CIP").Read
strSQL = "INSERT INTO z_tag_TBF03_CIP (tag_value, created) values(" & tag_TBF03_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_CIP = Nothing
Dim tag_TBF03_run
tag_TBF03_run = HMIRuntime.Tags("TBF03_run").Read
strSQL = "INSERT INTO z_tag_TBF03_run (tag_value, created) values(" & tag_TBF03_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_run = Nothing
Dim tag_TBF03_seq
tag_TBF03_seq = HMIRuntime.Tags("TBF03_seq").Read
strSQL = "INSERT INTO z_tag_TBF03_seq (tag_value, created) values(" & tag_TBF03_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_seq = Nothing
Dim tag_TBF03_Sequence
tag_TBF03_Sequence = HMIRuntime.Tags("TBF03_Sequence").Read
strSQL = "INSERT INTO z_tag_TBF03_Sequence (tag_value, created) values(" & tag_TBF03_Sequence & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_Sequence = Nothing
Dim tag_TBF03_Volum
tag_TBF03_Volum = HMIRuntime.Tags("TBF03_Volum").Read
strSQL = "INSERT INTO z_tag_TBF03_Volum (tag_value, created) values(" & tag_TBF03_Volum & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF03_Volum = Nothing
Dim tag_TBF04_auto
tag_TBF04_auto = HMIRuntime.Tags("TBF04_auto").Read
strSQL = "INSERT INTO z_tag_TBF04_auto (tag_value, created) values(" & tag_TBF04_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_auto = Nothing
Dim tag_TBF04_CIP
tag_TBF04_CIP = HMIRuntime.Tags("TBF04_CIP").Read
strSQL = "INSERT INTO z_tag_TBF04_CIP (tag_value, created) values(" & tag_TBF04_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_CIP = Nothing
Dim tag_TBF04_run
tag_TBF04_run = HMIRuntime.Tags("TBF04_run").Read
strSQL = "INSERT INTO z_tag_TBF04_run (tag_value, created) values(" & tag_TBF04_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_run = Nothing
Dim tag_TBF04_seq
tag_TBF04_seq = HMIRuntime.Tags("TBF04_seq").Read
strSQL = "INSERT INTO z_tag_TBF04_seq (tag_value, created) values(" & tag_TBF04_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_seq = Nothing
Dim tag_TBF04_Sequence
tag_TBF04_Sequence = HMIRuntime.Tags("TBF04_Sequence").Read
strSQL = "INSERT INTO z_tag_TBF04_Sequence (tag_value, created) values(" & tag_TBF04_Sequence & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_Sequence = Nothing
Dim tag_TBF04_Volum
tag_TBF04_Volum = HMIRuntime.Tags("TBF04_Volum").Read
strSQL = "INSERT INTO z_tag_TBF04_Volum (tag_value, created) values(" & tag_TBF04_Volum & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF04_Volum = Nothing
Dim tag_TBF05_CIP
tag_TBF05_CIP = HMIRuntime.Tags("TBF05_CIP").Read
strSQL = "INSERT INTO z_tag_TBF05_CIP (tag_value, created) values(" & tag_TBF05_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF05_CIP = Nothing
Dim tag_TBF06_CIP
tag_TBF06_CIP = HMIRuntime.Tags("TBF06_CIP").Read
strSQL = "INSERT INTO z_tag_TBF06_CIP (tag_value, created) values(" & tag_TBF06_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF06_CIP = Nothing
Dim tag_TBF07_Total
tag_TBF07_Total = HMIRuntime.Tags("TBF07_Total").Read
strSQL = "INSERT INTO z_tag_TBF07_Total (tag_value, created) values(" & tag_TBF07_Total & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF07_Total = Nothing
Dim tag_TBF08_Total
tag_TBF08_Total = HMIRuntime.Tags("TBF08_Total").Read
strSQL = "INSERT INTO z_tag_TBF08_Total (tag_value, created) values(" & tag_TBF08_Total & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF08_Total = Nothing
Dim tag_TBF09_TET
tag_TBF09_TET = HMIRuntime.Tags("TBF09_TET").Read
strSQL = "INSERT INTO z_tag_TBF09_TET (tag_value, created) values(" & tag_TBF09_TET & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF09_TET = Nothing
Dim tag_TBF09_Total
tag_TBF09_Total = HMIRuntime.Tags("TBF09_Total").Read
strSQL = "INSERT INTO z_tag_TBF09_Total (tag_value, created) values(" & tag_TBF09_Total & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBF09_Total = Nothing
Dim tag_TBFNEW_CIP
tag_TBFNEW_CIP = HMIRuntime.Tags("TBFNEW_CIP").Read
strSQL = "INSERT INTO z_tag_TBFNEW_CIP (tag_value, created) values(" & tag_TBFNEW_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TBFNEW_CIP = Nothing
Dim tag_TET_DA01
tag_TET_DA01 = HMIRuntime.Tags("TET_DA01").Read
strSQL = "INSERT INTO z_tag_TET_DA01 (tag_value, created) values(" & tag_TET_DA01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_DA01 = Nothing
Dim tag_TET_DA02
tag_TET_DA02 = HMIRuntime.Tags("TET_DA02").Read
strSQL = "INSERT INTO z_tag_TET_DA02 (tag_value, created) values(" & tag_TET_DA02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_DA02 = Nothing
Dim tag_TET_TACHMEN70hl
tag_TET_TACHMEN70hl = HMIRuntime.Tags("TET_TACHMEN70hl").Read
strSQL = "INSERT INTO z_tag_TET_TACHMEN70hl (tag_value, created) values(" & tag_TET_TACHMEN70hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_TACHMEN70hl = Nothing
Dim tag_TET_TACHMEN100hl
tag_TET_TACHMEN100hl = HMIRuntime.Tags("TET_TACHMEN100hl").Read
strSQL = "INSERT INTO z_tag_TET_TACHMEN100hl (tag_value, created) values(" & tag_TET_TACHMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_TACHMEN100hl = Nothing
Dim tag_TET_TANKMEN100hl
tag_TET_TANKMEN100hl = HMIRuntime.Tags("TET_TANKMEN100hl").Read
strSQL = "INSERT INTO z_tag_TET_TANKMEN100hl (tag_value, created) values(" & tag_TET_TANKMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_TANKMEN100hl = Nothing
Dim tag_TET_TBF08_IN
tag_TET_TBF08_IN = HMIRuntime.Tags("TET_TBF08_IN").Read
strSQL = "INSERT INTO z_tag_TET_TBF08_IN (tag_value, created) values(" & tag_TET_TBF08_IN & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TET_TBF08_IN = Nothing
Dim tag_thoigianda_T1
tag_thoigianda_T1 = HMIRuntime.Tags("thoigianda_T1").Read
strSQL = "INSERT INTO z_tag_thoigianda_T1 (tag_value, created) values(" & tag_thoigianda_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T1 = Nothing
Dim tag_thoigianda_T2
tag_thoigianda_T2 = HMIRuntime.Tags("thoigianda_T2").Read
strSQL = "INSERT INTO z_tag_thoigianda_T2 (tag_value, created) values(" & tag_thoigianda_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T2 = Nothing
Dim tag_thoigianda_T3
tag_thoigianda_T3 = HMIRuntime.Tags("thoigianda_T3").Read
strSQL = "INSERT INTO z_tag_thoigianda_T3 (tag_value, created) values(" & tag_thoigianda_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T3 = Nothing
Dim tag_thoigianda_T4
tag_thoigianda_T4 = HMIRuntime.Tags("thoigianda_T4").Read
strSQL = "INSERT INTO z_tag_thoigianda_T4 (tag_value, created) values(" & tag_thoigianda_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T4 = Nothing
Dim tag_thoigianda_T5
tag_thoigianda_T5 = HMIRuntime.Tags("thoigianda_T5").Read
strSQL = "INSERT INTO z_tag_thoigianda_T5 (tag_value, created) values(" & tag_thoigianda_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T5 = Nothing
Dim tag_thoigianda_T6
tag_thoigianda_T6 = HMIRuntime.Tags("thoigianda_T6").Read
strSQL = "INSERT INTO z_tag_thoigianda_T6 (tag_value, created) values(" & tag_thoigianda_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T6 = Nothing
Dim tag_thoigianda_T7
tag_thoigianda_T7 = HMIRuntime.Tags("thoigianda_T7").Read
strSQL = "INSERT INTO z_tag_thoigianda_T7 (tag_value, created) values(" & tag_thoigianda_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T7 = Nothing
Dim tag_thoigianda_T8
tag_thoigianda_T8 = HMIRuntime.Tags("thoigianda_T8").Read
strSQL = "INSERT INTO z_tag_thoigianda_T8 (tag_value, created) values(" & tag_thoigianda_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T8 = Nothing
Dim tag_thoigianda_T9
tag_thoigianda_T9 = HMIRuntime.Tags("thoigianda_T9").Read
strSQL = "INSERT INTO z_tag_thoigianda_T9 (tag_value, created) values(" & tag_thoigianda_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T9 = Nothing
Dim tag_thoigianda_T10
tag_thoigianda_T10 = HMIRuntime.Tags("thoigianda_T10").Read
strSQL = "INSERT INTO z_tag_thoigianda_T10 (tag_value, created) values(" & tag_thoigianda_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T10 = Nothing
Dim tag_thoigianda_T11
tag_thoigianda_T11 = HMIRuntime.Tags("thoigianda_T11").Read
strSQL = "INSERT INTO z_tag_thoigianda_T11 (tag_value, created) values(" & tag_thoigianda_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T11 = Nothing
Dim tag_thoigianda_T12
tag_thoigianda_T12 = HMIRuntime.Tags("thoigianda_T12").Read
strSQL = "INSERT INTO z_tag_thoigianda_T12 (tag_value, created) values(" & tag_thoigianda_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T12 = Nothing
Dim tag_thoigianda_T13
tag_thoigianda_T13 = HMIRuntime.Tags("thoigianda_T13").Read
strSQL = "INSERT INTO z_tag_thoigianda_T13 (tag_value, created) values(" & tag_thoigianda_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T13 = Nothing
Dim tag_thoigianda_T14
tag_thoigianda_T14 = HMIRuntime.Tags("thoigianda_T14").Read
strSQL = "INSERT INTO z_tag_thoigianda_T14 (tag_value, created) values(" & tag_thoigianda_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T14 = Nothing
Dim tag_thoigianda_T15
tag_thoigianda_T15 = HMIRuntime.Tags("thoigianda_T15").Read
strSQL = "INSERT INTO z_tag_thoigianda_T15 (tag_value, created) values(" & tag_thoigianda_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T15 = Nothing
Dim tag_thoigianda_T16
tag_thoigianda_T16 = HMIRuntime.Tags("thoigianda_T16").Read
strSQL = "INSERT INTO z_tag_thoigianda_T16 (tag_value, created) values(" & tag_thoigianda_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T16 = Nothing
Dim tag_thoigianda_T17
tag_thoigianda_T17 = HMIRuntime.Tags("thoigianda_T17").Read
strSQL = "INSERT INTO z_tag_thoigianda_T17 (tag_value, created) values(" & tag_thoigianda_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T17 = Nothing
Dim tag_thoigianda_T18
tag_thoigianda_T18 = HMIRuntime.Tags("thoigianda_T18").Read
strSQL = "INSERT INTO z_tag_thoigianda_T18 (tag_value, created) values(" & tag_thoigianda_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T18 = Nothing
Dim tag_thoigianda_T19
tag_thoigianda_T19 = HMIRuntime.Tags("thoigianda_T19").Read
strSQL = "INSERT INTO z_tag_thoigianda_T19 (tag_value, created) values(" & tag_thoigianda_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T19 = Nothing
Dim tag_thoigianda_T20
tag_thoigianda_T20 = HMIRuntime.Tags("thoigianda_T20").Read
strSQL = "INSERT INTO z_tag_thoigianda_T20 (tag_value, created) values(" & tag_thoigianda_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T20 = Nothing
Dim tag_thoigianda_T21
tag_thoigianda_T21 = HMIRuntime.Tags("thoigianda_T21").Read
strSQL = "INSERT INTO z_tag_thoigianda_T21 (tag_value, created) values(" & tag_thoigianda_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T21 = Nothing
Dim tag_thoigianda_T22
tag_thoigianda_T22 = HMIRuntime.Tags("thoigianda_T22").Read
strSQL = "INSERT INTO z_tag_thoigianda_T22 (tag_value, created) values(" & tag_thoigianda_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T22 = Nothing
Dim tag_thoigianda_T23
tag_thoigianda_T23 = HMIRuntime.Tags("thoigianda_T23").Read
strSQL = "INSERT INTO z_tag_thoigianda_T23 (tag_value, created) values(" & tag_thoigianda_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T23 = Nothing
Dim tag_thoigianda_T24
tag_thoigianda_T24 = HMIRuntime.Tags("thoigianda_T24").Read
strSQL = "INSERT INTO z_tag_thoigianda_T24 (tag_value, created) values(" & tag_thoigianda_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T24 = Nothing
Dim tag_thoigianda_T25
tag_thoigianda_T25 = HMIRuntime.Tags("thoigianda_T25").Read
strSQL = "INSERT INTO z_tag_thoigianda_T25 (tag_value, created) values(" & tag_thoigianda_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T25 = Nothing
Dim tag_thoigianda_T26
tag_thoigianda_T26 = HMIRuntime.Tags("thoigianda_T26").Read
strSQL = "INSERT INTO z_tag_thoigianda_T26 (tag_value, created) values(" & tag_thoigianda_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T26 = Nothing
Dim tag_thoigianda_T27
tag_thoigianda_T27 = HMIRuntime.Tags("thoigianda_T27").Read
strSQL = "INSERT INTO z_tag_thoigianda_T27 (tag_value, created) values(" & tag_thoigianda_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T27 = Nothing
Dim tag_thoigianda_T28
tag_thoigianda_T28 = HMIRuntime.Tags("thoigianda_T28").Read
strSQL = "INSERT INTO z_tag_thoigianda_T28 (tag_value, created) values(" & tag_thoigianda_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T28 = Nothing
Dim tag_thoigianda_T29
tag_thoigianda_T29 = HMIRuntime.Tags("thoigianda_T29").Read
strSQL = "INSERT INTO z_tag_thoigianda_T29 (tag_value, created) values(" & tag_thoigianda_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T29 = Nothing
Dim tag_thoigianda_T30
tag_thoigianda_T30 = HMIRuntime.Tags("thoigianda_T30").Read
strSQL = "INSERT INTO z_tag_thoigianda_T30 (tag_value, created) values(" & tag_thoigianda_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianda_T30 = Nothing
Dim tag_thoigianhnd_T1
tag_thoigianhnd_T1 = HMIRuntime.Tags("thoigianhnd_T1").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T1 (tag_value, created) values(" & tag_thoigianhnd_T1 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T1 = Nothing
Dim tag_thoigianhnd_T2
tag_thoigianhnd_T2 = HMIRuntime.Tags("thoigianhnd_T2").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T2 (tag_value, created) values(" & tag_thoigianhnd_T2 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T2 = Nothing
Dim tag_thoigianhnd_T3
tag_thoigianhnd_T3 = HMIRuntime.Tags("thoigianhnd_T3").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T3 (tag_value, created) values(" & tag_thoigianhnd_T3 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T3 = Nothing
Dim tag_thoigianhnd_T4
tag_thoigianhnd_T4 = HMIRuntime.Tags("thoigianhnd_T4").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T4 (tag_value, created) values(" & tag_thoigianhnd_T4 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T4 = Nothing
Dim tag_thoigianhnd_T5
tag_thoigianhnd_T5 = HMIRuntime.Tags("thoigianhnd_T5").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T5 (tag_value, created) values(" & tag_thoigianhnd_T5 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T5 = Nothing
Dim tag_thoigianhnd_T6
tag_thoigianhnd_T6 = HMIRuntime.Tags("thoigianhnd_T6").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T6 (tag_value, created) values(" & tag_thoigianhnd_T6 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T6 = Nothing
Dim tag_thoigianhnd_T7
tag_thoigianhnd_T7 = HMIRuntime.Tags("thoigianhnd_T7").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T7 (tag_value, created) values(" & tag_thoigianhnd_T7 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T7 = Nothing
Dim tag_thoigianhnd_T8
tag_thoigianhnd_T8 = HMIRuntime.Tags("thoigianhnd_T8").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T8 (tag_value, created) values(" & tag_thoigianhnd_T8 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T8 = Nothing
Dim tag_thoigianhnd_T9
tag_thoigianhnd_T9 = HMIRuntime.Tags("thoigianhnd_T9").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T9 (tag_value, created) values(" & tag_thoigianhnd_T9 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T9 = Nothing
Dim tag_thoigianhnd_T10
tag_thoigianhnd_T10 = HMIRuntime.Tags("thoigianhnd_T10").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T10 (tag_value, created) values(" & tag_thoigianhnd_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T10 = Nothing
Dim tag_thoigianhnd_T11
tag_thoigianhnd_T11 = HMIRuntime.Tags("thoigianhnd_T11").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T11 (tag_value, created) values(" & tag_thoigianhnd_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T11 = Nothing
Dim tag_thoigianhnd_T12
tag_thoigianhnd_T12 = HMIRuntime.Tags("thoigianhnd_T12").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T12 (tag_value, created) values(" & tag_thoigianhnd_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T12 = Nothing
Dim tag_thoigianhnd_T13
tag_thoigianhnd_T13 = HMIRuntime.Tags("thoigianhnd_T13").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T13 (tag_value, created) values(" & tag_thoigianhnd_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T13 = Nothing
Dim tag_thoigianhnd_T14
tag_thoigianhnd_T14 = HMIRuntime.Tags("thoigianhnd_T14").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T14 (tag_value, created) values(" & tag_thoigianhnd_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T14 = Nothing
Dim tag_thoigianhnd_T15
tag_thoigianhnd_T15 = HMIRuntime.Tags("thoigianhnd_T15").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T15 (tag_value, created) values(" & tag_thoigianhnd_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T15 = Nothing
Dim tag_thoigianhnd_T16
tag_thoigianhnd_T16 = HMIRuntime.Tags("thoigianhnd_T16").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T16 (tag_value, created) values(" & tag_thoigianhnd_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T16 = Nothing
Dim tag_thoigianhnd_T17
tag_thoigianhnd_T17 = HMIRuntime.Tags("thoigianhnd_T17").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T17 (tag_value, created) values(" & tag_thoigianhnd_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T17 = Nothing
Dim tag_thoigianhnd_T18
tag_thoigianhnd_T18 = HMIRuntime.Tags("thoigianhnd_T18").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T18 (tag_value, created) values(" & tag_thoigianhnd_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T18 = Nothing
Dim tag_thoigianhnd_T19
tag_thoigianhnd_T19 = HMIRuntime.Tags("thoigianhnd_T19").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T19 (tag_value, created) values(" & tag_thoigianhnd_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T19 = Nothing
Dim tag_thoigianhnd_T20
tag_thoigianhnd_T20 = HMIRuntime.Tags("thoigianhnd_T20").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T20 (tag_value, created) values(" & tag_thoigianhnd_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T20 = Nothing
Dim tag_thoigianhnd_T21
tag_thoigianhnd_T21 = HMIRuntime.Tags("thoigianhnd_T21").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T21 (tag_value, created) values(" & tag_thoigianhnd_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T21 = Nothing
Dim tag_thoigianhnd_T22
tag_thoigianhnd_T22 = HMIRuntime.Tags("thoigianhnd_T22").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T22 (tag_value, created) values(" & tag_thoigianhnd_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T22 = Nothing
Dim tag_thoigianhnd_T23
tag_thoigianhnd_T23 = HMIRuntime.Tags("thoigianhnd_T23").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T23 (tag_value, created) values(" & tag_thoigianhnd_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T23 = Nothing
Dim tag_thoigianhnd_T24
tag_thoigianhnd_T24 = HMIRuntime.Tags("thoigianhnd_T24").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T24 (tag_value, created) values(" & tag_thoigianhnd_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T24 = Nothing
Dim tag_thoigianhnd_T25
tag_thoigianhnd_T25 = HMIRuntime.Tags("thoigianhnd_T25").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T25 (tag_value, created) values(" & tag_thoigianhnd_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T25 = Nothing
Dim tag_thoigianhnd_T26
tag_thoigianhnd_T26 = HMIRuntime.Tags("thoigianhnd_T26").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T26 (tag_value, created) values(" & tag_thoigianhnd_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T26 = Nothing
Dim tag_thoigianhnd_T27
tag_thoigianhnd_T27 = HMIRuntime.Tags("thoigianhnd_T27").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T27 (tag_value, created) values(" & tag_thoigianhnd_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T27 = Nothing
Dim tag_thoigianhnd_T28
tag_thoigianhnd_T28 = HMIRuntime.Tags("thoigianhnd_T28").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T28 (tag_value, created) values(" & tag_thoigianhnd_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T28 = Nothing
Dim tag_thoigianhnd_T29
tag_thoigianhnd_T29 = HMIRuntime.Tags("thoigianhnd_T29").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T29 (tag_value, created) values(" & tag_thoigianhnd_T29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T29 = Nothing
Dim tag_thoigianhnd_T30
tag_thoigianhnd_T30 = HMIRuntime.Tags("thoigianhnd_T30").Read
strSQL = "INSERT INTO z_tag_thoigianhnd_T30 (tag_value, created) values(" & tag_thoigianhnd_T30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianhnd_T30 = Nothing
Dim tag_thoigianrun_0701M03
tag_thoigianrun_0701M03 = HMIRuntime.Tags("thoigianrun_0701M03").Read
strSQL = "INSERT INTO z_tag_thoigianrun_0701M03 (tag_value, created) values(" & tag_thoigianrun_0701M03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianrun_0701M03 = Nothing
Dim tag_thoigianrun_0701M04
tag_thoigianrun_0701M04 = HMIRuntime.Tags("thoigianrun_0701M04").Read
strSQL = "INSERT INTO z_tag_thoigianrun_0701M04 (tag_value, created) values(" & tag_thoigianrun_0701M04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianrun_0701M04 = Nothing
Dim tag_thoigianrun_0703M01
tag_thoigianrun_0703M01 = HMIRuntime.Tags("thoigianrun_0703M01").Read
strSQL = "INSERT INTO z_tag_thoigianrun_0703M01 (tag_value, created) values(" & tag_thoigianrun_0703M01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianrun_0703M01 = Nothing
Dim tag_thoigianrun_0703M02
tag_thoigianrun_0703M02 = HMIRuntime.Tags("thoigianrun_0703M02").Read
strSQL = "INSERT INTO z_tag_thoigianrun_0703M02 (tag_value, created) values(" & tag_thoigianrun_0703M02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_thoigianrun_0703M02 = Nothing
Dim tag_Total_BBT01
tag_Total_BBT01 = HMIRuntime.Tags("Total_BBT01").Read
strSQL = "INSERT INTO z_tag_Total_BBT01 (tag_value, created) values(" & tag_Total_BBT01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Total_BBT01 = Nothing
Dim tag_Total_BBT02
tag_Total_BBT02 = HMIRuntime.Tags("Total_BBT02").Read
strSQL = "INSERT INTO z_tag_Total_BBT02 (tag_value, created) values(" & tag_Total_BBT02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Total_BBT02 = Nothing
Dim tag_Total_BBT03
tag_Total_BBT03 = HMIRuntime.Tags("Total_BBT03").Read
strSQL = "INSERT INTO z_tag_Total_BBT03 (tag_value, created) values(" & tag_Total_BBT03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Total_BBT03 = Nothing
Dim tag_Total_BBT04
tag_Total_BBT04 = HMIRuntime.Tags("Total_BBT04").Read
strSQL = "INSERT INTO z_tag_Total_BBT04 (tag_value, created) values(" & tag_Total_BBT04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Total_BBT04 = Nothing
Dim tag_TotalTank01
tag_TotalTank01 = HMIRuntime.Tags("TotalTank01").Read
strSQL = "INSERT INTO z_tag_TotalTank01 (tag_value, created) values(" & tag_TotalTank01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank01 = Nothing
Dim tag_TotalTank02
tag_TotalTank02 = HMIRuntime.Tags("TotalTank02").Read
strSQL = "INSERT INTO z_tag_TotalTank02 (tag_value, created) values(" & tag_TotalTank02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank02 = Nothing
Dim tag_TotalTank03
tag_TotalTank03 = HMIRuntime.Tags("TotalTank03").Read
strSQL = "INSERT INTO z_tag_TotalTank03 (tag_value, created) values(" & tag_TotalTank03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank03 = Nothing
Dim tag_TotalTank04
tag_TotalTank04 = HMIRuntime.Tags("TotalTank04").Read
strSQL = "INSERT INTO z_tag_TotalTank04 (tag_value, created) values(" & tag_TotalTank04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank04 = Nothing
Dim tag_TotalTank05
tag_TotalTank05 = HMIRuntime.Tags("TotalTank05").Read
strSQL = "INSERT INTO z_tag_TotalTank05 (tag_value, created) values(" & tag_TotalTank05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank05 = Nothing
Dim tag_TotalTank06
tag_TotalTank06 = HMIRuntime.Tags("TotalTank06").Read
strSQL = "INSERT INTO z_tag_TotalTank06 (tag_value, created) values(" & tag_TotalTank06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank06 = Nothing
Dim tag_TotalTank07
tag_TotalTank07 = HMIRuntime.Tags("TotalTank07").Read
strSQL = "INSERT INTO z_tag_TotalTank07 (tag_value, created) values(" & tag_TotalTank07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank07 = Nothing
Dim tag_TotalTank08
tag_TotalTank08 = HMIRuntime.Tags("TotalTank08").Read
strSQL = "INSERT INTO z_tag_TotalTank08 (tag_value, created) values(" & tag_TotalTank08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank08 = Nothing
Dim tag_TotalTank09
tag_TotalTank09 = HMIRuntime.Tags("TotalTank09").Read
strSQL = "INSERT INTO z_tag_TotalTank09 (tag_value, created) values(" & tag_TotalTank09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank09 = Nothing
Dim tag_TotalTank10
tag_TotalTank10 = HMIRuntime.Tags("TotalTank10").Read
strSQL = "INSERT INTO z_tag_TotalTank10 (tag_value, created) values(" & tag_TotalTank10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank10 = Nothing
Dim tag_TotalTank11
tag_TotalTank11 = HMIRuntime.Tags("TotalTank11").Read
strSQL = "INSERT INTO z_tag_TotalTank11 (tag_value, created) values(" & tag_TotalTank11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank11 = Nothing
Dim tag_TotalTank12
tag_TotalTank12 = HMIRuntime.Tags("TotalTank12").Read
strSQL = "INSERT INTO z_tag_TotalTank12 (tag_value, created) values(" & tag_TotalTank12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank12 = Nothing
Dim tag_TotalTank13
tag_TotalTank13 = HMIRuntime.Tags("TotalTank13").Read
strSQL = "INSERT INTO z_tag_TotalTank13 (tag_value, created) values(" & tag_TotalTank13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank13 = Nothing
Dim tag_TotalTank14
tag_TotalTank14 = HMIRuntime.Tags("TotalTank14").Read
strSQL = "INSERT INTO z_tag_TotalTank14 (tag_value, created) values(" & tag_TotalTank14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank14 = Nothing
Dim tag_TotalTank15
tag_TotalTank15 = HMIRuntime.Tags("TotalTank15").Read
strSQL = "INSERT INTO z_tag_TotalTank15 (tag_value, created) values(" & tag_TotalTank15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank15 = Nothing
Dim tag_TotalTank16
tag_TotalTank16 = HMIRuntime.Tags("TotalTank16").Read
strSQL = "INSERT INTO z_tag_TotalTank16 (tag_value, created) values(" & tag_TotalTank16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank16 = Nothing
Dim tag_TotalTank17
tag_TotalTank17 = HMIRuntime.Tags("TotalTank17").Read
strSQL = "INSERT INTO z_tag_TotalTank17 (tag_value, created) values(" & tag_TotalTank17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank17 = Nothing
Dim tag_TotalTank18
tag_TotalTank18 = HMIRuntime.Tags("TotalTank18").Read
strSQL = "INSERT INTO z_tag_TotalTank18 (tag_value, created) values(" & tag_TotalTank18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank18 = Nothing
Dim tag_TotalTank19
tag_TotalTank19 = HMIRuntime.Tags("TotalTank19").Read
strSQL = "INSERT INTO z_tag_TotalTank19 (tag_value, created) values(" & tag_TotalTank19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank19 = Nothing
Dim tag_TotalTank20
tag_TotalTank20 = HMIRuntime.Tags("TotalTank20").Read
strSQL = "INSERT INTO z_tag_TotalTank20 (tag_value, created) values(" & tag_TotalTank20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank20 = Nothing
Dim tag_TotalTank21
tag_TotalTank21 = HMIRuntime.Tags("TotalTank21").Read
strSQL = "INSERT INTO z_tag_TotalTank21 (tag_value, created) values(" & tag_TotalTank21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank21 = Nothing
Dim tag_TotalTank22
tag_TotalTank22 = HMIRuntime.Tags("TotalTank22").Read
strSQL = "INSERT INTO z_tag_TotalTank22 (tag_value, created) values(" & tag_TotalTank22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank22 = Nothing
Dim tag_TotalTank23
tag_TotalTank23 = HMIRuntime.Tags("TotalTank23").Read
strSQL = "INSERT INTO z_tag_TotalTank23 (tag_value, created) values(" & tag_TotalTank23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank23 = Nothing
Dim tag_TotalTank24
tag_TotalTank24 = HMIRuntime.Tags("TotalTank24").Read
strSQL = "INSERT INTO z_tag_TotalTank24 (tag_value, created) values(" & tag_TotalTank24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank24 = Nothing
Dim tag_TotalTank25
tag_TotalTank25 = HMIRuntime.Tags("TotalTank25").Read
strSQL = "INSERT INTO z_tag_TotalTank25 (tag_value, created) values(" & tag_TotalTank25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank25 = Nothing
Dim tag_TotalTank26
tag_TotalTank26 = HMIRuntime.Tags("TotalTank26").Read
strSQL = "INSERT INTO z_tag_TotalTank26 (tag_value, created) values(" & tag_TotalTank26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank26 = Nothing
Dim tag_TotalTank27
tag_TotalTank27 = HMIRuntime.Tags("TotalTank27").Read
strSQL = "INSERT INTO z_tag_TotalTank27 (tag_value, created) values(" & tag_TotalTank27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank27 = Nothing
Dim tag_TotalTank28
tag_TotalTank28 = HMIRuntime.Tags("TotalTank28").Read
strSQL = "INSERT INTO z_tag_TotalTank28 (tag_value, created) values(" & tag_TotalTank28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank28 = Nothing
Dim tag_TotalTank29
tag_TotalTank29 = HMIRuntime.Tags("TotalTank29").Read
strSQL = "INSERT INTO z_tag_TotalTank29 (tag_value, created) values(" & tag_TotalTank29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank29 = Nothing
Dim tag_TotalTank30
tag_TotalTank30 = HMIRuntime.Tags("TotalTank30").Read
strSQL = "INSERT INTO z_tag_TotalTank30 (tag_value, created) values(" & tag_TotalTank30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank30 = Nothing
Dim tag_TotalTank31
tag_TotalTank31 = HMIRuntime.Tags("TotalTank31").Read
strSQL = "INSERT INTO z_tag_TotalTank31 (tag_value, created) values(" & tag_TotalTank31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank31 = Nothing
Dim tag_TotalTank32
tag_TotalTank32 = HMIRuntime.Tags("TotalTank32").Read
strSQL = "INSERT INTO z_tag_TotalTank32 (tag_value, created) values(" & tag_TotalTank32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank32 = Nothing
Dim tag_TotalTank33
tag_TotalTank33 = HMIRuntime.Tags("TotalTank33").Read
strSQL = "INSERT INTO z_tag_TotalTank33 (tag_value, created) values(" & tag_TotalTank33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank33 = Nothing
Dim tag_TotalTank34
tag_TotalTank34 = HMIRuntime.Tags("TotalTank34").Read
strSQL = "INSERT INTO z_tag_TotalTank34 (tag_value, created) values(" & tag_TotalTank34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTank34 = Nothing
Dim tag_TotalTimeT01
tag_TotalTimeT01 = HMIRuntime.Tags("TotalTimeT01").Read
strSQL = "INSERT INTO z_tag_TotalTimeT01 (tag_value, created) values(" & tag_TotalTimeT01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT01 = Nothing
Dim tag_TotalTimeT02
tag_TotalTimeT02 = HMIRuntime.Tags("TotalTimeT02").Read
strSQL = "INSERT INTO z_tag_TotalTimeT02 (tag_value, created) values(" & tag_TotalTimeT02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT02 = Nothing
Dim tag_TotalTimeT03
tag_TotalTimeT03 = HMIRuntime.Tags("TotalTimeT03").Read
strSQL = "INSERT INTO z_tag_TotalTimeT03 (tag_value, created) values(" & tag_TotalTimeT03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT03 = Nothing
Dim tag_TotalTimeT04
tag_TotalTimeT04 = HMIRuntime.Tags("TotalTimeT04").Read
strSQL = "INSERT INTO z_tag_TotalTimeT04 (tag_value, created) values(" & tag_TotalTimeT04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT04 = Nothing
Dim tag_TotalTimeT05
tag_TotalTimeT05 = HMIRuntime.Tags("TotalTimeT05").Read
strSQL = "INSERT INTO z_tag_TotalTimeT05 (tag_value, created) values(" & tag_TotalTimeT05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT05 = Nothing
Dim tag_TotalTimeT06
tag_TotalTimeT06 = HMIRuntime.Tags("TotalTimeT06").Read
strSQL = "INSERT INTO z_tag_TotalTimeT06 (tag_value, created) values(" & tag_TotalTimeT06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT06 = Nothing
Dim tag_TotalTimeT07
tag_TotalTimeT07 = HMIRuntime.Tags("TotalTimeT07").Read
strSQL = "INSERT INTO z_tag_TotalTimeT07 (tag_value, created) values(" & tag_TotalTimeT07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT07 = Nothing
Dim tag_TotalTimeT08
tag_TotalTimeT08 = HMIRuntime.Tags("TotalTimeT08").Read
strSQL = "INSERT INTO z_tag_TotalTimeT08 (tag_value, created) values(" & tag_TotalTimeT08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT08 = Nothing
Dim tag_TotalTimeT09
tag_TotalTimeT09 = HMIRuntime.Tags("TotalTimeT09").Read
strSQL = "INSERT INTO z_tag_TotalTimeT09 (tag_value, created) values(" & tag_TotalTimeT09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT09 = Nothing
Dim tag_TotalTimeT10
tag_TotalTimeT10 = HMIRuntime.Tags("TotalTimeT10").Read
strSQL = "INSERT INTO z_tag_TotalTimeT10 (tag_value, created) values(" & tag_TotalTimeT10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT10 = Nothing
Dim tag_TotalTimeT11
tag_TotalTimeT11 = HMIRuntime.Tags("TotalTimeT11").Read
strSQL = "INSERT INTO z_tag_TotalTimeT11 (tag_value, created) values(" & tag_TotalTimeT11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT11 = Nothing
Dim tag_TotalTimeT12
tag_TotalTimeT12 = HMIRuntime.Tags("TotalTimeT12").Read
strSQL = "INSERT INTO z_tag_TotalTimeT12 (tag_value, created) values(" & tag_TotalTimeT12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT12 = Nothing
Dim tag_TotalTimeT13
tag_TotalTimeT13 = HMIRuntime.Tags("TotalTimeT13").Read
strSQL = "INSERT INTO z_tag_TotalTimeT13 (tag_value, created) values(" & tag_TotalTimeT13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT13 = Nothing
Dim tag_TotalTimeT14
tag_TotalTimeT14 = HMIRuntime.Tags("TotalTimeT14").Read
strSQL = "INSERT INTO z_tag_TotalTimeT14 (tag_value, created) values(" & tag_TotalTimeT14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT14 = Nothing
Dim tag_TotalTimeT15
tag_TotalTimeT15 = HMIRuntime.Tags("TotalTimeT15").Read
strSQL = "INSERT INTO z_tag_TotalTimeT15 (tag_value, created) values(" & tag_TotalTimeT15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT15 = Nothing
Dim tag_TotalTimeT16
tag_TotalTimeT16 = HMIRuntime.Tags("TotalTimeT16").Read
strSQL = "INSERT INTO z_tag_TotalTimeT16 (tag_value, created) values(" & tag_TotalTimeT16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT16 = Nothing
Dim tag_TotalTimeT17
tag_TotalTimeT17 = HMIRuntime.Tags("TotalTimeT17").Read
strSQL = "INSERT INTO z_tag_TotalTimeT17 (tag_value, created) values(" & tag_TotalTimeT17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT17 = Nothing
Dim tag_TotalTimeT18
tag_TotalTimeT18 = HMIRuntime.Tags("TotalTimeT18").Read
strSQL = "INSERT INTO z_tag_TotalTimeT18 (tag_value, created) values(" & tag_TotalTimeT18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT18 = Nothing
Dim tag_TotalTimeT19
tag_TotalTimeT19 = HMIRuntime.Tags("TotalTimeT19").Read
strSQL = "INSERT INTO z_tag_TotalTimeT19 (tag_value, created) values(" & tag_TotalTimeT19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT19 = Nothing
Dim tag_TotalTimeT20
tag_TotalTimeT20 = HMIRuntime.Tags("TotalTimeT20").Read
strSQL = "INSERT INTO z_tag_TotalTimeT20 (tag_value, created) values(" & tag_TotalTimeT20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT20 = Nothing
Dim tag_TotalTimeT21
tag_TotalTimeT21 = HMIRuntime.Tags("TotalTimeT21").Read
strSQL = "INSERT INTO z_tag_TotalTimeT21 (tag_value, created) values(" & tag_TotalTimeT21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT21 = Nothing
Dim tag_TotalTimeT22
tag_TotalTimeT22 = HMIRuntime.Tags("TotalTimeT22").Read
strSQL = "INSERT INTO z_tag_TotalTimeT22 (tag_value, created) values(" & tag_TotalTimeT22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT22 = Nothing
Dim tag_TotalTimeT23
tag_TotalTimeT23 = HMIRuntime.Tags("TotalTimeT23").Read
strSQL = "INSERT INTO z_tag_TotalTimeT23 (tag_value, created) values(" & tag_TotalTimeT23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT23 = Nothing
Dim tag_TotalTimeT24
tag_TotalTimeT24 = HMIRuntime.Tags("TotalTimeT24").Read
strSQL = "INSERT INTO z_tag_TotalTimeT24 (tag_value, created) values(" & tag_TotalTimeT24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT24 = Nothing
Dim tag_TotalTimeT25
tag_TotalTimeT25 = HMIRuntime.Tags("TotalTimeT25").Read
strSQL = "INSERT INTO z_tag_TotalTimeT25 (tag_value, created) values(" & tag_TotalTimeT25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT25 = Nothing
Dim tag_TotalTimeT26
tag_TotalTimeT26 = HMIRuntime.Tags("TotalTimeT26").Read
strSQL = "INSERT INTO z_tag_TotalTimeT26 (tag_value, created) values(" & tag_TotalTimeT26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT26 = Nothing
Dim tag_TotalTimeT27
tag_TotalTimeT27 = HMIRuntime.Tags("TotalTimeT27").Read
strSQL = "INSERT INTO z_tag_TotalTimeT27 (tag_value, created) values(" & tag_TotalTimeT27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT27 = Nothing
Dim tag_TotalTimeT28
tag_TotalTimeT28 = HMIRuntime.Tags("TotalTimeT28").Read
strSQL = "INSERT INTO z_tag_TotalTimeT28 (tag_value, created) values(" & tag_TotalTimeT28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT28 = Nothing
Dim tag_TotalTimeT29
tag_TotalTimeT29 = HMIRuntime.Tags("TotalTimeT29").Read
strSQL = "INSERT INTO z_tag_TotalTimeT29 (tag_value, created) values(" & tag_TotalTimeT29 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT29 = Nothing
Dim tag_TotalTimeT30
tag_TotalTimeT30 = HMIRuntime.Tags("TotalTimeT30").Read
strSQL = "INSERT INTO z_tag_TotalTimeT30 (tag_value, created) values(" & tag_TotalTimeT30 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT30 = Nothing
Dim tag_TotalTimeT31
tag_TotalTimeT31 = HMIRuntime.Tags("TotalTimeT31").Read
strSQL = "INSERT INTO z_tag_TotalTimeT31 (tag_value, created) values(" & tag_TotalTimeT31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT31 = Nothing
Dim tag_TotalTimeT32
tag_TotalTimeT32 = HMIRuntime.Tags("TotalTimeT32").Read
strSQL = "INSERT INTO z_tag_TotalTimeT32 (tag_value, created) values(" & tag_TotalTimeT32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT32 = Nothing
Dim tag_TotalTimeT33
tag_TotalTimeT33 = HMIRuntime.Tags("TotalTimeT33").Read
strSQL = "INSERT INTO z_tag_TotalTimeT33 (tag_value, created) values(" & tag_TotalTimeT33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT33 = Nothing
Dim tag_TotalTimeT34
tag_TotalTimeT34 = HMIRuntime.Tags("TotalTimeT34").Read
strSQL = "INSERT INTO z_tag_TotalTimeT34 (tag_value, created) values(" & tag_TotalTimeT34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_TotalTimeT34 = Nothing
Dim tag_Type_Beer_T01
tag_Type_Beer_T01 = HMIRuntime.Tags("Type_Beer_T01").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T01 (tag_value, created) values(" & tag_Type_Beer_T01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T01 = Nothing
Dim tag_Type_Beer_T02
tag_Type_Beer_T02 = HMIRuntime.Tags("Type_Beer_T02").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T02 (tag_value, created) values(" & tag_Type_Beer_T02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T02 = Nothing
Dim tag_Type_Beer_T03
tag_Type_Beer_T03 = HMIRuntime.Tags("Type_Beer_T03").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T03 (tag_value, created) values(" & tag_Type_Beer_T03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T03 = Nothing
Dim tag_Type_Beer_T04
tag_Type_Beer_T04 = HMIRuntime.Tags("Type_Beer_T04").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T04 (tag_value, created) values(" & tag_Type_Beer_T04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T04 = Nothing
Dim tag_Type_Beer_T05
tag_Type_Beer_T05 = HMIRuntime.Tags("Type_Beer_T05").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T05 (tag_value, created) values(" & tag_Type_Beer_T05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T05 = Nothing
Dim tag_Type_Beer_T06
tag_Type_Beer_T06 = HMIRuntime.Tags("Type_Beer_T06").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T06 (tag_value, created) values(" & tag_Type_Beer_T06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T06 = Nothing
Dim tag_Type_Beer_T07
tag_Type_Beer_T07 = HMIRuntime.Tags("Type_Beer_T07").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T07 (tag_value, created) values(" & tag_Type_Beer_T07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T07 = Nothing
Dim tag_Type_Beer_T08
tag_Type_Beer_T08 = HMIRuntime.Tags("Type_Beer_T08").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T08 (tag_value, created) values(" & tag_Type_Beer_T08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T08 = Nothing
Dim tag_Type_Beer_T09
tag_Type_Beer_T09 = HMIRuntime.Tags("Type_Beer_T09").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T09 (tag_value, created) values(" & tag_Type_Beer_T09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T09 = Nothing
Dim tag_Type_Beer_T10
tag_Type_Beer_T10 = HMIRuntime.Tags("Type_Beer_T10").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T10 (tag_value, created) values(" & tag_Type_Beer_T10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T10 = Nothing
Dim tag_Type_Beer_T11
tag_Type_Beer_T11 = HMIRuntime.Tags("Type_Beer_T11").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T11 (tag_value, created) values(" & tag_Type_Beer_T11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T11 = Nothing
Dim tag_Type_Beer_T12
tag_Type_Beer_T12 = HMIRuntime.Tags("Type_Beer_T12").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T12 (tag_value, created) values(" & tag_Type_Beer_T12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T12 = Nothing
Dim tag_Type_Beer_T13
tag_Type_Beer_T13 = HMIRuntime.Tags("Type_Beer_T13").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T13 (tag_value, created) values(" & tag_Type_Beer_T13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T13 = Nothing
Dim tag_Type_Beer_T14
tag_Type_Beer_T14 = HMIRuntime.Tags("Type_Beer_T14").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T14 (tag_value, created) values(" & tag_Type_Beer_T14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T14 = Nothing
Dim tag_Type_Beer_T15
tag_Type_Beer_T15 = HMIRuntime.Tags("Type_Beer_T15").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T15 (tag_value, created) values(" & tag_Type_Beer_T15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T15 = Nothing
Dim tag_Type_Beer_T16
tag_Type_Beer_T16 = HMIRuntime.Tags("Type_Beer_T16").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T16 (tag_value, created) values(" & tag_Type_Beer_T16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T16 = Nothing
Dim tag_Type_Beer_T17
tag_Type_Beer_T17 = HMIRuntime.Tags("Type_Beer_T17").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T17 (tag_value, created) values(" & tag_Type_Beer_T17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T17 = Nothing
Dim tag_Type_Beer_T18
tag_Type_Beer_T18 = HMIRuntime.Tags("Type_Beer_T18").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T18 (tag_value, created) values(" & tag_Type_Beer_T18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T18 = Nothing
Dim tag_Type_Beer_T19
tag_Type_Beer_T19 = HMIRuntime.Tags("Type_Beer_T19").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T19 (tag_value, created) values(" & tag_Type_Beer_T19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T19 = Nothing
Dim tag_Type_Beer_T20
tag_Type_Beer_T20 = HMIRuntime.Tags("Type_Beer_T20").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T20 (tag_value, created) values(" & tag_Type_Beer_T20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T20 = Nothing
Dim tag_Type_Beer_T21
tag_Type_Beer_T21 = HMIRuntime.Tags("Type_Beer_T21").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T21 (tag_value, created) values(" & tag_Type_Beer_T21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T21 = Nothing
Dim tag_Type_Beer_T22
tag_Type_Beer_T22 = HMIRuntime.Tags("Type_Beer_T22").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T22 (tag_value, created) values(" & tag_Type_Beer_T22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T22 = Nothing
Dim tag_Type_Beer_T23
tag_Type_Beer_T23 = HMIRuntime.Tags("Type_Beer_T23").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T23 (tag_value, created) values(" & tag_Type_Beer_T23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T23 = Nothing
Dim tag_Type_Beer_T24
tag_Type_Beer_T24 = HMIRuntime.Tags("Type_Beer_T24").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T24 (tag_value, created) values(" & tag_Type_Beer_T24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T24 = Nothing
Dim tag_Type_Beer_T25
tag_Type_Beer_T25 = HMIRuntime.Tags("Type_Beer_T25").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T25 (tag_value, created) values(" & tag_Type_Beer_T25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T25 = Nothing
Dim tag_Type_Beer_T26
tag_Type_Beer_T26 = HMIRuntime.Tags("Type_Beer_T26").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T26 (tag_value, created) values(" & tag_Type_Beer_T26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T26 = Nothing
Dim tag_Type_Beer_T27
tag_Type_Beer_T27 = HMIRuntime.Tags("Type_Beer_T27").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T27 (tag_value, created) values(" & tag_Type_Beer_T27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T27 = Nothing
Dim tag_Type_Beer_T28
tag_Type_Beer_T28 = HMIRuntime.Tags("Type_Beer_T28").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T28 (tag_value, created) values(" & tag_Type_Beer_T28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T28 = Nothing
Dim tag_Type_Beer_T31
tag_Type_Beer_T31 = HMIRuntime.Tags("Type_Beer_T31").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T31 (tag_value, created) values(" & tag_Type_Beer_T31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T31 = Nothing
Dim tag_Type_Beer_T32
tag_Type_Beer_T32 = HMIRuntime.Tags("Type_Beer_T32").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T32 (tag_value, created) values(" & tag_Type_Beer_T32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T32 = Nothing
Dim tag_Type_Beer_T33
tag_Type_Beer_T33 = HMIRuntime.Tags("Type_Beer_T33").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T33 (tag_value, created) values(" & tag_Type_Beer_T33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T33 = Nothing
Dim tag_Type_Beer_T34
tag_Type_Beer_T34 = HMIRuntime.Tags("Type_Beer_T34").Read
strSQL = "INSERT INTO z_tag_Type_Beer_T34 (tag_value, created) values(" & tag_Type_Beer_T34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_T34 = Nothing
Dim tag_Type_Beer_Tank31
tag_Type_Beer_Tank31 = HMIRuntime.Tags("Type_Beer_Tank31").Read
strSQL = "INSERT INTO z_tag_Type_Beer_Tank31 (tag_value, created) values(" & tag_Type_Beer_Tank31 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_Tank31 = Nothing
Dim tag_Type_Beer_Tank32
tag_Type_Beer_Tank32 = HMIRuntime.Tags("Type_Beer_Tank32").Read
strSQL = "INSERT INTO z_tag_Type_Beer_Tank32 (tag_value, created) values(" & tag_Type_Beer_Tank32 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_Tank32 = Nothing
Dim tag_Type_Beer_Tank33
tag_Type_Beer_Tank33 = HMIRuntime.Tags("Type_Beer_Tank33").Read
strSQL = "INSERT INTO z_tag_Type_Beer_Tank33 (tag_value, created) values(" & tag_Type_Beer_Tank33 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_Tank33 = Nothing
Dim tag_Type_Beer_Tank34
tag_Type_Beer_Tank34 = HMIRuntime.Tags("Type_Beer_Tank34").Read
strSQL = "INSERT INTO z_tag_Type_Beer_Tank34 (tag_value, created) values(" & tag_Type_Beer_Tank34 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_Tank34 = Nothing
Dim tag_Type_Beer_TBF01
tag_Type_Beer_TBF01 = HMIRuntime.Tags("Type_Beer_TBF01").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF01 (tag_value, created) values(" & tag_Type_Beer_TBF01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF01 = Nothing
Dim tag_Type_Beer_TBF02
tag_Type_Beer_TBF02 = HMIRuntime.Tags("Type_Beer_TBF02").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF02 (tag_value, created) values(" & tag_Type_Beer_TBF02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF02 = Nothing
Dim tag_Type_Beer_TBF03
tag_Type_Beer_TBF03 = HMIRuntime.Tags("Type_Beer_TBF03").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF03 (tag_value, created) values(" & tag_Type_Beer_TBF03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF03 = Nothing
Dim tag_Type_Beer_TBF04
tag_Type_Beer_TBF04 = HMIRuntime.Tags("Type_Beer_TBF04").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF04 (tag_value, created) values(" & tag_Type_Beer_TBF04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF04 = Nothing
Dim tag_Type_Beer_TBF07
tag_Type_Beer_TBF07 = HMIRuntime.Tags("Type_Beer_TBF07").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF07 (tag_value, created) values(" & tag_Type_Beer_TBF07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF07 = Nothing
Dim tag_Type_Beer_TBF08
tag_Type_Beer_TBF08 = HMIRuntime.Tags("Type_Beer_TBF08").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF08 (tag_value, created) values(" & tag_Type_Beer_TBF08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF08 = Nothing
Dim tag_Type_Beer_TBF09
tag_Type_Beer_TBF09 = HMIRuntime.Tags("Type_Beer_TBF09").Read
strSQL = "INSERT INTO z_tag_Type_Beer_TBF09 (tag_value, created) values(" & tag_Type_Beer_TBF09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Type_Beer_TBF09 = Nothing
Dim tag_ValvedayBBT01
tag_ValvedayBBT01 = HMIRuntime.Tags("ValvedayBBT01").Read
strSQL = "INSERT INTO z_tag_ValvedayBBT01 (tag_value, created) values(" & tag_ValvedayBBT01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayBBT01 = Nothing
Dim tag_ValvedayBBT02
tag_ValvedayBBT02 = HMIRuntime.Tags("ValvedayBBT02").Read
strSQL = "INSERT INTO z_tag_ValvedayBBT02 (tag_value, created) values(" & tag_ValvedayBBT02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayBBT02 = Nothing
Dim tag_ValvedayBBT03
tag_ValvedayBBT03 = HMIRuntime.Tags("ValvedayBBT03").Read
strSQL = "INSERT INTO z_tag_ValvedayBBT03 (tag_value, created) values(" & tag_ValvedayBBT03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayBBT03 = Nothing
Dim tag_ValvedayBBT04
tag_ValvedayBBT04 = HMIRuntime.Tags("ValvedayBBT04").Read
strSQL = "INSERT INTO z_tag_ValvedayBBT04 (tag_value, created) values(" & tag_ValvedayBBT04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayBBT04 = Nothing
Dim tag_ValvedayT01
tag_ValvedayT01 = HMIRuntime.Tags("ValvedayT01").Read
strSQL = "INSERT INTO z_tag_ValvedayT01 (tag_value, created) values(" & tag_ValvedayT01 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT01 = Nothing
Dim tag_ValvedayT02
tag_ValvedayT02 = HMIRuntime.Tags("ValvedayT02").Read
strSQL = "INSERT INTO z_tag_ValvedayT02 (tag_value, created) values(" & tag_ValvedayT02 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT02 = Nothing
Dim tag_ValvedayT03
tag_ValvedayT03 = HMIRuntime.Tags("ValvedayT03").Read
strSQL = "INSERT INTO z_tag_ValvedayT03 (tag_value, created) values(" & tag_ValvedayT03 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT03 = Nothing
Dim tag_ValvedayT04
tag_ValvedayT04 = HMIRuntime.Tags("ValvedayT04").Read
strSQL = "INSERT INTO z_tag_ValvedayT04 (tag_value, created) values(" & tag_ValvedayT04 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT04 = Nothing
Dim tag_ValvedayT05
tag_ValvedayT05 = HMIRuntime.Tags("ValvedayT05").Read
strSQL = "INSERT INTO z_tag_ValvedayT05 (tag_value, created) values(" & tag_ValvedayT05 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT05 = Nothing
Dim tag_ValvedayT06
tag_ValvedayT06 = HMIRuntime.Tags("ValvedayT06").Read
strSQL = "INSERT INTO z_tag_ValvedayT06 (tag_value, created) values(" & tag_ValvedayT06 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT06 = Nothing
Dim tag_ValvedayT07
tag_ValvedayT07 = HMIRuntime.Tags("ValvedayT07").Read
strSQL = "INSERT INTO z_tag_ValvedayT07 (tag_value, created) values(" & tag_ValvedayT07 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT07 = Nothing
Dim tag_ValvedayT08
tag_ValvedayT08 = HMIRuntime.Tags("ValvedayT08").Read
strSQL = "INSERT INTO z_tag_ValvedayT08 (tag_value, created) values(" & tag_ValvedayT08 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT08 = Nothing
Dim tag_ValvedayT09
tag_ValvedayT09 = HMIRuntime.Tags("ValvedayT09").Read
strSQL = "INSERT INTO z_tag_ValvedayT09 (tag_value, created) values(" & tag_ValvedayT09 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT09 = Nothing
Dim tag_ValvedayT10
tag_ValvedayT10 = HMIRuntime.Tags("ValvedayT10").Read
strSQL = "INSERT INTO z_tag_ValvedayT10 (tag_value, created) values(" & tag_ValvedayT10 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT10 = Nothing
Dim tag_ValvedayT11
tag_ValvedayT11 = HMIRuntime.Tags("ValvedayT11").Read
strSQL = "INSERT INTO z_tag_ValvedayT11 (tag_value, created) values(" & tag_ValvedayT11 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT11 = Nothing
Dim tag_ValvedayT12
tag_ValvedayT12 = HMIRuntime.Tags("ValvedayT12").Read
strSQL = "INSERT INTO z_tag_ValvedayT12 (tag_value, created) values(" & tag_ValvedayT12 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT12 = Nothing
Dim tag_ValvedayT13
tag_ValvedayT13 = HMIRuntime.Tags("ValvedayT13").Read
strSQL = "INSERT INTO z_tag_ValvedayT13 (tag_value, created) values(" & tag_ValvedayT13 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT13 = Nothing
Dim tag_ValvedayT14
tag_ValvedayT14 = HMIRuntime.Tags("ValvedayT14").Read
strSQL = "INSERT INTO z_tag_ValvedayT14 (tag_value, created) values(" & tag_ValvedayT14 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT14 = Nothing
Dim tag_ValvedayT15
tag_ValvedayT15 = HMIRuntime.Tags("ValvedayT15").Read
strSQL = "INSERT INTO z_tag_ValvedayT15 (tag_value, created) values(" & tag_ValvedayT15 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT15 = Nothing
Dim tag_ValvedayT16
tag_ValvedayT16 = HMIRuntime.Tags("ValvedayT16").Read
strSQL = "INSERT INTO z_tag_ValvedayT16 (tag_value, created) values(" & tag_ValvedayT16 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT16 = Nothing
Dim tag_ValvedayT17
tag_ValvedayT17 = HMIRuntime.Tags("ValvedayT17").Read
strSQL = "INSERT INTO z_tag_ValvedayT17 (tag_value, created) values(" & tag_ValvedayT17 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT17 = Nothing
Dim tag_ValvedayT18
tag_ValvedayT18 = HMIRuntime.Tags("ValvedayT18").Read
strSQL = "INSERT INTO z_tag_ValvedayT18 (tag_value, created) values(" & tag_ValvedayT18 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT18 = Nothing
Dim tag_ValvedayT19
tag_ValvedayT19 = HMIRuntime.Tags("ValvedayT19").Read
strSQL = "INSERT INTO z_tag_ValvedayT19 (tag_value, created) values(" & tag_ValvedayT19 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT19 = Nothing
Dim tag_ValvedayT20
tag_ValvedayT20 = HMIRuntime.Tags("ValvedayT20").Read
strSQL = "INSERT INTO z_tag_ValvedayT20 (tag_value, created) values(" & tag_ValvedayT20 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT20 = Nothing
Dim tag_ValvedayT21
tag_ValvedayT21 = HMIRuntime.Tags("ValvedayT21").Read
strSQL = "INSERT INTO z_tag_ValvedayT21 (tag_value, created) values(" & tag_ValvedayT21 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT21 = Nothing
Dim tag_ValvedayT22
tag_ValvedayT22 = HMIRuntime.Tags("ValvedayT22").Read
strSQL = "INSERT INTO z_tag_ValvedayT22 (tag_value, created) values(" & tag_ValvedayT22 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT22 = Nothing
Dim tag_ValvedayT23
tag_ValvedayT23 = HMIRuntime.Tags("ValvedayT23").Read
strSQL = "INSERT INTO z_tag_ValvedayT23 (tag_value, created) values(" & tag_ValvedayT23 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT23 = Nothing
Dim tag_ValvedayT24
tag_ValvedayT24 = HMIRuntime.Tags("ValvedayT24").Read
strSQL = "INSERT INTO z_tag_ValvedayT24 (tag_value, created) values(" & tag_ValvedayT24 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT24 = Nothing
Dim tag_ValvedayT25
tag_ValvedayT25 = HMIRuntime.Tags("ValvedayT25").Read
strSQL = "INSERT INTO z_tag_ValvedayT25 (tag_value, created) values(" & tag_ValvedayT25 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT25 = Nothing
Dim tag_ValvedayT26
tag_ValvedayT26 = HMIRuntime.Tags("ValvedayT26").Read
strSQL = "INSERT INTO z_tag_ValvedayT26 (tag_value, created) values(" & tag_ValvedayT26 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT26 = Nothing
Dim tag_ValvedayT27
tag_ValvedayT27 = HMIRuntime.Tags("ValvedayT27").Read
strSQL = "INSERT INTO z_tag_ValvedayT27 (tag_value, created) values(" & tag_ValvedayT27 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT27 = Nothing
Dim tag_ValvedayT28
tag_ValvedayT28 = HMIRuntime.Tags("ValvedayT28").Read
strSQL = "INSERT INTO z_tag_ValvedayT28 (tag_value, created) values(" & tag_ValvedayT28 & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_ValvedayT28 = Nothing
Dim tag_Valveglylol_ncDA
tag_Valveglylol_ncDA = HMIRuntime.Tags("Valveglylol_ncDA").Read
strSQL = "INSERT INTO z_tag_Valveglylol_ncDA (tag_value, created) values(" & tag_Valveglylol_ncDA & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_Valveglylol_ncDA = Nothing
Dim tag_VancapglycolnuocDA
tag_VancapglycolnuocDA = HMIRuntime.Tags("VancapglycolnuocDA").Read
strSQL = "INSERT INTO z_tag_VancapglycolnuocDA (tag_value, created) values(" & tag_VancapglycolnuocDA & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_VancapglycolnuocDA = Nothing
Dim tag_VANDAY_TACHMEN70hl
tag_VANDAY_TACHMEN70hl = HMIRuntime.Tags("VANDAY_TACHMEN70hl").Read
strSQL = "INSERT INTO z_tag_VANDAY_TACHMEN70hl (tag_value, created) values(" & tag_VANDAY_TACHMEN70hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_VANDAY_TACHMEN70hl = Nothing
Dim tag_VANDAY_TACHMEN100hl
tag_VANDAY_TACHMEN100hl = HMIRuntime.Tags("VANDAY_TACHMEN100hl").Read
strSQL = "INSERT INTO z_tag_VANDAY_TACHMEN100hl (tag_value, created) values(" & tag_VANDAY_TACHMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_VANDAY_TACHMEN100hl = Nothing
Dim tag_VANDAY_TANKMEN100hl
tag_VANDAY_TANKMEN100hl = HMIRuntime.Tags("VANDAY_TANKMEN100hl").Read
strSQL = "INSERT INTO z_tag_VANDAY_TANKMEN100hl (tag_value, created) values(" & tag_VANDAY_TANKMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_VANDAY_TANKMEN100hl = Nothing
Dim tag_WET_TACHMEN70hl
tag_WET_TACHMEN70hl = HMIRuntime.Tags("WET_TACHMEN70hl").Read
strSQL = "INSERT INTO z_tag_WET_TACHMEN70hl (tag_value, created) values(" & tag_WET_TACHMEN70hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WET_TACHMEN70hl = Nothing
Dim tag_WET_TACHMEN100hl
tag_WET_TACHMEN100hl = HMIRuntime.Tags("WET_TACHMEN100hl").Read
strSQL = "INSERT INTO z_tag_WET_TACHMEN100hl (tag_value, created) values(" & tag_WET_TACHMEN100hl & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_WET_TACHMEN100hl = Nothing
Dim tag_YeastTank01_auto
tag_YeastTank01_auto = HMIRuntime.Tags("YeastTank01_auto").Read
strSQL = "INSERT INTO z_tag_YeastTank01_auto (tag_value, created) values(" & tag_YeastTank01_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank01_auto = Nothing
Dim tag_YeastTank01_CIP
tag_YeastTank01_CIP = HMIRuntime.Tags("YeastTank01_CIP").Read
strSQL = "INSERT INTO z_tag_YeastTank01_CIP (tag_value, created) values(" & tag_YeastTank01_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank01_CIP = Nothing
Dim tag_YeastTank01_run
tag_YeastTank01_run = HMIRuntime.Tags("YeastTank01_run").Read
strSQL = "INSERT INTO z_tag_YeastTank01_run (tag_value, created) values(" & tag_YeastTank01_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank01_run = Nothing
Dim tag_YeastTank01_seq
tag_YeastTank01_seq = HMIRuntime.Tags("YeastTank01_seq").Read
strSQL = "INSERT INTO z_tag_YeastTank01_seq (tag_value, created) values(" & tag_YeastTank01_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank01_seq = Nothing
Dim tag_YeastTank02_auto
tag_YeastTank02_auto = HMIRuntime.Tags("YeastTank02_auto").Read
strSQL = "INSERT INTO z_tag_YeastTank02_auto (tag_value, created) values(" & tag_YeastTank02_auto & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank02_auto = Nothing
Dim tag_YeastTank02_CIP
tag_YeastTank02_CIP = HMIRuntime.Tags("YeastTank02_CIP").Read
strSQL = "INSERT INTO z_tag_YeastTank02_CIP (tag_value, created) values(" & tag_YeastTank02_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank02_CIP = Nothing
Dim tag_YeastTank02_run
tag_YeastTank02_run = HMIRuntime.Tags("YeastTank02_run").Read
strSQL = "INSERT INTO z_tag_YeastTank02_run (tag_value, created) values(" & tag_YeastTank02_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank02_run = Nothing
Dim tag_YeastTank02_seq
tag_YeastTank02_seq = HMIRuntime.Tags("YeastTank02_seq").Read
strSQL = "INSERT INTO z_tag_YeastTank02_seq (tag_value, created) values(" & tag_YeastTank02_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank02_seq = Nothing
Dim tag_YeastTank03_CIP
tag_YeastTank03_CIP = HMIRuntime.Tags("YeastTank03_CIP").Read
strSQL = "INSERT INTO z_tag_YeastTank03_CIP (tag_value, created) values(" & tag_YeastTank03_CIP & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank03_CIP = Nothing
Dim tag_YeastTank03_run
tag_YeastTank03_run = HMIRuntime.Tags("YeastTank03_run").Read
strSQL = "INSERT INTO z_tag_YeastTank03_run (tag_value, created) values(" & tag_YeastTank03_run & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank03_run = Nothing
Dim tag_YeastTank03_seq
tag_YeastTank03_seq = HMIRuntime.Tags("YeastTank03_seq").Read
strSQL = "INSERT INTO z_tag_YeastTank03_seq (tag_value, created) values(" & tag_YeastTank03_seq & ", CURRENT_TIMESTAMP);"
Set objCommand = CreateObject("ADODB.Command")
With objCommand
    .ActiveConnection = objConnection
    .CommandText = strSQL
End With
objCommand.Execute
Set objCommand = Nothing
Set tag_YeastTank03_seq = Nothing

objConnection.Close
Set objConnection = Nothing
End Function