Attribute VB_Name = "basUpgradeDB3"
Option Explicit


Public Sub Go_to_UpgradeDB3_module()

On Error GoTo ErrorTrap

    If gstrDBVersion <= "0005.9974.0000" Then
        DB_Upgrade_5_9974_00_To_5_9977_00
    End If

    If gstrDBVersion <= "0005.9977.0000" Then
        DB_Upgrade_5_9977_00_To_5_9977_02
    End If

    If gstrDBVersion <= "0005.9977.0002" Then
        DB_Upgrade_5_9977_02_To_5_9977_03
    End If

    If gstrDBVersion <= "0005.9977.0003" Then
        DB_Upgrade_5_9977_03_To_5_9978_00
    End If

    If gstrDBVersion <= "0005.9978.0000" Then
        DB_Upgrade_5_9978_00_To_5_99778_01
    End If

    If gstrDBVersion <= "0005.9978.0001" Then
        DB_Upgrade_5_9978_01_To_5_99779_00
    End If
    
    If gstrDBVersion <= "0005.9979.0000" Then
        DB_Upgrade_5_9979_00_To_5_99779_01
    End If

    If gstrDBVersion <= "0005.9979.0001" Then
        DB_Upgrade_5_9979_01_To_5_99779_02
    End If

    If gstrDBVersion <= "0005.9979.0002" Then
        DB_Upgrade_5_9979_02_To_5_9979_08
    End If

    If gstrDBVersion <= "0005.9979.0008" Then
        DB_Upgrade_5_9979_08_To_5_9979_09
    End If
    
    If gstrDBVersion <= "0005.9979.0009" Then
        DB_Upgrade_5_9979_09_To_5_9979_10
    End If
 
    If gstrDBVersion <= "0005.9979.0010" Then
        DB_Upgrade_0005_9979_0010_To_0005_9979_0011
    End If

    If gstrDBVersion <= "0005.9979.0011" Then
        DB_Upgrade_0005_9979_0011_To_0005_9979_0012
    End If
    
    If gstrDBVersion <= "0005.9979.0012" Then
        DB_Upgrade_0005_9979_0012_To_0005_9979_0013
    End If
    
    If gstrDBVersion <= "0005.9979.0013" Then
        DB_Upgrade_0005_9979_0013_To_0005_9979_0014
    End If
    
    If gstrDBVersion <= "0005.9979.0014" Then
        DB_Upgrade_0005_9979_0014_To_0005_9979_0015
    End If
    
    If gstrDBVersion <= "0005.9979.0015" Then
        DB_Upgrade_0005_9979_0015_To_0005_9979_0016
    End If
    
    If gstrDBVersion <= "0005.9979.0016" Then
        DB_Upgrade_0005_9979_0016_To_0005_9979_0017
    End If
    
    If gstrDBVersion <= "0005.9979.0017" Then
        DB_Upgrade_0005_9979_0017_To_0005_9979_0018
    End If
    
    If gstrDBVersion <= "0005.9979.0018" Then
        DB_Upgrade_0005_9979_0018_To_0005_9979_0019
    End If
    
    If gstrDBVersion <= "0005.9979.0019" Then
        DB_Upgrade_0005_9979_0019_To_0005_9979_0020
    End If
    
    If gstrDBVersion <= "0005.9979.0020" Then
        DB_Upgrade_0005_9979_0020_To_0005_9979_0021
    End If
    
    If gstrDBVersion <= "0005.9979.0021" Then
        DB_Upgrade_0005_9979_0021_To_0005_9979_0022
    End If
    
    If gstrDBVersion <= "0005.9979.0022" Then
        DB_Upgrade_0005_9979_0022_To_0005_9979_0023
    End If
    
    If gstrDBVersion <= "0005.9979.0023" Then
        DB_Upgrade_0005_9979_0023_To_0005_9979_0024
    End If
    
    If gstrDBVersion <= "0005.9979.0024" Then
        DB_Upgrade_0005_9979_0024_To_0005_9979_0025
    End If
    
    If gstrDBVersion <= "0005.9979.0025" Then
        DB_Upgrade_0005_9979_0024_To_0005_9979_0026
    End If
    
    If gstrDBVersion <= "0005.9979.0026" Then
        DB_Upgrade_0005_9979_0026_To_0005_9979_0027
    End If
    
    If gstrDBVersion <= "0005.9979.0027" Then
        DB_Upgrade_0005_9979_0027_To_0005_9979_0028
    End If
    
    If gstrDBVersion <= "0005.9979.0028" Then
        DB_Upgrade_0005_9979_0028_To_0005_9979_0029
    End If
    
    If gstrDBVersion <= "0005.9979.0029" Then
        DB_Upgrade_0005_9979_0029_To_0005_9979_0030
    End If
    
    If gstrDBVersion <= "0005.9979.0030" Then
        DB_Upgrade_0005_9979_0030_To_0005_9979_0031
    End If
        
    If gstrDBVersion <= "0005.9979.0031" Then
        DB_Upgrade_0005_9979_0031_To_0005_9979_0032
    End If
        
    If gstrDBVersion <= "0005.9979.0032" Then
        DB_Upgrade_0005_9979_0032_To_0005_9979_0033
    End If
    
    If gstrDBVersion <= "0005.9979.0033" Then
        DB_Upgrade_0005_9979_0033_To_0005_9979_0034
    End If
    
    If gstrDBVersion <= "0005.9979.0034" Then
        DB_Upgrade_0005_9979_0034_To_0005_9979_0035
    End If
    
    If gstrDBVersion <= "0005.9979.0035" Then
        DB_Upgrade_0005_9979_0035_To_0005_9979_0037
    End If
    
    If gstrDBVersion <= "0005.9979.0037" Then
        DB_Upgrade_0005_9979_0037_To_0005_9979_0038
    End If
    
    If gstrDBVersion <= "0005.9979.0038" Then
        DB_Upgrade_0005_9979_0038_To_0005_9979_0039
    End If
    
    If gstrDBVersion <= "0005.9979.0039" Then
        DB_Upgrade_0005_9979_0039_To_0005_9979_0040
    End If
    
    If gstrDBVersion <= "0005.9979.0040" Then
        DB_Upgrade_0005_9979_0040_To_0005_9979_0041
    End If
    
    If gstrDBVersion <= "0005.9979.0041" Then
        DB_Upgrade_0005_9979_0041_To_0005_9979_0042
    End If
    
    If gstrDBVersion <= "0005.9979.0042" Then
        DB_Upgrade_0005_9979_0042_To_0005_9979_0043
    End If
    
    If gstrDBVersion <= "0005.9979.0043" Then
        DB_Upgrade_0005_9979_0043_To_0005_9979_0044
    End If
    
    If gstrDBVersion <= "0005.9979.0044" Then
        DB_Upgrade_0005_9979_0044_To_0005_9979_0045
    End If
    
    If gstrDBVersion <= "0005.9979.0045" Then
        DB_Upgrade_0005_9979_0045_To_0005_9979_0046
    End If
    
    If gstrDBVersion <= "0005.9979.0046" Then
        DB_Upgrade_0005_9979_0046_To_0005_9979_0047
    End If
    
    If gstrDBVersion <= "0005.9979.0047" Then
        DB_Upgrade_0005_9979_0047_To_0005_9979_0048
    End If
    
    If gstrDBVersion <= "0005.9979.0048" Then
        DB_Upgrade_0005_9979_0048_To_0005_9979_0049
    End If
    
    If gstrDBVersion <= "0005.9979.0049" Then
        DB_Upgrade_0005_9979_0049_To_0005_9979_0050
    End If
    
    If gstrDBVersion <= "0005.9979.0050" Then
        DB_Upgrade_0005_9979_0050_To_0005_9979_0051
    End If
    
    If gstrDBVersion <= "0005.9979.0051" Then
        DB_Upgrade_0005_9979_0051_To_0005_9979_0052
    End If
    
    If gstrDBVersion <= "0005.9979.0052" Then
        DB_Upgrade_0005_9979_0052_To_0005_9979_0053
    End If
    
    If gstrDBVersion <= "0005.9979.0053" Then
        DB_Upgrade_0005_9979_0053_To_0005_9979_0054
    End If
    
    If gstrDBVersion <= "0005.9979.0054" Then
        DB_Upgrade_0005_9979_0054_To_0005_9979_0055
    End If
    
    If gstrDBVersion <= "0005.9979.0055" Then
        DB_Upgrade_0005_9979_0055_To_0005_9979_0056
    End If
    
    If gstrDBVersion <= "0005.9979.0056" Then
        DB_Upgrade_0005_9979_0056_To_0005_9979_0057
    End If
    
    If gstrDBVersion <= "0005.9979.0057" Then
        DB_Upgrade_0005_9979_0057_To_0005_9979_0058
    End If
    
    If gstrDBVersion <= "0005.9979.0058" Then
        DB_Upgrade_0005_9979_0058_To_0005_9979_0059
    End If
    
    If gstrDBVersion <= "0005.9979.0059" Then
        DB_Upgrade_0005_9979_0059_To_0005_9979_0060
    End If
    
    If gstrDBVersion <= "0005.9979.0060" Then
        DB_Upgrade_0005_9979_0060_To_0005_9979_0061
    End If
    
    If gstrDBVersion <= "0005.9979.0061" Then
        DB_Upgrade_0005_9979_0061_To_0005_9979_0062
    End If
    
    If gstrDBVersion <= "0005.9979.0062" Then
        DB_Upgrade_0005_9979_0062_To_0005_9979_0063
    End If
    
    If gstrDBVersion <= "0005.9979.0063" Then
        DB_Upgrade_0005_9979_0063_To_0005_9979_0064
    End If
    
    If gstrDBVersion <= "0005.9979.0064" Then
        DB_Upgrade_0005_9979_0064_To_0005_9979_0065
    End If
    
    If gstrDBVersion <= "0005.9979.0065" Then
        DB_Upgrade_0005_9979_0065_To_0005_9979_0066
    End If
    
    If gstrDBVersion <= "0005.9979.0066" Then
        DB_Upgrade_0005_9979_0066_To_0005_9979_0067
    End If
    
    If gstrDBVersion <= "0005.9979.0067" Then
        DB_Upgrade_0005_9979_0067_To_0005_9979_0068
    End If
    
    If gstrDBVersion <= "0005.9979.0068" Then
        DB_Upgrade_0005_9979_0068_To_0005_9979_0069
    End If
    
    If gstrDBVersion <= "0005.9979.0069" Then
        DB_Upgrade_0005_9979_0069_To_0005_9979_0070
    End If
    
    If gstrDBVersion <= "0005.9979.0070" Then
        DB_Upgrade_0005_9979_0070_To_0005_9979_0071
    End If
    
    If gstrDBVersion <= "0005.9979.0071" Then
        DB_Upgrade_0005_9979_0071_To_0005_9979_0072
    End If
    
    If gstrDBVersion <= "0005.9979.0072" Then
        DB_Upgrade_0005_9979_0072_To_0005_9979_0073
    End If
    
    If gstrDBVersion <= "0005.9979.0073" Then
        DB_Upgrade_0005_9979_0073_To_0005_9979_0074
    End If
    
    If gstrDBVersion <= "0005.9979.0074" Then
        DB_Upgrade_0005_9979_0074_To_0005_9979_0075
    End If
    
    If gstrDBVersion <= "0005.9979.0075" Then
        DB_Upgrade_0005_9979_0075_To_0005_9979_0076
    End If
    
    If gstrDBVersion <= "0005.9979.0076" Then
        DB_Upgrade_0005_9979_0076_To_0005_9979_0077
    End If
    
    If gstrDBVersion <= "0005.9979.0077" Then
        DB_Upgrade_0005_9979_0077_To_0005_9979_0078
    End If
    
    If gstrDBVersion <= "0005.9979.0078" Then
        DB_Upgrade_0005_9979_0078_To_0005_9979_0079
    End If
    
    If gstrDBVersion <= "0005.9979.0079" Then
        DB_Upgrade_0005_9979_0079_To_0005_9979_0080
    End If
    
    If gstrDBVersion <= "0005.9979.0080" Then
        DB_Upgrade_0005_9979_0080_To_0005_9979_0081
    End If
    
    If gstrDBVersion <= "0005.9979.0081" Then
        DB_Upgrade_0005_9979_0081_To_0005_9979_0082
    End If
    
    If gstrDBVersion <= "0005.9979.0082" Then
        DB_Upgrade_0005_9979_0082_To_0005_9979_0083
    End If
    
    If gstrDBVersion <= "0005.9979.0083" Then
        DB_Upgrade_0005_9979_0083_To_0005_9979_0084
    End If
    
    If gstrDBVersion <= "0005.9979.0084" Then
        DB_Upgrade_0005_9979_0084_To_0005_9979_0085
    End If
    
    If gstrDBVersion <= "0005.9979.0085" Then
        DB_Upgrade_0005_9979_0085_To_0005_9979_0086
    End If
    
    If gstrDBVersion <= "0005.9979.0086" Then
        DB_Upgrade_0005_9979_0086_To_0005_9979_0087
    End If
    
    If gstrDBVersion <= "0005.9979.0087" Then
        DB_Upgrade_0005_9979_0087_To_0005_9979_0088
    End If
    
    If gstrDBVersion <= "0005.9979.0088" Then
        DB_Upgrade_0005_9979_0088_To_0005_9979_0089
    End If
    
    If gstrDBVersion <= "0005.9979.0089" Then
        DB_Upgrade_0005_9979_0089_To_0005_9979_0090
    End If
    
    If gstrDBVersion <= "0005.9979.0090" Then
        DB_Upgrade_0005_9979_0090_To_0005_9979_0091
    End If
    
    If gstrDBVersion <= "0005.9979.0091" Then
        DB_Upgrade_0005_9979_0091_To_0005_9979_0092
    End If
    
    If gstrDBVersion <= "0005.9979.0092" Then
        DB_Upgrade_0005_9979_0092_To_0005_9979_0093
    End If
    
    If gstrDBVersion <= "0005.9979.0093" Then
        DB_Upgrade_0005_9979_0093_To_0005_9979_0094
    End If
    
    If gstrDBVersion <= "0005.9979.0094" Then
        DB_Upgrade_0005_9979_0094_To_0005_9979_0095
    End If
    
    If gstrDBVersion <= "0005.9979.0095" Then
        DB_Upgrade_0005_9979_0095_To_0005_9979_0096
    End If
    
    If gstrDBVersion <= "0005.9979.0096" Then
        DB_Upgrade_0005_9979_0096_To_0005_9979_0097
    End If
    
    If gstrDBVersion <= "0005.9979.0097" Then
        DB_Upgrade_0005_9979_0097_To_0005_9979_0098
    End If
    
    If gstrDBVersion <= "0005.9979.0098" Then
        DB_Upgrade_0005_9979_0098_To_0005_9979_0099
    End If
    
    If gstrDBVersion <= "0005.9979.0099" Then
        DB_Upgrade_0005_9979_0099_To_0005_9979_0100
    End If
    
    If gstrDBVersion <= "0005.9979.0100" Then
        DB_Upgrade_0005_9979_0100_To_0005_9979_0101
    End If
    
    If gstrDBVersion <= "0005.9979.0101" Then
        DB_Upgrade_0005_9979_0101_To_0005_9979_0102
    End If
    
    If gstrDBVersion <= "0005.9979.0102" Then
        DB_Upgrade_0005_9979_0102_To_0005_9979_0103
    End If
    
    If gstrDBVersion <= "0005.9979.0103" Then
        DB_Upgrade_0005_9979_0103_To_0005_9979_0104
    End If
    
    If gstrDBVersion <= "0005.9979.0104" Then
        DB_Upgrade_0005_9979_0104_To_0005_9979_0105
    End If
    
    If gstrDBVersion <= "0005.9979.0105" Then
        DB_Upgrade_0005_9979_0105_To_0005_9979_0106
    End If
    
    If gstrDBVersion <= "0005.9979.0106" Then
        DB_Upgrade_0005_9979_0106_To_0005_9979_0107
    End If
    
    If gstrDBVersion <= "0005.9979.0107" Then
        DB_Upgrade_0005_9979_0107_To_0005_9979_0108
    End If
    
    If gstrDBVersion <= "0005.9979.0108" Then
        DB_Upgrade_0005_9979_0108_To_0005_9979_0109
    End If
    
    If gstrDBVersion <= "0005.9979.0109" Then
        DB_Upgrade_0005_9979_0109_To_0005_9979_0110
    End If
    
    If gstrDBVersion <= "0005.9979.0110" Then
        DB_Upgrade_0005_9979_0110_To_0005_9979_0111
    End If
    
    If gstrDBVersion <= "0005.9979.0111" Then
        DB_Upgrade_0005_9979_0111_To_0005_9979_0112
    End If
    
    If gstrDBVersion <= "0005.9979.0112" Then
        DB_Upgrade_0005_9979_0112_To_0005_9979_0113
    End If
    
    If gstrDBVersion <= "0005.9979.0113" Then
        DB_Upgrade_0005_9979_0113_To_0005_9979_0114
    End If
    
    If gstrDBVersion <= "0005.9979.0115" Then
        DB_Upgrade_0005_9979_0114_To_0005_9979_0116
    End If
    
    If gstrDBVersion <= "0005.9979.0116" Then
        DB_Upgrade_0005_9979_0116_To_0005_9979_0117
    End If
    
    If gstrDBVersion <= "0005.9979.0117" Then
        DB_Upgrade_0005_9979_0117_To_0005_9979_0118
    End If
    
    If gstrDBVersion <= "0005.9979.0118" Then
        DB_Upgrade_0005_9979_0118_To_0005_9979_0119
    End If
    
    If gstrDBVersion <= "0005.9979.0119" Then
        DB_Upgrade_0005_9979_0119_To_0005_9979_0120
    End If
    
    If gstrDBVersion <= "0005.9979.0120" Then
        DB_Upgrade_0005_9979_0120_To_0005_9979_0121
    End If
    
    If gstrDBVersion <= "0005.9979.0121" Then
        DB_Upgrade_0005_9979_0121_To_0005_9979_0122
    End If
    
    If gstrDBVersion <= "0005.9979.0122" Then
        DB_Upgrade_0005_9979_0122_To_0005_9979_0123
    End If
    
    If gstrDBVersion <= "0005.9979.0124" Then
        DB_Upgrade_0005_9979_0124_To_0005_9979_0125
    End If
    
    If gstrDBVersion <= "0005.9979.0125" Then
        DB_Upgrade_0005_9979_0125_To_0005_9979_0126
    End If
    
    If gstrDBVersion <= "0005.9979.0126" Then
        DB_Upgrade_0005_9979_0126_To_0005_9979_0127
    End If
    
    If gstrDBVersion <= "0005.9979.0127" Then
        DB_Upgrade_0005_9979_0127_To_0005_9979_0128
    End If
    
    If gstrDBVersion <= "0005.9979.0128" Then
        DB_Upgrade_0005_9979_0128_To_0005_9979_0129
    End If
    
    If gstrDBVersion <= "0005.9979.0129" Then
        DB_Upgrade_0005_9979_0129_To_0005_9979_0130
    End If
    
    If gstrDBVersion <= "0005.9979.0130" Then
        DB_Upgrade_0005_9979_0130_To_0005_9979_0131
    End If
    
    If gstrDBVersion <= "0005.9979.0131" Then
        DB_Upgrade_0005_9979_0131_To_0005_9979_0132
    End If
    
    If gstrDBVersion <= "0005.9979.0132" Then
        DB_Upgrade_0005_9979_0132_To_0005_9979_0133
    End If
    
    If gstrDBVersion <= "0005.9979.0133" Then
        DB_Upgrade_0005_9979_0133_To_0005_9979_0134
    End If
    
    If gstrDBVersion <= "0005.9979.0134" Then
        DB_Upgrade_0005_9979_0134_To_0005_9979_0135
    End If
    
    If gstrDBVersion <= "0005.9979.0135" Then
        DB_Upgrade_0005_9979_0135_To_0005_9979_0136
    End If
    
    If gstrDBVersion <= "0005.9979.0136" Then
        DB_Upgrade_0005_9979_0136_To_0005_9979_0137
    End If
    
    If gstrDBVersion <= "0005.9979.0140" Then
        DB_Upgrade_0005_9979_0137_To_0005_9979_0141
    End If
    
    If gstrDBVersion <= "0005.9979.0141" Then
        DB_Upgrade_0005_9979_0141_To_0005_9979_0142
    End If
    
    If gstrDBVersion <= "0005.9979.0142" Then
        DB_Upgrade_0005_9979_0142_To_0005_9979_0143
    End If
    
    If gstrDBVersion <= "0005.9979.0143" Then
        DB_Upgrade_0005_9979_0143_To_0005_9979_0144
    End If
    
    If gstrDBVersion <= "0005.9979.0144" Then
        DB_Upgrade_0005_9979_0144_To_0005_9979_0145
    End If
    
    If gstrDBVersion <= "0005.9979.0145" Then
        DB_Upgrade_0005_9979_0145_To_0005_9979_0146
    End If
    
    If gstrDBVersion <= "0005.9979.0146" Then
        DB_Upgrade_0005_9979_0146_To_0005_9979_0147
    End If
    
    If gstrDBVersion <= "0005.9979.0147" Then
        DB_Upgrade_0005_9979_0147_To_0005_9979_0148
    End If
    
    If gstrDBVersion <= "0005.9979.0148" Then
        DB_Upgrade_0005_9979_0148_To_0005_9979_0149
    End If
    
    If gstrDBVersion <= "0005.9979.0149" Then
        DB_Upgrade_0005_9979_0149_To_0005_9979_0150
    End If
    
    If gstrDBVersion <= "0005.9979.0150" Then
        DB_Upgrade_0005_9979_0150_To_0005_9979_0151
    End If
    
    If gstrDBVersion <= "0005.9979.0151" Then
        DB_Upgrade_0005_9979_0151_To_0005_9979_0152
    End If
    
    If gstrDBVersion <= "0005.9979.0153" Then
        DB_Upgrade_0005_9979_0152_To_0005_9979_0154
    End If
    
    If gstrDBVersion <= "0005.9979.0154" Then
        DB_Upgrade_0005_9979_0154_To_0005_9979_0155
    End If
    
    If gstrDBVersion <= "0005.9979.0155" Then
        DB_Upgrade_0005_9979_0155_To_0005_9979_0156
    End If
    
    If gstrDBVersion <= "0005.9979.0156" Then
        DB_Upgrade_0005_9979_0156_To_0005_9979_0157
    End If
    
    If gstrDBVersion <= "0005.9979.0157" Then
        DB_Upgrade_0005_9979_0157_To_0005_9979_0158
    End If
    
    Go_to_UpgradeDB4_module
    
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub DB_Upgrade_5_9974_00_To_5_9977_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('YearsHistoryToInclude_Min', " & _
                          " 10, " & _
                          " 'Initial value = 10')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('YearsHistoryToInclude_Accounts', " & _
                          " 10, " & _
                          " 'Initial value = 10')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('YearsHistoryToInclude_Calendar', " & _
                          " 30, " & _
                          " 'Initial value = 30')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('YearsHistoryToInclude_Meetings', " & _
                          " 5, " & _
                          " 'Initial value = 5')"
                          
    CMSDB.Execute "DELETE FROM tblConstants WHERE FldName = 'YearRangeCheckPublicSpeaker' AND NumVal = 160 "
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9977.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_5_9977_00_To_5_9977_02()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CreateField ErrorCode, "tblPublicMtgWtg", "DoesBoth", "YESNO"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9977.2"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9977_02_To_5_9977_03()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('IncludeInactiveAsIrregular', " & _
                          " FALSE, " & _
                          " 'Initial value = FALSE')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9977.3"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9977_03_To_5_9978_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field

    CreateField ErrorCode, "tblPrintAddresses", "GroupName", "TEXT", "250"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblPrintAddresses").Fields("GroupName").Required = False
    CMSDB.TableDefs.Refresh
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9978.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9978_00_To_5_99778_01()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MaxResultRows', " & _
                          " 2000, " & _
                          " 'Initial value = 2000')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('AccountsWarningsDefault', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9978.1"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9978_01_To_5_99779_00()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('UseMsgBoxForMessages', " & _
                          " FALSE, " & _
                          " 'Initial value = FALSE')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('SuppressMessages', " & _
                          " FALSE, " & _
                          " 'Initial value = FALSE')"
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.0"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9979_00_To_5_99779_01()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('NewTransaction_SkipDesc', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE. (Skip description on Tran Entry form after enter tran Code) ')"
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.1"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9979_01_To_5_99779_02()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('MaxImageSizeKB_Map', " & _
                          " 1000, " & _
                          " 'Initial value = 1000') "
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.2"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_5_9979_02_To_5_9979_08()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = -0.3 " & _
                  " WHERE FldName = 'CongRepTopMargin' "
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.8"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_5_9979_08_To_5_9979_09()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('AutoSelectPersonIfOnlyMatch', " & _
                          " True, " & _
                          " 'Initial value = True; applies to Personal Details form') "
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.9"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_5_9979_09_To_5_9979_10()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, NewField As DAO.Field
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakX_Side2_4_05', " & _
                          " 0, " & _
                          " ' ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES('PubCardTweakY_Side2_4_05', " & _
                          " 0.05, " & _
                          " ' ')"
                          
    CreateField ErrorCode, "tblPubCardTypes", "CardSideInfo", "TEXT", "250"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblPubCardTypes").Fields("CardSideInfo").Required = False
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblPubCardTypes SET CardSideInfo = " & _
                "'Side 1 is printed higher than side 2' " & _
                "WHERE CardTypeID = 1"
                              
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "5.9979.10"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0010_To_0005_9979_0011()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "ALTER TABLE tblTransactionTypes " & _
                    "ALTER COLUMN TranCode TEXT(10)" & ";"
                           
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0011"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0011_To_0005_9979_0012()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  "AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('PubRecCardVersion_CongMin', " & _
                          " 1, " & _
                          "'5/02', " & _
                          " 'Initial value = 1 (ie S-21 5/02)')"
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0012"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0012_To_0005_9979_0013()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CreateField ErrorCode, "tblTransactionDates", "BookGroupNo", "LONG"
    
    CMSDB.Execute "UPDATE tblTransactionDates SET BookGroupNo = 0"
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0013"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0013_To_0005_9979_0014()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  "AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('TranTypesContributedAtGroups', " & _
                          "'C,'," & _
                          " 'Initial value = C,')"
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0014"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0014_To_0005_9979_0015()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                             
    CMSDB.Execute "UPDATE tblConstants SET AlphaVal = 'C,G,' " & _
                  "WHERE FldName = 'TranTypesContributedAtGroups'"
                          
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0015"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0015_To_0005_9979_0016()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property
    
    DeleteTable "tblTransactionSubTypes"

    CreateTable ErrorCode, "tblTransactionSubTypes", "Description", "TEXT", "255", , True, "TranSubCodeID"
    CreateField ErrorCode, "tblTransactionSubTypes", "LinkedToTranCode", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblTransactionSubTypes")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Describes various sub-types of transaction (eg 'Refurb fund' as sub-type of 'Congregation Contributions')")
    Else
        prp.value = "Describes various sub-types of transaction (eg 'Refurb fund' as sub-type of 'Congregation Contributions')"
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap
                          
'****************************************************************************************************

    CreateField ErrorCode, "tblTransactionDates", "TranSubTypeID", "LONG", , ""
    
    CMSDB.Execute "UPDATE tblTransactionDates SET TranSubTypeID = 0"
                              
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0016"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0016_To_0005_9979_0017()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('CongContributionTransactionCode', " & _
                          " 'C', " & _
                          " 'Initial value = C')"
                                                       
    CMSDB.Execute "UPDATE tblConstants SET AlphaVal = 'C,' " & _
                  "WHERE FldName = 'TranTypesContributedAtGroups'"
                                                    
                              
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0017"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0017_To_0005_9979_0018()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim tdf As TableDef, prp As DAO.Property

    DeleteTable "tblPubCardSideAndYearForPerson"

    CreateTable ErrorCode, "tblPubCardSideAndYearForPerson", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblPubCardSideAndYearForPerson", "CardTypeID", "LONG"
    CreateField ErrorCode, "tblPubCardSideAndYearForPerson", "ServiceYear", "LONG"
    CreateField ErrorCode, "tblPubCardSideAndYearForPerson", "CardSide", "LONG"
    
    CMSDB.TableDefs.Refresh
    
    CreateIndex ErrorCode, "tblPubCardSideAndYearForPerson", "PersonID, CardTypeID, ServiceYear", _
                "IX1", True, False, True
    
    CMSDB.TableDefs.Refresh
    
    Set tdf = CMSDB.TableDefs("tblPubCardSideAndYearForPerson")
    
    On Error Resume Next
    
    Set prp = tdf.Properties("Description")
    
    If Err.number <> 0 Then
        tdf.Properties.Append tdf.CreateProperty("Description", _
            dbText, "Stores the side of the Pub Rec Card for person and year.")
    Else
        prp.value = "Stores the side of the Pub Rec Card for person and year."
    End If
    
    CMSDB.TableDefs.Refresh
    
    On Error GoTo ErrorTrap

     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0018"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0018_To_0005_9979_0019()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMTP_UserName', " & _
                          " '', " & _
                          " 'Initial value = BLANK')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMTP_Password', " & _
                          " '', " & _
                          " 'Initial value = BLANK')"
                                                                                                      
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMTP_SuccessCode', " & _
                          " '', " & _
                          " 'Initial value = 334')"
                              
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0019"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0019_To_0005_9979_0020()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    FixMeetingTypes
                              
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0020"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0020_To_0005_9979_0021()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('CSVExportHasDoubleQuotes', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
                                                        
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0021"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0021_To_0005_9979_0022()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblTransactionDates SET TranSubTypeID = 0 WHERE TranSubTypeID IS NULL "
                                                        
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0022"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_0005_9979_0022_To_0005_9979_0023()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'meeting durations in minutes
    '
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('PublicMeetingDurationMins', " & _
                          "35, " & _
                          " '')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('WatchtowerDurationMins', " & _
                          "70, " & _
                          " '')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSDurationMins', " & _
                          "35, " & _
                          " '')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('CongBibleStudyDurationMins', " & _
                          "30, " & _
                          " '')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServiceMeetingDurationMins', " & _
                          "40, " & _
                          " '')"
                          
    '
    'Now build new tblCongBibleStudyRota table
    '
    DeleteTable "tblCongBibleStudyRota"
    CreateTable ErrorCode, "tblCongBibleStudyRota", "MeetingDate", "DATE", , , False
    CreateField ErrorCode, "tblCongBibleStudyRota", "ConductorID", "LONG"
    CreateField ErrorCode, "tblCongBibleStudyRota", "ReaderID", "LONG"
       
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblCongBibleStudyRota " & _
                  "   (MeetingDate) " & _
                  "WITH PRIMARY"
                  
    CMSDB.Execute "UPDATE tblTasks SET Description = 'Congregation Bible Study Conductor' WHERE Task = 14"
    CMSDB.Execute "UPDATE tblTasks SET Description = 'Congregation Bible Study Assistant' WHERE Task = 15"
    CMSDB.Execute "UPDATE tblTasks SET Description = 'Congregation Bible Study Reader' WHERE Task = 16"
    CMSDB.Execute "UPDATE tblTasks SET Description = 'Congregation Bible Study Prayer' WHERE Task = 17"
                  
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0023"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0023_To_0005_9979_0024()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "DELETE FROM tblTaskAndPerson WHERE Person = 0 "
                                                        
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0024"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0024_To_0005_9979_0025()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblTasks", "RequiresExemplaryBro", "YESNO", ""
     
    CMSDB.Execute "UPDATE tblTasks SET RequiresExemplaryBro = TRUE " & _
                  "WHERE Task NOT IN (90,86,88,95,54,55,56,83)"
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('AlertForIrregularBrosWithCongDuties', " & _
                          "TRUE, " & _
                          " '')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0025"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0024_To_0005_9979_0026()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ShowGiftAidPayerInTranEntryForm', " & _
                          "TRUE, " & _
                          " '')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0026"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0026_To_0005_9979_0027()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                              
    '
    'Now build new tblReportSentToBranch table
    '
    DeleteTable "tblReportSentToBranch"
    CreateTable ErrorCode, "tblReportSentToBranch", "SocietyReportingPeriod", "DATE", , , False
       
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblReportSentToBranch " & _
                  "   (SocietyReportingPeriod) " & _
                  "WITH PRIMARY"
                  
    CMSDB.Execute "INSERT INTO tblReportSentToBranch " & _
                    "SELECT DISTINCT SocietyReportingPeriod FROM tblMinReports " & _
                    "WHERE DATEDIFF('m', SocietyReportingPeriod, Now) > 2 "
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0027"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0027_To_0005_9979_0028()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('WebsiteForMinReports', " & _
                          "'www.jw.org', " & _
                          " '')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0028"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0028_To_0005_9979_0029()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ShowZeroAmountTransactions', " & _
                          "FALSE, " & _
                          " 'FALSE')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0029"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_0005_9979_0029_To_0005_9979_0030()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                                            
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('AccountsCheckDupeEntries', " & _
                          "TRUE, " & _
                          " 'TRUE')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0030"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub DB_Upgrade_0005_9979_0030_To_0005_9979_0031()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                                            
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('AddPublicSpeakerOutline', " & _
                          "'ASK', " & _
                          " 'Values: ASK, AUTO, NO')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0031"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0031_To_0005_9979_0032()
                              
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                                            
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('AccountsReportingPrintHeightCM', " & _
                          "29, " & _
                          " 'This is landscape, so actually refers to A4 height. Default: 29 ')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('AccountsReportingPrintWidthCM', " & _
                          "20.5, " & _
                          " 'This is landscape, so actually refers to A4 Width. Default: 20.5 ')"
     
     '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0032"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub DB_Upgrade_0005_9979_0032_To_0005_9979_0033()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('DayOfMonthForCongStats', " & _
                          " 1, " & _
                          " 'Initial value = 1')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0033"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0033_To_0005_9979_0034()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    '
    'Add new Attendant Overseer Task
    '
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description) " & _
                  " VALUES (6, " & _
                          " 11, " & _
                          " 96, " & _
                          " 'Attendants Overseer')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0034"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0034_To_0005_9979_0035()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DropField ErrorCode, "tblEventAutoAlarms", "TriggerSQL"
    DropField ErrorCode, "tblEventAutoAlarms", "AccessLevels"
    CreateField ErrorCode, "tblEventAutoAlarms", "TriggerSQL", "MEMO"
    CreateField ErrorCode, "tblEventAutoAlarms", "AccessLevels", "TEXT", "100"
    
    CMSDB.TableDefs("tblEventAutoAlarms").Fields("TriggerSQL").AllowZeroLength = True
    CMSDB.TableDefs("tblEventAutoAlarms").Fields("AccessLevels").AllowZeroLength = True
    
    CMSDB.Execute "UPDATE tblEventAutoAlarms SET TriggerSQL = ''"
    CMSDB.Execute "UPDATE tblEventAutoAlarms SET AccessLevels = ''"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0035"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0035_To_0005_9979_0037()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    'datafix
    CMSDB.Execute "UPDATE tblPublisherDates SET EndReason = -1 WHERE EndDate = #9999/12/01#"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0037"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0037_To_0005_9979_0038()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CMSDB.Execute "UPDATE tblAccessLevelDescriptions SET AccessDesc = 'Public Talks and Service Meeting' WHERE AccessLevel = 8"
    
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmMainMenu' " & _
                  "AND ControlNameProperty = 'cmdServiceMtg' " & _
                  "AND SecurityLevel = 5"
                  
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmMainMenu', 'cmdServiceMtg', 8)"
                  
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmExportDB', 'chkExportItem(8)', 8)"
                  
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmExportDB' " & _
                  "AND ControlNameProperty = 'chkExportItem(8)' " & _
                  "AND SecurityLevel = 5"
                  
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmMainMenu' " & _
                  "AND ControlNameProperty = 'cmdAccounts' " & _
                  "AND SecurityLevel = 5"
        
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmPublicMeetingMenu' " & _
                  "AND ControlNameProperty = 'cmdChairmansNotes' " & _
                  "AND SecurityLevel = 5"
                  
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmOptionsMenu', 'cmdFieldMinAdvanced', 11)"
                  
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmMainMenu' " & _
                  "AND ControlNameProperty = 'cmdFieldMinistry' " & _
                  "AND SecurityLevel = 5"
                  
    CMSDB.Execute "DELETE FROM tblObjectSecurity " & _
                  "WHERE FormNameProperty = 'frmFieldMinistryMenu' " & _
                  "AND ControlNameProperty = 'cmdFieldServiceReports' " & _
                  "AND SecurityLevel = 5"
                  
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmMainMenu', 'cmdOpenCalendar', 11)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmMainMenu', 'cmdOpenCalendar', 8)"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0038"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0038_To_0005_9979_0039()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CMSDB.Execute "UPDATE tblAccessLevelDescriptions SET AccessDesc = 'Public Talks and Service Meeting' WHERE AccessLevel = 8"
    CMSDB.Execute "INSERT INTO tblAccessLevelDescriptions " & _
                  "VALUES (110, 12, 'Personal Details')"

    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmMainMenu', 'cmdOpenPersonalDetails', 12)"
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0039"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0039_To_0005_9979_0040()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmExportDB', 'chkExportItem(11)', 1)"
    CMSDB.Execute "INSERT INTO tblObjectSecurity (FormNameProperty, ControlNameProperty, SecurityLevel) " & _
                  "VALUES ('frmExportDB', 'chkExportItem(11)', 5)"
                  
    CMSDB.Execute "INSERT INTO tblExportDetails " & _
                  "(ExportDataType, " & _
                  " OrderingForSQL, " & _
                  " IncludeForExport, " & _
                  " Description) " & _
                  " VALUES (12, " & _
                          " 1200, " & _
                          "FALSE, " & _
                          "'User Defined Queries and Rotas')"
                  
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ImportItem11', " & _
                          " False, " & _
                          " 'Initial value = FALSE (User-defined queries and rotas)')"
                  
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0040"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0040_To_0005_9979_0041()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
    
    CreateField ErrorCode, "tblUserQueries", "Private", "YESNO"
    CreateField ErrorCode, "tblCustomRotaDetails", "Private", "YESNO"
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0041"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0041_To_0005_9979_0042()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                   
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSMinAge', " & _
                          " 6, " & _
                          " 'Initial value = 6')"
                  
                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0042"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0042_To_0005_9979_0043()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                
    CreateField ErrorCode, "tblPublicMtgSchedule", "Provisional", "YESNO"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0043"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0043_To_0005_9979_0044()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                
    DeleteTable "tblOutlookSync"
    CreateTable ErrorCode, "tblOutlookSync", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblOutlookSync", "OutlookEntryID", "TEXT", "255"
    CMSDB.TableDefs.Refresh
    CreateIndex ErrorCode, "tblOutlookSync", "PersonID", "IX1", True, False, True
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('OutlookSynch_AllowEmailAddrAccess', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
    
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0044"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0044_To_0005_9979_0045()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    DeleteTable "tblBankAccounts"
    CreateTable ErrorCode, "tblBankAccounts", "AccountID", "LONG", , , False
    CreateField ErrorCode, "tblBankAccounts", "AccountName", "TEXT", "255"
    CreateField ErrorCode, "tblBankAccounts", "StartAmount", "SINGLE"
    CreateField ErrorCode, "tblBankAccounts", "StartDate", "DATE"
    CMSDB.TableDefs.Refresh
    CreateIndex ErrorCode, "tblBankAccounts", "AccountID", "IX1", True, False, True
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "INSERT INTO tblBankAccounts " & _
                  "(AccountID, " & _
                  " AccountName, " & _
                  " StartAmount," & _
                  " StartDate) " & _
                  " VALUES (0, " & _
                          " 'Current Account', " & _
                          GlobalParms.GetValue("AccountBalanceAtStartOfMonth", "NumFloat", 0) & ", " & _
                          GetDateStringForSQLWhere(GlobalParms.GetValue("AccountBalanceAtStartOfMonth", "DateVal", "01/01/2000")) & ")"
        
    CMSDB.Execute "DELETE FROM tblConstants WHERE FldName = 'AccountBalanceAtStartOfMonth'"
      
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0045"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0045_To_0005_9979_0046()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTransactionDates", "AccountID", "LONG"
    CreateField ErrorCode, "tblTransactionDates", "TfrAccountID", "LONG"
            
    CMSDB.Execute "UPDATE tblTransactionDates SET AccountID = 0, TfrAccountID = -1"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0046"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0046_To_0005_9979_0047()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTransactionTypes", "AccountID", "LONG"
    CreateField ErrorCode, "tblTransactionTypes", "TfrAccountID", "LONG"
            
    CMSDB.Execute "UPDATE tblTransactionTypes SET AccountID = 0, TfrAccountID = -1"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0047"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0047_To_0005_9979_0048()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('AccountTransferTransactionCode', " & _
                          " 'TX', " & _
                          " 'Initial value = TX')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0048"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0048_To_0005_9979_0049()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset

    CMSDB.Execute "INSERT INTO tblTransactionTypes " & _
                  "(TranCode, " & _
                  " Description, " & _
                  " InOutTypeID, " & _
                  " AutoDayOfMonth, " & _
                  " Amount, " & _
                  " OnReceipt, " & _
                  " Ref, " & _
                  " AccountID, " & _
                  " TfrAccountID) " & _
                  " VALUES ('TX', " & _
                          " 'Account Transfer IN', " & _
                          " 1, 0, 0, false, 0, 0, -1)"
                          
    CMSDB.Execute "INSERT INTO tblTransactionTypes " & _
                  "(TranCode, " & _
                  " Description, " & _
                  " InOutTypeID, " & _
                  " AutoDayOfMonth, " & _
                  " Amount, " & _
                  " OnReceipt, " & _
                  " Ref, " & _
                  " AccountID, " & _
                  " TfrAccountID) " & _
                  " VALUES ('TX', " & _
                          " 'Account Transfer OUT', " & _
                          " 2, 0, 0, false, 0, 0, -1)"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0049"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0049_To_0005_9979_0050()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('OutlookSynch_RemoveCMSFldWhenBlankOutlookFld', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
                          
    SetUpTranCode False, "W", "Worldwide Work Contributions Received", 5, 28, 0, True, 0, 0, -1
    SetUpTranCode False, "G", "Gift Aid Contributions Received", 1, 0, 0, True, 0, 0, -1
    SetUpTranCode False, "C", "Congregation Contributions Received", 1, 0, 0, True, 0, 0, -1
    SetUpTranCode False, "K", "Society Kingdom Hall Fund Contributions Received", 5, 28, 0, True, 0, 0, -1
    SetUpTranCode False, "U", "Kingdom Hall Operation and Maintenance", 4, 1, -96, False, 0, 0, -1
    SetUpTranCode False, "O", "Other Payment", 2, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "I", "Bank Interest Received", 1, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "X", "Tax Repayment", 1, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "W", "Regular Donation To IBSA for WWW", 4, 28, -10, False, 0, 0, -1
    SetUpTranCode False, "K", "Regular Donation to WBTS for SKHF", 4, 28, -15, False, 0, 0, -1
    SetUpTranCode False, "O", "Other Receipt", 1, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "S", "Regular Donation to WBTS", 4, 28, -10, False, 0, 0, -1
    SetUpTranCode False, "A", "Donations from Cong Funds (eg Pio School, circuit)", 2, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "V", "CO and Visiting Speaker Expenses", 2, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "Q", "Equipment Purchased for the Hall", 2, 0, 0, False, 0, 0, -1
    SetUpTranCode False, "TO", "Travelling Overseer Assistance Arrangement", 4, 28, -19, False, 0, 0, -1

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('OutlookSynch_CategoryFilter', " & _
                          " 'JWs', " & _
                          " 'Initial value = JWs')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0050"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0050_To_0005_9979_0051()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    DeleteTable "tblTableUpdateDateTimes"
    CreateTable ErrorCode, "tblTableUpdateDateTimes", "TableName", "TEXT", "150"
    CreateField ErrorCode, "tblTableUpdateDateTimes", "LastUpdateDate", "DATE"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0051"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0051_To_0005_9979_0052()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CMSDB.Execute "UPDATE tblMeetingTypes SET MeetingType = 'Congregation Bible Study' WHERE MeetingTypeID = 4"
    
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro) " & _
                  " VALUES (5, " & _
                          " 9, " & _
                          " 97, " & _
                          " 'Field Service Group Overseer', FALSE, TRUE)"
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro) " & _
                  " VALUES (5, " & _
                          " 9, " & _
                          " 98, " & _
                          " 'Field Service Group Assistant', FALSE, TRUE)"
                          
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0052"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0052_To_0005_9979_0053()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTasks", "TaskComment", "TEXT", , ""
    
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 99, " & _
                          " 'Talk #1 (Reading Assignment)', TRUE, FALSE, " & _
                          " 'For TMS in 2009 onwards') "

    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 100, " & _
                          " 'Talk #2',  TRUE, FALSE, " & _
                          " 'For TMS in 2009 onwards') "
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro, " & _
                  " TaskComment) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 101, " & _
                          " 'Talk #3', TRUE, FALSE,  " & _
                          " 'For TMS in 2009 onwards') "
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0053"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0053_To_0005_9979_0054()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    'transfer TMS students to new 2009 roles
    
    CMSDB.Execute "DELETE FROM tblTaskAndPerson WHERE Task IN (99, 100, 101)"
    CMSDB.Execute "DELETE FROM tblTaskPersonSuspendDates WHERE Task IN (99, 100, 101)"
    
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 99, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T1 " & _
                  "WHERE Task = 36"
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 99, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T2 " & _
                  "WHERE Task = 37 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 99)"
                
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 100, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T1 " & _
                  "WHERE Task = 38"
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 100, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T2 " & _
                  "WHERE Task = 39 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 100)"
                                
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 101, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T1 " & _
                  "WHERE Task = 40 "
    CMSDB.Execute "INSERT INTO tblTaskAndPerson " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person,OnSunday,OnMidweek) " & _
                  "SELECT CongNo, 4, 6, 101, Person, FALSE, FALSE " & _
                  "FROM tblTaskAndPerson T2 " & _
                  "WHERE Task = 41 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskAndPerson T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 101)"
''''
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 99, Person " & _
                  "FROM tblTaskPersonSuspendDates T1 " & _
                  "WHERE Task = 36"
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 99, Person " & _
                  "FROM tblTaskPersonSuspendDates T2 " & _
                  "WHERE Task = 37 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskPersonSuspendDates T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 99)"
                
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 100, Person " & _
                  "FROM tblTaskPersonSuspendDates T1 " & _
                  "WHERE Task = 38"
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 100, Person " & _
                  "FROM tblTaskPersonSuspendDates T2 " & _
                  "WHERE Task = 39 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskPersonSuspendDates T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 100)"
                                
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 101, Person " & _
                  "FROM tblTaskPersonSuspendDates T1 " & _
                  "WHERE Task = 40 "
    CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                  " (CongNo,TaskCategory,TaskSubCategory,Task,Person) " & _
                  "SELECT CongNo, 4, 6, 101, Person " & _
                  "FROM tblTaskPersonSuspendDates T2 " & _
                  "WHERE Task = 41 " & _
                  "AND NOT EXISTS (SELECT 1 FROM tblTaskPersonSuspendDates T3 " & _
                                  "WHERE T3.Person = T2.Person " & _
                                  "AND T3.Task = 101)"

                                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0054"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0054_To_0005_9979_0055()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "No2AssistantName", "TEXT", "255"
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "No2Setting", "TEXT", "255"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0055"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0055_To_0005_9979_0056()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CMSDB.TableDefs("tblTMSPrintWorkSheet").Fields("No2AssistantName").Required = False
    CMSDB.TableDefs("tblTMSPrintWorkSheet").Fields("No2Setting").Required = False
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0056"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0056_To_0005_9979_0057()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTMSPrintSchedule", "No1BroSchool2", "TEXT", "100"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No1BroSchool3", "TEXT", "100"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No2AsstSchool1", "TEXT", "100"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No2AsstSchool2", "TEXT", "100"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No2AsstSchool3", "TEXT", "100"
    
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No1BroSchool2").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No1BroSchool3").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No2AsstSchool1").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No2AsstSchool2").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No2AsstSchool3").Required = False
                
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSPrintStudentDtlsFor2009Format', " & _
                          " TRUE, " & _
                          " 'Initial value = TRUE')"
    
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0057"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0057_To_0005_9979_0058()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                    
    CMSDB.Execute "UPDATE tblTMSAssignmentsForSearch " & _
                  "SET AssignmentDesc = 'Bible Highlights' " & _
                  "WHERE SeqNum = 3"
    CMSDB.Execute "UPDATE tblTMSAssignmentsForSearch " & _
                  "SET AssignmentDesc = 'Talk No 1' " & _
                  "WHERE SeqNum = 4"
    
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0058"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0058_To_0005_9979_0059()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSTMSNo1Weighting_2009', " & _
                          " 1, " & _
                          " 'Initial value = 1')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSTMSNo2Weighting_2009', " & _
                          " 5, " & _
                          " 'Initial value = 1')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMSTMSNo3Weighting_2009', " & _
                          " 5, " & _
                          " 'Initial value = 1')"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0059"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0059_To_0005_9979_0060()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                    
    CreateField ErrorCode, "tblTMSTalkNoDesc", "Order2009", "LONG"
    
    CMSDB.Execute "UPDATE tblTMSTalkNoDesc " & _
                  "SET Order2009 = SeqNum "
                  
    CMSDB.Execute "UPDATE tblTMSTalkNoDesc " & _
                  "SET Order2009 = 4 " & _
                  "WHERE SeqNum = 3"
    CMSDB.Execute "UPDATE tblTMSTalkNoDesc " & _
                  "SET Order2009 = 3 " & _
                  "WHERE SeqNum = 4"
                    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0060"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0060_To_0005_9979_0061()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMS_HTTPPhoneNoParmName', " & _
                          " 'to_num', " & _
                          " 'Initial value = to_num')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMS_HTTPMessageParmName', " & _
                          " 'message', " & _
                          " 'Initial value = message')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMS_HTTPUserNameParmName', " & _
                          " 'username', " & _
                          " 'Initial value = username')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMS_HTTPPasswordParmName', " & _
                          " 'password', " & _
                          " 'Initial value = password')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('SMS_HTTPWordDelimiter', " & _
                          " '+', " & _
                          " 'Initial value = +')"
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('SendSMSUsingHTTP', " & _
                          " False, " & _
                          " 'Initial value = False')"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0061"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0061_To_0005_9979_0062()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DeleteTable "tblTMSItemsMaster"
                    
    CMSDB.Execute "SELECT * INTO tblTMSItemsMaster FROM tblTMSItems"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0062"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0062_To_0005_9979_0063()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblCongBibleStudyRota", "PrayerID", "LONG"
    CMSDB.Execute "UPDATE tblCongBibleStudyRota SET PrayerID = 0"
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0063"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0063_To_0005_9979_0064()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgStartMins_2009', " & _
                            65 & ", " & _
                          " 'Initial value = 65')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgCOVisitStartMins_2009', " & _
                            35 & ", " & _
                          " 'Initial value = 35')"
       
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0064"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0064_To_0005_9979_0065()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('OutlookCategoryForImport', " & _
                            "'' , " & _
                          " 'Initial value = BLANK')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0065"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0065_To_0005_9979_0066()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateTable ErrorCode, "tblMidWkMtgTempStartTime", "MeetingDate", "DATE"
    CreateField ErrorCode, "tblMidWkMtgTempStartTime", "NewTime", "DATE"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0066"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0066_To_0005_9979_0067()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    Set fld = CMSDB.TableDefs("tblCong").CreateField("SeqNo", dbLong)
    fld.Attributes = fld.Attributes Or dbAutoIncrField
    CMSDB.TableDefs("tblCong").Fields.Append fld
    CMSDB.TableDefs("tblCong").Fields.Refresh
    
    Set fld = Nothing
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0067"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0067_To_0005_9979_0068()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSWarnAboutSettingsInCounselDialog', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0068"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0068_To_0005_9979_0069()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('ExportItem11', " & _
                            "FALSE , " & _
                          " 'Initial value = FALSE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0069"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0069_To_0005_9979_0070()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblTasks " & _
                  "SET AllowSuspend = FALSE " & _
                  "WHERE Task IN (33,35,36,37,38,39,40,41,42,47)"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0070"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0070_To_0005_9979_0071()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    DropIndex ErrorCode, "tblMeetingTypes", "IX1"
    
    CMSDB.Execute "UPDATE tblMeetingTypes " & _
                  "SET MeetingTypeID = 4, " & _
                  "    AltOrder = 4 " & _
                  "WHERE MeetingType = 'Theocratic Ministry School' "
    CMSDB.Execute "UPDATE tblMeetingTypes " & _
                  "SET MeetingTypeID = 2, " & _
                  "    AltOrder = 0 " & _
                  "WHERE MeetingType = 'Congregation Bible Study' "
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0071"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0071_To_0005_9979_0072()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    
    CreateField ErrorCode, "tblTMSPrintSchedule", "BHTheme", "TEXT", "200"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No1Theme", "TEXT", "200"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No2Theme", "TEXT", "200"
    CreateField ErrorCode, "tblTMSPrintSchedule", "No3Theme", "TEXT", "200"
    
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("BHTheme").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No1Theme").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No2Theme").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No3Theme").Required = False
    CMSDB.TableDefs.Refresh
    
    For Each fld In CMSDB.TableDefs("tblTMSPrintSchedule").Fields
        If fld.Type = dbText Then
            fld.AllowZeroLength = True
        End If
    Next
    
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSShowThemeInSchedulePrint_2009', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0072"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0072_To_0005_9979_0073()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    
    CreateField ErrorCode, "tblCongBibleStudyRota", "OpeningSong", "LONG"
    
    CMSDB.Execute "UPDATE tblCongBibleStudyRota SET OpeningSong = 0"
    
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0073"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0073_To_0005_9979_0074()
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSUsePortionOfScheduleForPrevNext', " & _
                            "TRUE , '6,3', " & _
                          " 'Initial value = TRUE; 6,3 (months prev/next)')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0074"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0074_To_0005_9979_0075()
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSAutoSwitchToAsstTabOnInsert', " & _
                            "TRUE, " & _
                          " 'Initial value = TRUE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0075"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0075_To_0005_9979_0076()
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSPrintSlipsInLandscape', " & _
                            "TRUE, " & _
                          " 'Initial value = TRUE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0076"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0076_To_0005_9979_0077()
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('HandleForeignChars', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
                                 
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0077"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0077_To_0005_9979_0078()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    CreateField ErrorCode, "tblTMSPrintSchedule", "No1Title_2009", "TEXT", "200"
    CMSDB.TableDefs.Refresh
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("No1Title_2009").Required = False
    CMSDB.TableDefs.Refresh

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0078"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0078_To_0005_9979_0079()
On Error GoTo ErrorTrap

    CMSDB.Execute "UPDATE tblTasks SET Description = 'Talk #2/#3 Assistant' WHERE Task = 43"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0079"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0079_To_0005_9979_0080()
Dim fso As New FileSystemObject
On Error Resume Next

'    Shell "regsvr32 /s /u " & """" & App.Path & "\smsSendMessage.dll" & """"
'    fso.DeleteFile App.Path & "\smsSendMessage.dll", True

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0080"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0080_To_0005_9979_0081()
On Error GoTo ErrorTrap

    CMSDB.Execute "UPDATE tblNameAddress SET DOB = 0 WHERE Active = FALSE"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0081"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0081_To_0005_9979_0082()
On Error GoTo ErrorTrap

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('PrintWholeMidWkMtg', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
                                 
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('PrintWholeMidWkMtgWithTMS', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0082"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0082_To_0005_9979_0083()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "OpeningPrayerBro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "CBSSong", "TEXT", "255"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "CBSConductor", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "CBSReader", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "BHBro", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1BroSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1BroSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1BroSch3", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2BroSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2BroSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2BroSch3", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3BroSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3BroSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3BroSch3", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1AsstSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1AsstSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No1AsstSch3", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2AsstSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2AsstSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No2AsstSch3", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3AsstSch1", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3AsstSch2", "TEXT", "100"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "No3AsstSch3", "TEXT", "100"
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("OpeningPrayerBro").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("CBSSong").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("CBSConductor").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("CBSReader").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("BHBro").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1BroSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1BroSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1BroSch3").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2BroSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2BroSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2BroSch3").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3BroSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3BroSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3BroSch3").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1AsstSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1AsstSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No1AsstSch3").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2AsstSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2AsstSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No2AsstSch3").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3AsstSch1").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3AsstSch2").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("No3AsstSch3").Required = False
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "ALTER TABLE tblServiceMtgSchedulePrint " & _
                    "ALTER COLUMN OpeningSong TEXT(255)" & ";"
    CMSDB.Execute "ALTER TABLE tblServiceMtgSchedulePrint " & _
                    "ALTER COLUMN ConcSong TEXT(255)" & ";"
    
    

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0083"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0083_To_0005_9979_0084()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field
    
    CMSDB.TableDefs("tblCongBibleStudyRota").Fields("ConductorID").Required = False
    CMSDB.TableDefs("tblCongBibleStudyRota").Fields("ReaderID").Required = False
    CMSDB.TableDefs("tblCongBibleStudyRota").Fields("PrayerID").Required = False
    CMSDB.TableDefs("tblCongBibleStudyRota").Fields("OpeningSong").Required = False
    CMSDB.TableDefs.Refresh

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0084"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0084_To_0005_9979_0085()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "CBSEndTime", "TEXT", "50"
    CreateField ErrorCode, "tblServiceMtgSchedulePrint", "TMSEndTime", "TEXT", "50"
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("CBSEndTime").Required = False
    CMSDB.TableDefs("tblServiceMtgSchedulePrint").Fields("TMSEndTime").Required = False
    
    CMSDB.TableDefs.Refresh
    
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 30 " & _
                  " WHERE FldName = 'CongBibleStudyDurationMins' "
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 40 " & _
                  " WHERE FldName = 'ServiceMeetingDurationMins' "
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 30 " & _
                  " WHERE FldName = 'TMSDurationMins' "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0085"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0085_To_0005_9979_0086()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 35 " & _
                  " WHERE FldName = 'ServMtgCOVisitStartMins_2009' "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0086"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0086_To_0005_9979_0087()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSWarnAboutMissingCounselInCounselDialog', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0087"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0087_To_0005_9979_0088()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSWarnAboutCloseAssignments', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSCloseAssignments_For_E_And_MS', " & _
                            "2, " & _
                          " 'Initial value = 2 (weeks)')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSCloseAssignments_For_Bros_1and3', " & _
                            "10, " & _
                          " 'Initial value = 10 (weeks)')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSCloseAssignments_For_Sisters_2and3', " & _
                            "16, " & _
                          " 'Initial value = 16 (weeks)')"


    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0088"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0088_To_0005_9979_0089()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " Comment) " & _
                  " VALUES ('NextTMSSchedulePrintStartDate', " & _
                          " 'Initial value = Null')"

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSScheduleDraftPrint', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0089"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0089_To_0005_9979_0090()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSAssignmentSlipsPrintAll', " & _
                            "FALSE , " & _
                          " 'Initial value = FALSE')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0090"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0090_To_0005_9979_0091()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSHighlightAsstDates', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0091"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0091_To_0005_9979_0092()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSSchedHighlightCurrentWk', " & _
                            "TRUE , " & _
                          " 'Initial value = TRUE')"
                          
                             

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0092"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0092_To_0005_9979_0093()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSNoSchoolsForSchedulePrint', " & _
                            "1, " & _
                          " 'Initial value = 1')"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0093"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0093_To_0005_9979_0094()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('ServMtgDefaultItemName', " & _
                            "'Item', " & _
                          " 'Initial value = Item')"
                          
    CMSDB.Execute "UPDATE tblTMSSchedule " & _
                    "SET CounselPointAssignedDate = NULL " & _
                    "WHERE CounselPointAssignedDate = 0"
                          
    CMSDB.Execute "UPDATE tblTMSSchedule " & _
                    "SET CounselPointCompletedDate = NULL " & _
                    "WHERE CounselPointCompletedDate = 0"
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0094"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0094_To_0005_9979_0095()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    GlobalParms.Save "TMSTMSNo1Weighting_2009", "NumFloat", 10
    GlobalParms.Save "TMSTMSNo2Weighting_2009", "NumFloat", 80
    GlobalParms.Save "TMSTMSNo3Weighting_2009", "NumFloat", 80
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0095"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0095_To_0005_9979_0096()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSNoSchoolsForCounsel', " & _
                            "1, " & _
                          " 'Initial value = 1')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0096"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0096_To_0005_9979_0097()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSHighlightsOnSch2', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSHighlightsOnSch3', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0097"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0097_To_0005_9979_0098()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblTMSPrintSchedule", "BHBroSch2", "TEXT", "100"
    CreateField ErrorCode, "tblTMSPrintSchedule", "BHBroSch3", "TEXT", "100"
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("BHBroSch2").Required = False
    CMSDB.TableDefs("tblTMSPrintSchedule").Fields("BHBroSch3").Required = False
    
    CMSDB.TableDefs.Refresh
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0098"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0098_To_0005_9979_0099()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "DELETE FROM tblTasks " & _
                  "WHERE Task IN (44,45,46) "
    CMSDB.Execute "DELETE FROM tblTaskAndPerson " & _
                  "WHERE Task IN (44,45,46) "
    CMSDB.Execute "DELETE FROM tblTaskPersonSuspendDates " & _
                  "WHERE Task IN (44,45,46) "
            
                  
    CMSDB.Execute "SELECT Person AS PersonID, 1 AS SchoolNo, Task " & _
                  "INTO tblTMSSchoolAndPerson " & _
                  "FROM tblTaskAndPerson " & _
                  "WHERE Task IN (99,100,101,43,34)"
    
    CreateIndex ErrorCode, "tblTMSSchoolAndPerson", "PersonID, SchoolNo, Task", _
                "IX1", True, False, True
    
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0099"

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0099_To_0005_9979_0100()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSDeleteFromUnusedSchoolsWhenNoSchoolsChanges', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0100"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0100_To_0005_9979_0101()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblTMSCounselPointComponents " & _
                  "SET SubPointDescription = 'Application made clear - isolate key words (p155)' " & _
                  "WHERE CounselPoint = 22 AND CounselSubPoint = 2"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0101"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0101_To_0005_9979_0102()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fso As New Scripting.FileSystemObject
Dim fsoTextStream As Scripting.TextStream, fsoFile As Scripting.File
Dim TheFileRec As String, SongNo As Long, SongName As String
    
    If Not fso.FileExists(JustTheDirectory & "\All Songs.txt") Then
        MsgBox "Cannot find 'All Songs.txt'", vbOKOnly + vbCritical, AppName
        EndProgram
    End If
    
    DelAllRows "tblSongNoAndSubject"
    DelAllRows "tblSongSubjects"
    DelAllRows "tblSongs"
    
    Set fsoFile = fso.GetFile(JustTheDirectory & "\All Songs.txt")
    Set fsoTextStream = fsoFile.OpenAsTextStream
    
    With fsoTextStream
    
    Do Until .AtEndOfStream
        TheFileRec = Trim(RemoveNonPrintingChars(.ReadLine()))
        Select Case True
        Case IsNumeric(Right(TheFileRec, 3))
            SongNo = CLng(Right(TheFileRec, 3))
            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 3)))
        Case IsNumeric(Right(TheFileRec, 2))
            SongNo = CLng(Right(TheFileRec, 2))
            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 2)))
        Case IsNumeric(Right(TheFileRec, 1))
            SongNo = CLng(Right(TheFileRec, 1))
            SongName = Trim$(DoubleUpSingleQuotes(Left(TheFileRec, Len(TheFileRec) - 1)))
        Case Else
            SongNo = 0
            SongName = ""
        End Select
        
        If SongNo > 0 Then
            CMSDB.Execute "INSERT INTO tblSongs " & _
                          "(SongNo, " & _
                          " SongTitle) " & _
                          " VALUES(" & SongNo & ", '" & _
                                       SongName & "')"
        End If
    Loop
    
    .Close
    
    End With
    
    fso.DeleteFile JustTheDirectory & "\All Songs.txt", True
    
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0102"
    
    Set fso = Nothing
    Set fsoTextStream = Nothing
    Set fsoFile = Nothing

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0102_To_0005_9979_0103()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSWarnOfUnprintedScheduleDays', " & _
                            "50, " & _
                          " 'Initial value = 50')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0103"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0103_To_0005_9979_0104()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, fld As DAO.Field

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET TrueFalse = FALSE, AlphaVal = 'DRAFT', Comment = 'Inital value=DRAFT; valid values: DRAFT, FINAL, REPRINT' " & _
                  " WHERE FldName = 'TMSScheduleDraftPrint' "

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0104"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0104_To_0005_9979_0105()
Dim TheError As Integer, fld As DAO.Field

    On Error Resume Next
    DeleteTable "tblTMSWeightings"
    On Error GoTo ErrorTrap
    CreateTable TheError, "tblTMSWeightings", "PersonID", "LONG"
    CreateField TheError, "tblTMSWeightings", "TMSWeighting", "DOUBLE"
    CreateField TheError, "tblTMSWeightings", "TMSPrayerWeighting", "DOUBLE"
    CreateField TheError, "tblTMSWeightings", "TMSAsstWeighting", "DOUBLE"
    
    CMSDB.TableDefs.Refresh
    
    
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 8, Comment = 'Initial value = 8 (weeks)' " & _
                  " WHERE FldName = 'TMSCloseAssignments_For_Bros_1and3' "
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 14, Comment = 'Initial value = 14 (weeks)' " & _
                  " WHERE FldName = 'TMSCloseAssignments_For_Sisters_2and3' "
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 30, Comment = 'Initial value = 30 (weeks)' " & _
                  " WHERE FldName = 'TMSWeeksToCheck' "
       

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0105"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0105_To_0005_9979_0106()
Dim TheError As Integer, fld As DAO.Field

Dim ErrorCode As Integer, NewField As DAO.Field

    DestroyGlobalObjects

    'new Volunteer column
    CMSDB.TableDefs.Refresh
    Set NewField = CMSDB.TableDefs("tblTMSSchedule").CreateField("IsVolunteer", dbBoolean)
    CMSDB.TableDefs("tblTMSSchedule").Fields.Append NewField
    NewField.Required = False
    
    
'    CreateField ErrorCode, "tblTMSSchedule", "IsVolunteer", "YESNO"

    CMSDB.TableDefs.Refresh
    
    SetUpGlobalObjects
    

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0106"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0106_To_0005_9979_0107()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSMonthsForRepeatAssistant', " & _
                            "12, " & _
                          " 'Initial value = 12')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0107"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0107_To_0005_9979_0108()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMS_NumberOfRowsToShowOnInsertForm', " & _
                            "0, " & _
                          " 'Initial value = 0 (0 means ALL)')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0108"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0108_To_0005_9979_0109()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 80, Comment = 'Initial value = 80' " & _
                  " WHERE FldName = 'TMSAsstWeighting' "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0109"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0109_To_0005_9979_0110()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblTMSPrintTypes VALUES (5, 'Counsel Forms')"
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0110"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0110_To_0005_9979_0111()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMS_MinAgeOfResponsibility', " & _
                            "16, " & _
                          " 'Initial value = 16')"
                          
    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 102, " & _
                          " 'School Assignments Early Notification', FALSE, FALSE)"
                              
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0111"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0111_To_0005_9979_0112()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, DateVal, " & _
                  " Comment) " & _
                  " VALUES ('TMS_FreqSuspendedReminderDays', " & _
                            "60, #" & Format(CDate(Now) - 60, "mm/dd/yyyy") & "#, " & _
                          " 'Initial value = 60')"
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0112"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0112_To_0005_9979_0113()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 1000000, Comment = 'Initial value = 1000000' " & _
                  " WHERE FldName = 'TMSWeightingIfAssistantOnly' "
                  
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 50, Comment = 'Initial value = 50' " & _
                  " WHERE FldName = 'TMSWeeksToCheck' "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0113"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0113_To_0005_9979_0114()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('Do3rdBackup', " & _
                            "FALSE, " & _
                          " 'Initial value = FALSE')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('3rdBackupLocation', " & _
                            "'', " & _
                          " 'Initial value = BLANK')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " Comment) " & _
                  " VALUES ('DocumentLocation2', " & _
                            "'', " & _
                          " 'Initial value = BLANK')"
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0114"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0114_To_0005_9979_0116()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMS_LastAsstWeighting', " & _
                            "5, " & _
                            "TRUE, " & _
                          " 'Initial values = 5, TRUE')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0116"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0116_To_0005_9979_0117()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMS_AwkwardCounselPoints', " & _
                            "'~43~45~46~47~', " & _
                            "TRUE, " & _
                          " 'Initial values = ~43~45~46~47~, TRUE (each SQ must be wrapped with ~nn~)')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0117"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0117_To_0005_9979_0118()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 103, " & _
                          " 'Oral Review Reader', TRUE, TRUE)"
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0118"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0118_To_0005_9979_0119()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMS_PrintSlipForReviewReader', " & _
                            "TRUE, " & _
                          " 'Initial value = TRUE ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0119"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub DB_Upgrade_0005_9979_0119_To_0005_9979_0120()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumFloat, " & _
                  " Comment) " & _
                  " VALUES ('TMS_ReviewReaderWeighting', " & _
                            "5, " & _
                          " 'Initial value = 5 ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0120"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0120_To_0005_9979_0121()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
   
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET TrueFalse = TRUE " & _
                  " WHERE FldName = 'TMSUseSubstituteSlips' "
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0121"

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub DB_Upgrade_0005_9979_0121_To_0005_9979_0122()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " AlphaVal, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMSPrintSlipsToFile', " & _
                            "'Microsoft XPS Document Writer', " & _
                            "FALSE, " & _
                          " 'Initial value = Microsoft XPS Document Writer; FALSE ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0122"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0122_To_0005_9979_0123()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSMaxNoMonthsForSistersTalks', " & _
                            "5, " & _
                          " 'Initial value = 5 ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0123"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0123_To_0005_9979_0124()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSNoMonthsSinceSistersLastTalk', " & _
                            "6, " & _
                          " 'Initial value = 7 ')"
                          
    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('TMSNoMonthsSinceBrosLastTalk', " & _
                            "5, " & _
                          " 'Initial value = 5 ')"
                          
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 700 " & _
                  " WHERE FldName = 'TMS_LastAsstWeighting' "
                          
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0124"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0124_To_0005_9979_0125()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
                          
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 40 " & _
                  " WHERE FldName = 'TMSTMSNo1Weighting_2009' "
                          
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0125"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0125_To_0005_9979_0126()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('RemoveUKInternationalDialCodes', " & _
                            "TRUE, " & _
                          " 'Initial value = TRUE ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0126"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0126_To_0005_9979_0127()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " NumVal, " & _
                  " Comment) " & _
                  " VALUES ('AutoSetupSrvMtgBlankTemplate', " & _
                            "TRUE, " & _
                            "2, " & _
                          " 'Initial value = TRUE;2 ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0127"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0127_To_0005_9979_0128()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblTasks " & _
                  "(TaskCategory, " & _
                  " TaskSubCategory, " & _
                  " Task, " & _
                  " Description, " & _
                  " AllowSuspend, " & _
                  " RequiresExemplaryBro) " & _
                  " VALUES (4, " & _
                          " 6, " & _
                          " 104, " & _
                          " 'Always mark slips as printed', FALSE, FALSE)"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0128"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0128_To_0005_9979_0129()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 40, Comment = 'Initial value = 40' " & _
                  " WHERE FldName = 'TMSAsstWeighting' "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0129"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0129_To_0005_9979_0130()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumFloat = 5, Comment = 'Initial value = 5' " & _
                  " WHERE FldName = 'TMS_LastAsstWeighting' "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0130"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0130_To_0005_9979_0131()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET AlphaVal = 'PDFCreator~!~Microsoft XPS Document Writer', Comment = 'Initial value = PDFCreator~!~Microsoft XPS Document Writer;FALSE' " & _
                  " WHERE FldName = 'TMSPrintSlipsToFile' "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0131"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0131_To_0005_9979_0132()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "DELETE FROM tblTMSAssignmentsForSearch " & _
                  " WHERE SeqNum IN (1,2,7) "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0132"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0132_To_0005_9979_0133()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "UPDATE tblTasks " & _
                    "SET AllowSuspend = FALSE, RequiresExemplaryBro = FALSE " & _
                    "WHERE Task = 39 "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0133"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0133_To_0005_9979_0134()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CreateField ErrorCode, "tblTMSItems", "Sourceless", "YESNO"
    CreateField ErrorCode, "tblTMSItemsMaster", "Sourceless", "YESNO"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0134"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0134_To_0005_9979_0135()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer

    CMSDB.Execute "INSERT INTO tblConstants " & _
                  "(FldName, " & _
                  " TrueFalse, " & _
                  " Comment) " & _
                  " VALUES ('TMS_EnableSourcelessFiltering', " & _
                            "TRUE, " & _
                          " 'Initial value = TRUE ')"
                          
                                  
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0135"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0135_To_0005_9979_0136()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
Dim fso As New FileSystemObject


    DeleteTable "tblPersonGroupingNames"
    DeleteTable "tblPersonGroupingMapping"
    
    '
    'tblPersonGroupingMapping
    '
    CreateTable ErrorCode, "tblPersonGroupingMapping", "PersonID", "LONG", , , False
    CreateField ErrorCode, "tblPersonGroupingMapping", "PersonGroupingID", "LONG"
    
    CMSDB.Execute "CREATE INDEX IX1 " & _
                  "ON tblPersonGroupingMapping " & _
                  "   (PersonID,PersonGroupingID) " & _
                  "WITH PRIMARY"
    
    '
    'tblPersonGroupingNames
    '
    CreateTable ErrorCode, "tblPersonGroupingNames", "PersonGroupingName", "TEXT", "100", , True, "PersonGroupingID"
    
    '
    'add new word templates folder
    '
    If Not fso.FolderExists(JustTheDirectory & "\Word Templates") Then
        fso.CreateFolder JustTheDirectory & "\Word Templates"
    End If
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0136"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0136_To_0005_9979_0137()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0137"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0137_To_0005_9979_0141()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    'no longer linking oral review reader to a school
    CMSDB.Execute "DELETE FROM tblTMSSchoolAndPerson " & _
              "WHERE Task = " & 103
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0141"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0141_To_0005_9979_0142()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 3, Comment = 'Initial value = TRUE;3' " & _
                  " WHERE FldName = 'AutoSetupSrvMtgBlankTemplate' "
                  
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 35, Comment = 'Initial value = 35 (includes 5 mins opening song' " & _
                  " WHERE FldName = 'CongBibleStudyDurationMins' "
        
    CMSDB.Execute "UPDATE tblTMSCounselPointComponents " & _
                  "SET SubPointDescription = 'In harmony with Faithful Slave (p224)' " & _
                  "WHERE CounselPoint = 40 AND CounselSubPoint = 2"
                  
    CMSDB.Execute "INSERT INTO tblTMSCounselPointComponents " & _
                   " (CounselPoint, CounselSubPoint, SubPointDescription) " & _
                   " VALUES (11, 4, 'Other feelings (p120 par1)') "
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0142"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0142_To_0005_9979_0143()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0143"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0143_To_0005_9979_0144()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " TrueFalse, " & _
             " Comment) " & _
             " VALUES ('TMS_InclSrvMtgItemNameInWorkshtPrt', " & _
                       "TRUE, " & _
                     " 'Initial value = TRUE ')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0144"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0144_To_0005_9979_0145()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 4, Comment = 'Initial value = 4 (weeks)' " & _
                  " WHERE FldName = 'TMSCloseAssignments_For_Bros_1and3' "
                  
    CMSDB.Execute "UPDATE tblConstants " & _
                  "SET NumVal = 8, Comment = 'Initial value = 8 (weeks)' " & _
                  " WHERE FldName = 'TMSCloseAssignments_For_Sisters_2and3' "
                         
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0145"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0145_To_0005_9979_0146()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " TrueFalse, " & _
             " Comment) " & _
             " VALUES ('TMS_PrintSQListWithWorksheet', " & _
                       "TRUE, " & _
                     " 'Initial value = TRUE ')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0146"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0146_To_0005_9979_0147()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " TrueFalse, " & _
             " Comment) " & _
             " VALUES ('EmailShellExecuteLineBreak', " & _
                       "TRUE, " & _
                     " 'Initial value = TRUE (Setting TRUE replaces vbCrLf with %0D%0A when sending text to an email app using ShellExecute. vbCrLf can be ignored.')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0147"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub DB_Upgrade_0005_9979_0147_To_0005_9979_0148()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset
                
    CreateField ErrorCode, "tblTMSPrintWorkSheet", "AssignmentDate2", "TEXT", "25"
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0148"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub DB_Upgrade_0005_9979_0148_To_0005_9979_0149()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer, rs As Recordset, sSQL As String, pos1 As Long, pos2 As Long
Dim s1 As String, s2 As String, s3 As String
                
    DestroyGlobalObjects
                    
    On Error Resume Next
    DropField ErrorCode, "tblTMSCounselPointList", "PageOfBeBook"
    On Error GoTo ErrorTrap

    CreateField ErrorCode, "tblTMSCounselPointList", "PageOfBeBook", "TEXT", "50"
    
    InstantiateGlobalObjects
                    
    sSQL = "SELECT CounselPoint, CounselDescription, PageOfBeBook " & _
                "FROM tblTMSCounselPointList "
    
    Set rs = CMSDB.OpenRecordset(sSQL, dbOpenDynaset)

    With rs
    Do Until .EOF Or .BOF
    
        s1 = rs!CounselDescription
        
        pos1 = InStr(1, s1, "(")
        
        If pos1 > 0 Then
            
            pos2 = InStr(pos1, s1, ")")
            
            s2 = Trim(Left(s1, pos1 - 1))
            s3 = Mid(s1, pos1 + 1, pos2 - 1 - pos1)
            
            .Edit
            !CounselDescription = s2
            !PageOfBeBook = s3
            .Update
            
            
        End If
        
        .MoveNext
        
    Loop
    End With
    
    Set rs = Nothing
                    
               
                
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0149"

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub DB_Upgrade_0005_9979_0149_To_0005_9979_0150()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    CMSDB.Execute "DELETE FROM tblConstants WHERE FldName = 'ServiceMeetingAddSeqNoToItemName'"
    CMSDB.Execute "DELETE FROM tblConstants WHERE FldName = 'ServiceMeetingDefaultItemLength'"

    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " TrueFalse, " & _
             " Comment) " & _
             " VALUES ('ServiceMeetingAddSeqNoToItemName', " & _
                       "TRUE, " & _
                     " 'Initial value = TRUE. Indicates whether to add item sequence number. eg Item 1')"
                     
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " NumVal, " & _
             " Comment) " & _
             " VALUES ('ServiceMeetingDefaultItemLength', " & _
                       "10, " & _
                     " 'Initial value = 10')"
                          
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0150"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0150_To_0005_9979_0151()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    CreateField ErrorCode, "tblMinReports", "NoTracts", "LONG"
    CMSDB.Execute "UPDATE tblMinReports SET NoTracts = 0"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0151"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0151_To_0005_9979_0152()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    CreateField ErrorCode, "tblAdvancedMinReporting", "NoTracts", "DOUBLE", , ""
    CreateField ErrorCode, "tblAdvancedMinReportingPrint", "AvgTracts", "DOUBLE"

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0152"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0152_To_0005_9979_0154()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    DestroyGlobalObjects True
    
    CreateIndex ErrorCode, "tblAccInOutTypes", "InOutID", "IX_InOutID", False, True, False
    CreateIndex ErrorCode, "tblAdvancedMinReporting", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblAdvancedMinReporting", "ActualMinDate", "IX_ActualMinDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblAdvancedMinReportingPrint", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblAuxPioDates", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblAuxPioDates", "StartDate", "IX_StartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblAuxPioDates", "EndDate", "IX_EndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblChildren", "Parent", "IX_Parent", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblChildren", "Child", "IX_Child", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblCleaningRota", "RotaDate", "IX_RotaDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblCongBibleStudyRota", "ConductorID", "IX_ConductorID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblCongBibleStudyRota", "ReaderID", "IX_ReaderID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblCongBibleStudyRota", "PrayerID", "IX_PrayerID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblConstants", "FldName", "IX_FldName", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblEvents", "SeqNum", "IX_SeqNum", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblEvents", "EventStartDate", "IX_EventStartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblGiftAidPayerActiveDates", "GiftAidNo", "IX_GiftAidNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblGiftAidPayers", "GiftAidNo", "IX_GiftAidNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblIDWeightings", "ID", "IX_ID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblInactivePubs", "MissingReportGroupID", "IX_MissingReportGroupID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblInactivePubs", "StartDate", "IX_StartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblInactivePubs", "EndDate", "IX_EndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblIrregularPubs", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblIrregularPubs", "IrregStartDate", "IX_IrregStartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblIrregularPubs", "IrregEndDate", "IX_IrregEndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "SocietyReportingMonth", "IX_SocietyReportingMonth", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "SocietyReportingYear", "IX_SocietyReportingYear", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "MinistryDoneInMonth", "IX_MinistryDoneInMonth", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "MinistryDoneInYear", "IX_MinistryDoneInYear", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "ActualMinPeriod", "IX_ActualMinPeriod", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMinReports", "SocietyReportingPeriod", "IX_SocietyReportingPeriod", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMissingReports", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMissingReports", "ServiceYear", "IX_ServiceYear", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMissingReports", "ServiceMonth", "IX_ServiceMonth", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMissingReports", "ActualMinDate", "IX_ActualMinDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblMissingReports", "MissingReportGroupID", "IX_MissingReportGroupID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPioHourCredit", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPioHourCredit", "MinDate", "IX_MinDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "MeetingDate", "IX_MeetingDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "SpeakerID", "IX_SpeakerID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "TalkNo", "IX_TalkNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "ChairmanID", "IX_ChairmanID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "WTReaderID", "IX_WTReaderID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "CongNoWhereMtgIs", "IX_CongNoWhereMtgIs", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublicMtgSchedule", "SpeakerID2", "IX_SpeakerID2", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublisherDates", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublisherDates", "StartDate", "IX_StartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublisherDates", "EndDate", "IX_EndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublisherDates", "StartReason", "IX_StartReason", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblPublisherDates", "EndReason", "IX_EndReason", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblRegPioDates", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblRegPioDates", "StartDate", "IX_StartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblRegPioDates", "EndDate", "IX_EndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblServiceMtgs", "MeetingDate", "IX_MeetingDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblServiceMtgs", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblSpecPioDates", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblSpecPioDates", "StartDate", "IX_StartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblSpecPioDates", "EndDate", "IX_EndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblBaptismDates", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblBaptismDates", "BaptismDate", "IX_BaptismDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskAndPerson", "TaskCategory", "IX_TaskCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskAndPerson", "TaskSubCategory", "IX_TaskSubCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskAndPerson", "Task", "IX_Task", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskAndPerson", "Person", "IX_Person", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "TaskCategory", "IX_TaskCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "TaskSubCategory", "IX_TaskSubCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "Task", "IX_Task", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "Person", "IX_Person", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "SuspendStartDate", "IX_SuspendStartDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTaskPersonSuspendDates", "SuspendEndDate", "IX_SuspendEndDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTasks", "TaskCategory", "IX_TaskCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTasks", "TaskSubCategory", "IX_TaskSubCategory", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTasks", "Task", "IX_Task", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTasks", "RequiresExemplaryBro", "IX_RequiresExemplaryBro", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTerritoryDNCs", "StreetSeqNum", "IX_StreetSeqNum", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTerritoryMapDates", "MapNo", "IX_MapNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTerritoryStreets", "MapNo", "IX_MapNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSItems", "ItemsSeqNum", "IX_ItemsSeqNum", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSItems", "AssignmentDate", "IX_AssignmentDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSItems", "TalkNo", "IX_TalkNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSItems", "TaskNo", "IX_TaskNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSItems", "TalkSeqNum", "IX_TalkSeqNum", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "ScheduleSeqNum", "IX_ScheduleSeqNum", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "AssignmentDate", "IX_AssignmentDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "TalkNo", "IX_TalkNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "SchoolNo", "IX_SchoolNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "Assistant1ID", "IX_Assistant1ID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "TalkCompleted", "IX_TalkCompleted", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "TalkDefaulted", "IX_TalkDefaulted", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchedule", "SlipPrinted", "IX_SlipPrinted", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchoolAndPerson", "PersonID", "IX_PersonID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchoolAndPerson", "SchoolNo", "IX_SchoolNo", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTMSSchoolAndPerson", "Task", "IX_Task", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "TranID", "IX_TranID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "TranCodeID", "IX_TranCodeID", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "TranDate", "IX_TranDate", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "FinancialYear", "IX_FinancialYear", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "FinancialMonth", "IX_FinancialMonth", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblTransactionDates", "FinancialQuarter", "IX_FinancialQuarter", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False
    CreateIndex ErrorCode, "tblNameAddress", "Active", "IX_Active", UniqueIX:=False, AllowNull:=True, CreatePrimary:=False

    InstantiateGlobalObjects True

    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0154"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub DB_Upgrade_0005_9979_0154_To_0005_9979_0155()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " AlphaVal, " & _
             " Comment) " & _
             " VALUES ('TMS_PrintSQListWithWorksheet_Type', " & _
                       "'PerStudent', " & _
                     " 'Initial value = PerStudent. Either PerStudent or OneSheet')"
                                               
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0155"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0155_To_0005_9979_0156()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CreateField ErrorCode, "tblTransactionTypes", "Suppressed", "YESNO"
    
                                               
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0156"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0156_To_0005_9979_0157()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CreateField ErrorCode, "tblTransactionSubTypes", "Suppressed", "YESNO"
    
                                               
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0157"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub
Private Sub DB_Upgrade_0005_9979_0157_To_0005_9979_0158()
On Error GoTo ErrorTrap
Dim ErrorCode As Integer
        
    CMSDB.Execute "INSERT INTO tblConstants " & _
             "(FldName, " & _
             " TrueFalse, " & _
             " Comment) " & _
             " VALUES ('TMS_FlagMissingSetting', " & _
                       "TRUE, " & _
                     " 'Initial value = TRUE')"
    
                                               
    '
    'Update DB version
    '
    GlobalParms.Save "CMS_Version", "AlphaVal", "0005.9979.0158"
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


