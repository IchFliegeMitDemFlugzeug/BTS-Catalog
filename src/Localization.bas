' Russian user-facing strings stored as Unicode code points.
Option Explicit

Public Function BTS_T(ByVal id As Long) As String
    ' Return a localized string without dependence on the VBA project codepage.
    Select Case id
        Case 1
            BTS_T = ChrW(1053) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1077) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077)
        Case 2
            BTS_T = ChrW(1055) & ChrW(1077) & ChrW(1088) & ChrW(1077) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1100) & " " & ChrW(1076) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1077) & ChrW(1081)
        Case 3
            BTS_T = ChrW(1064) & ChrW(1090) & ChrW(1091) & ChrW(1094) & ChrW(1077) & ChrW(1088) & ChrW(1099)
        Case 4
            BTS_T = ChrW(1050) & ChrW(1088) & ChrW(1099) & ChrW(1096) & ChrW(1082) & ChrW(1080)
        Case 5
            BTS_T = ChrW(1044) & ChrW(1059) & ChrW(1058)
        Case 6
            BTS_T = ChrW(1044) & ChrW(1058)
        Case 7
            BTS_T = ChrW(1057) & ChrW(1090) & ChrW(1072) & ChrW(1085) & ChrW(1076) & ChrW(1072) & ChrW(1088) & ChrW(1090) & ChrW(1085) & ChrW(1099) & ChrW(1077) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1103)
        Case 8
            BTS_T = ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079)
        Case 9
            BTS_T = ChrW(1057) & ChrW(1087) & ChrW(1088) & ChrW(1072) & ChrW(1074) & ChrW(1082) & ChrW(1072)
        Case 10
            BTS_T = ChrW(1054) & ChrW(1073) & ChrW(1086) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077)
        Case 11
            BTS_T = ChrW(1053) & ChrW(1072) & ChrW(1080) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077)
        Case 12
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083)
        Case 13
            BTS_T = ChrW(1050) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086)
        Case 14
            BTS_T = ChrW(1043) & ChrW(1088) & ChrW(1091) & ChrW(1087) & ChrW(1087) & ChrW(1072)
        Case 15
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & " " & "1"
        Case 16
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & " " & "2"
        Case 17
            BTS_T = ChrW(1044) & ChrW(1080) & ChrW(1072) & ChrW(1084) & ChrW(1077) & ChrW(1090) & ChrW(1088)
        Case 18
            BTS_T = ChrW(1059) & ChrW(1075) & ChrW(1086) & ChrW(1083)
        Case 19
            BTS_T = ChrW(1053) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088)
        Case 20
            BTS_T = ChrW(1057) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1074)
        Case 21
            BTS_T = ChrW(1056) & ChrW(1086) & ChrW(1083) & ChrW(1100)
        Case 22
            BTS_T = ChrW(1054) & ChrW(1089) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077)
        Case 23
            BTS_T = ChrW(1043) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1082) & ChrW(1072)
        Case 24
            BTS_T = ChrW(1055) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1077) & ChrW(1077)
        Case 25
            BTS_T = ChrW(1059) & ChrW(1076) & ChrW(1072) & ChrW(1083) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1074) & ChrW(1099) & ChrW(1073) & ChrW(1088) & ChrW(1072) & ChrW(1085) & ChrW(1085) & ChrW(1086) & ChrW(1077)
        Case 26
            BTS_T = ChrW(1044) & ChrW(1086) & ChrW(1073) & ChrW(1072) & ChrW(1074) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077)
        Case 27
            BTS_T = ChrW(1057) & ChrW(1086) & ChrW(1093) & ChrW(1088) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1080) & ChrW(1079) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1103)
        Case 28
            BTS_T = ChrW(1044) & ChrW(1086) & ChrW(1073) & ChrW(1072) & ChrW(1074) & ChrW(1080) & ChrW(1090) & ChrW(1100)
        Case 29
            BTS_T = ChrW(1054) & ChrW(1095) & ChrW(1080) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1079) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079)
        Case 30
            BTS_T = ChrW(1057) & ChrW(1092) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1080) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1090) & ChrW(1100) & " " & ChrW(1079) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079)
        Case 31
            BTS_T = ChrW(1064) & ChrW(1090) & ChrW(1091) & ChrW(1094) & ChrW(1077) & ChrW(1088) & " " & ChrW(1087) & ChrW(1088) & ChrW(1103) & ChrW(1084) & ChrW(1086) & ChrW(1081) & "," & " " & ChrW(1087) & ChrW(1086) & ChrW(1076) & " " & ChrW(1090) & ChrW(1088) & ChrW(1091) & ChrW(1073) & ChrW(1082) & ChrW(1091) & " " & ChrW(216)
        Case 32
            BTS_T = ChrW(1064) & ChrW(1090) & ChrW(1091) & ChrW(1094) & ChrW(1077) & ChrW(1088) & " " & ChrW(1091) & ChrW(1075) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1086) & ChrW(1081) & " "
        Case 33
            BTS_T = ChrW(176) & "," & " " & ChrW(1087) & ChrW(1086) & ChrW(1076) & " " & ChrW(1090) & ChrW(1088) & ChrW(1091) & ChrW(1073) & ChrW(1082) & ChrW(1091) & " " & ChrW(216)
        Case 34
            BTS_T = " " & ChrW(1084) & ChrW(1084)
        Case 35
            BTS_T = ChrW(1054) & ChrW(1073) & ChrW(1086) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077) & " " & ChrW(1091) & ChrW(1078) & ChrW(1077) & " " & ChrW(1079) & ChrW(1072) & ChrW(1088) & ChrW(1077) & ChrW(1075) & ChrW(1080) & ChrW(1089) & ChrW(1090) & ChrW(1088) & ChrW(1080) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1086) & "."
        Case 37
            BTS_T = ChrW(1048) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077) & " " & ChrW(1089) & ChrW(1086) & ChrW(1093) & ChrW(1088) & ChrW(1072) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1086) & "."
        Case 38
            BTS_T = ChrW(1044) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1080)
        Case 40
            BTS_T = ChrW(1050) & ChrW(1086) & ChrW(1076)
        Case 41
            BTS_T = ChrW(1048) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077)
        Case 42
            BTS_T = ChrW(1044) & ChrW(1083) & ChrW(1103) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1080) & " " & ChrW(1091) & ChrW(1082) & ChrW(1072) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1076) & ChrW(1074) & ChrW(1072) & " " & ChrW(1082) & ChrW(1086) & ChrW(1076) & ChrW(1072) & " " & ChrW(1084) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & ChrW(1086) & ChrW(1074) & "."
        Case 43
            BTS_T = ChrW(1055) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1077) & ChrW(1088) & ChrW(1100) & ChrW(1090) & ChrW(1077) & " " & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1087) & ChrW(1086) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1099) & " " & ChrW(1080) & " " & ChrW(1080) & ChrW(1093) & " " & ChrW(1082) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086) & "."
        Case 44
            BTS_T = ChrW(1054) & ChrW(1073) & ChrW(1085) & ChrW(1072) & ChrW(1088) & ChrW(1091) & ChrW(1078) & ChrW(1077) & ChrW(1085) & ChrW(1086) & " " & ChrW(1094) & ChrW(1080) & ChrW(1082) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1082) & ChrW(1086) & ChrW(1077) & " " & ChrW(1074) & ChrW(1082) & ChrW(1083) & ChrW(1102) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1082) & "."
        Case 45
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & " " & ChrW(1086) & ChrW(1089) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1103) & " " & ChrW(1080) & ChrW(1083) & ChrW(1080) & " " & ChrW(1075) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1082) & ChrW(1080) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1089) & ChrW(1086) & ChrW(1074) & ChrW(1087) & ChrW(1072) & ChrW(1076) & ChrW(1072) & ChrW(1077) & ChrW(1090) & " " & ChrW(1089) & " " & ChrW(1086) & ChrW(1073) & ChrW(1086) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1077) & ChrW(1084) & "."
        Case 46
            BTS_T = ChrW(1044) & ChrW(1083) & ChrW(1103) & " " & ChrW(1096) & ChrW(1090) & ChrW(1091) & ChrW(1094) & ChrW(1077) & ChrW(1088) & ChrW(1072) & " " & ChrW(1085) & ChrW(1091) & ChrW(1078) & ChrW(1085) & ChrW(1086) & " " & ChrW(1086) & ChrW(1076) & ChrW(1085) & ChrW(1086) & " " & ChrW(1086) & ChrW(1089) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077) & " " & ChrW(1080) & " " & ChrW(1086) & ChrW(1076) & ChrW(1085) & ChrW(1072) & " " & ChrW(1075) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1082) & ChrW(1072) & "."
        Case 47
            BTS_T = ChrW(1059) & ChrW(1082) & ChrW(1072) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1085) & ChrW(1072) & ChrW(1080) & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1077) & " " & ChrW(1076) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1080) & "."
        Case 48
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1088) & ChrW(1082) & ChrW(1072) & " " & ChrW(1084) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & ChrW(1072) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1089) & ChrW(1086) & ChrW(1086) & ChrW(1090) & ChrW(1074) & ChrW(1077) & ChrW(1090) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1091) & ChrW(1077) & ChrW(1090) & " " & ChrW(1082) & ChrW(1086) & ChrW(1076) & ChrW(1091) & "."
        Case 49
            BTS_T = ChrW(1048) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077) & " " & ChrW(1091) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1072) & ChrW(1085) & ChrW(1085) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1074) & ChrW(1080) & ChrW(1076) & ChrW(1072) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1085) & ChrW(1072) & ChrW(1081) & ChrW(1076) & ChrW(1077) & ChrW(1085) & ChrW(1086) & "."
        Case 50
            BTS_T = ChrW(1053) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1086) & ChrW(1081) & " " & ChrW(1077) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1080) & ChrW(1094) & ChrW(1077) & "?" & " " & ChrW(1044) & ChrW(1072) & " " & "-" & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1072) & "," & " " & ChrW(1053) & ChrW(1077) & ChrW(1090) & " " & "-" & " " & ChrW(1076) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100) & "."
        Case 51
            BTS_T = ChrW(1057) & ChrW(1074) & ChrW(1086) & ChrW(1073) & ChrW(1086) & ChrW(1076) & ChrW(1085) & ChrW(1086) & ChrW(1075) & ChrW(1086) & " " & ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088) & ChrW(1072) & " " & ChrW(1085) & ChrW(1077) & ChrW(1090) & "."
        Case 52
            BTS_T = ChrW(1059) & ChrW(1076) & ChrW(1072) & ChrW(1083) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077) & "?" & " " & ChrW(1047) & ChrW(1072) & ChrW(1090) & ChrW(1088) & ChrW(1086) & ChrW(1085) & ChrW(1091) & ChrW(1090) & ChrW(1099) & ChrW(1077) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1080) & ":"
        Case 53
            BTS_T = ChrW(1048) & ChrW(1089) & ChrW(1082) & ChrW(1083) & ChrW(1102) & ChrW(1095) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1101) & ChrW(1090) & ChrW(1091) & " " & ChrW(1074) & ChrW(1077) & ChrW(1090) & ChrW(1074) & ChrW(1100) & " " & ChrW(1090) & ChrW(1086) & ChrW(1083) & ChrW(1100) & ChrW(1082) & ChrW(1086) & " " & ChrW(1080) & ChrW(1079) & " " & ChrW(1079) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & ChrW(1072) & "?"
        Case 54
            BTS_T = ChrW(1048) & ChrW(1089) & ChrW(1082) & ChrW(1083) & ChrW(1102) & ChrW(1095) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1087) & ChrW(1086) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1090) & " " & ChrW(1080) & ChrW(1079) & " " & ChrW(1089) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1074) & ChrW(1072) & ":" & " "
        Case 55
            BTS_T = ChrW(1057) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1072) & ChrW(1103) & " " & ChrW(1077) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1080) & ChrW(1094) & ChrW(1072) & ":" & " "
        Case 56
            BTS_T = ChrW(1050) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086) & " " & ChrW(1076) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1085) & ChrW(1086) & " " & ChrW(1073) & ChrW(1099) & ChrW(1090) & ChrW(1100) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & ChrW(1083) & ChrW(1100) & ChrW(1085) & ChrW(1099) & ChrW(1084) & " " & ChrW(1094) & ChrW(1077) & ChrW(1083) & ChrW(1099) & ChrW(1084) & "."
        Case 57
            BTS_T = ChrW(1054) & ChrW(1095) & ChrW(1080) & ChrW(1089) & ChrW(1090) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1079) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & "?"
        Case 58
            BTS_T = ChrW(1057) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1072) & ChrW(1083) & ChrW(1072) & " " & ChrW(1089) & ChrW(1086) & ChrW(1093) & ChrW(1088) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1086) & ChrW(1089) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1085) & ChrW(1091) & ChrW(1102) & " " & ChrW(1082) & ChrW(1085) & ChrW(1080) & ChrW(1075) & ChrW(1091) & "."
        Case 59
            BTS_T = ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & " " & ChrW(1087) & ChrW(1091) & ChrW(1089) & ChrW(1090) & "."
        Case 60
            BTS_T = ChrW(1056) & ChrW(1072) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083)
        Case 61
            BTS_T = ChrW(1044) & "1" & "6" & ChrW(1058)
        Case 62
            BTS_T = ChrW(1041) & ChrW(1088) & ChrW(1040) & ChrW(1046) & "9" & "-" & "4"
        Case 63
            BTS_T = ChrW(1051) & ChrW(1057) & "5" & "9" & "-" & "1"
        Case 64
            BTS_T = "1" & "2" & ChrW(1061) & "1" & "8" & ChrW(1053) & "1" & "0" & ChrW(1058)
        Case 65
            BTS_T = ChrW(1057) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1072) & ChrW(1103) & " " & ChrW(1077) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1080) & ChrW(1094) & ChrW(1072)
        Case 66
            BTS_T = ChrW(1044) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100)
        Case 67
            BTS_T = ChrW(1047) & ChrW(1072) & ChrW(1075) & ChrW(1088) & ChrW(1091) & ChrW(1079) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077)
        Case 68
            BTS_T = ChrW(1055) & ChrW(1086) & ChrW(1076) & ChrW(1086) & ChrW(1073) & ChrW(1088) & ChrW(1072) & ChrW(1090) & ChrW(1100) & " " & ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088)
        Case 69
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1088) & ChrW(1082) & ChrW(1072) & " " & ChrW(1084) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & ChrW(1072)
        Case 70
            BTS_T = ChrW(1057) & ChrW(1090) & ChrW(1072) & ChrW(1085) & ChrW(1076) & ChrW(1072) & ChrW(1088) & ChrW(1090) & ChrW(1085) & ChrW(1099) & ChrW(1077) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1103)
        Case 71
            BTS_T = ChrW(1047) & ChrW(1072) & ChrW(1082) & ChrW(1072) & ChrW(1079) & " " & ChrW(1089) & ChrW(1092) & ChrW(1086) & ChrW(1088) & ChrW(1084) & ChrW(1080) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ":"
        Case 72
            BTS_T = ChrW(1042) & ChrW(1080) & ChrW(1076) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1103)
        Case 73
            BTS_T = ChrW(1044) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1100)
        Case 74
            BTS_T = ChrW(1057) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1072) & ChrW(1103) & " " & ChrW(1077) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1080) & ChrW(1094) & ChrW(1072)
        Case 75
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & " " & ChrW(1086) & ChrW(1089) & ChrW(1085) & ChrW(1086) & ChrW(1074) & ChrW(1072) & ChrW(1085) & ChrW(1080) & ChrW(1103)
        Case 76
            BTS_T = ChrW(1052) & ChrW(1072) & ChrW(1090) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1072) & ChrW(1083) & " " & ChrW(1075) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1074) & ChrW(1082) & ChrW(1080)
        Case 77
            BTS_T = ChrW(1055) & ChrW(1088) & ChrW(1086) & ChrW(1074) & ChrW(1077) & ChrW(1088) & ChrW(1100) & ChrW(1090) & ChrW(1077) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1103) & " " & ChrW(1086) & ChrW(1073) & ChrW(1086) & ChrW(1079) & ChrW(1085) & ChrW(1072) & ChrW(1095) & ChrW(1077) & ChrW(1085) & ChrW(1080) & ChrW(1103) & "." & " " & ChrW(1053) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1080) & " " & ChrW(1076) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1077) & ChrW(1085) & " " & ChrW(1086) & ChrW(1082) & ChrW(1072) & ChrW(1085) & ChrW(1095) & ChrW(1080) & ChrW(1074) & ChrW(1072) & ChrW(1090) & ChrW(1100) & ChrW(1089) & ChrW(1103) & " " & ChrW(1085) & ChrW(1072) & " " & "0" & "," & " " & ChrW(1085) & ChrW(1086) & ChrW(1084) & ChrW(1077) & ChrW(1088) & " " & ChrW(1076) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1080) & " " & ChrW(1085) & ChrW(1072) & " " & "1" & "-" & "9" & "."
        Case 78
            BTS_T = ChrW(1050) & ChrW(1086) & ChrW(1084) & ChrW(1087) & ChrW(1086) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1090)
        Case 79
            BTS_T = ChrW(1044) & ChrW(1077) & ChrW(1090) & ChrW(1072) & ChrW(1083) & ChrW(1080) & " " & ChrW(1080) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1080)
        Case 80
            BTS_T = ChrW(8470)
        Case 81
            BTS_T = ChrW(1059) & ChrW(1076) & ChrW(1072) & ChrW(1083) & ChrW(1080) & ChrW(1090) & ChrW(1100)
        Case 82
            BTS_T = ChrW(1044) & ChrW(1086) & ChrW(1073) & ChrW(1072) & ChrW(1074) & ChrW(1080) & ChrW(1090) & ChrW(1100) & " " & ChrW(1074) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1082) & ChrW(1091)
        Case 83
            BTS_T = ChrW(1042) & " " & ChrW(1089) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1074) & ChrW(1077) & " " & ChrW(1089) & ChrW(1073) & ChrW(1086) & ChrW(1088) & ChrW(1086) & ChrW(1095) & ChrW(1085) & ChrW(1086) & ChrW(1081) & " " & ChrW(1077) & ChrW(1076) & ChrW(1080) & ChrW(1085) & ChrW(1080) & ChrW(1094) & ChrW(1099) & " " & ChrW(1076) & ChrW(1086) & ChrW(1083) & ChrW(1078) & ChrW(1085) & ChrW(1086) & " " & ChrW(1073) & ChrW(1099) & ChrW(1090) & ChrW(1100) & " " & ChrW(1085) & ChrW(1077) & " " & ChrW(1084) & ChrW(1077) & ChrW(1085) & ChrW(1077) & ChrW(1077) & " " & ChrW(1076) & ChrW(1074) & ChrW(1091) & ChrW(1093) & " " & ChrW(1088) & ChrW(1072) & ChrW(1079) & ChrW(1085) & ChrW(1099) & ChrW(1093) & " " & ChrW(1082) & ChrW(1086) & ChrW(1084) & ChrW(1087) & ChrW(1086) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1090) & ChrW(1086) & ChrW(1074) & "."
        Case 86
            BTS_T = ChrW(1042) & ChrW(1099) & ChrW(1073) & ChrW(1077) & ChrW(1088) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1089) & ChrW(1091) & ChrW(1097) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1091) & ChrW(1102) & ChrW(1097) & ChrW(1077) & ChrW(1077) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1077) & " " & ChrW(1080) & " " & ChrW(1091) & ChrW(1082) & ChrW(1072) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & " " & ChrW(1094) & ChrW(1077) & ChrW(1083) & ChrW(1086) & ChrW(1077) & " " & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1086) & ChrW(1078) & ChrW(1080) & ChrW(1090) & ChrW(1077) & ChrW(1083) & ChrW(1100) & ChrW(1085) & ChrW(1086) & ChrW(1077) & " " & ChrW(1082) & ChrW(1086) & ChrW(1083) & ChrW(1080) & ChrW(1095) & ChrW(1077) & ChrW(1089) & ChrW(1090) & ChrW(1074) & ChrW(1086) & "."
        Case 87
            BTS_T = ChrW(1058) & ChrW(1072) & ChrW(1073) & ChrW(1083) & ChrW(1080) & ChrW(1094) & ChrW(1072) & " " & ChrW(1089) & ChrW(1086) & ChrW(1089) & ChrW(1090) & ChrW(1072) & ChrW(1074) & ChrW(1072) & " " & ChrW(1079) & ChrW(1072) & ChrW(1087) & ChrW(1086) & ChrW(1083) & ChrW(1085) & ChrW(1077) & ChrW(1085) & ChrW(1072) & "."
        Case 89
            BTS_T = ChrW(1043) & ChrW(1054) & ChrW(1057) & ChrW(1058) & " " & ChrW(1056)
        Case 90
            BTS_T = ChrW(1043) & ChrW(1054) & ChrW(1057) & ChrW(1058)
        Case 91
            BTS_T = ChrW(1057) & ChrW(1058) & ChrW(1054)
        Case 92
            BTS_T = ChrW(1054) & ChrW(1057) & ChrW(1058)
        Case 93
            BTS_T = ChrW(1042) & ChrW(1080) & ChrW(1085) & ChrW(1090)
        Case 94
            BTS_T = ChrW(1043) & ChrW(1072) & ChrW(1081) & ChrW(1082) & ChrW(1072)
        Case 95
            BTS_T = ChrW(1064) & ChrW(1072) & ChrW(1081) & ChrW(1073) & ChrW(1072)
        Case 96
            BTS_T = ChrW(1041) & ChrW(1086) & ChrW(1083) & ChrW(1090)
        Case 97
            BTS_T = ChrW(1064) & ChrW(1087) & ChrW(1080) & ChrW(1083) & ChrW(1100) & ChrW(1082) & ChrW(1072)
        Case 98
            BTS_T = ChrW(1050) & ChrW(1088) & ChrW(1077) & ChrW(1087) & ChrW(1077) & ChrW(1078) & ChrW(1085) & ChrW(1099) & ChrW(1077) & " " & ChrW(1080) & ChrW(1079) & ChrW(1076) & ChrW(1077) & ChrW(1083) & ChrW(1080) & ChrW(1103)
        Case Else
            BTS_T = ""
    End Select
End Function
