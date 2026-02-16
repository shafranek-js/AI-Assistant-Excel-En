Attribute VB_Name = "CommandExecutor"
Option Explicit

Public Function ExecuteCommands(commands As String) As String
    On Error GoTo ErrorHandler
    
    Dim lines() As String
    Dim i As Long
    Dim cmd As String
    Dim executedCount As Long
    Dim rejectedCount As Long
    Dim failedCount As Long
    Dim result As String
    Dim validationError As String
    Dim details As String
    Dim detailCount As Long
    Dim maxDetails As Long
    
    If Len(commands) = 0 Then
        ExecuteCommands = ""
        Exit Function
    End If
    
    lines = Split(commands, vbLf)
    executedCount = 0
    rejectedCount = 0
    failedCount = 0
    details = ""
    detailCount = 0
    maxDetails = 5
    
    For i = 0 To UBound(lines)
        cmd = Trim(Replace(lines(i), vbCr, ""))
        If Len(cmd) > 0 Then
            validationError = ""
            If ValidateCommandStrict(cmd, validationError) Then
                If ExecuteSingleCommand(cmd) Then
                    executedCount = executedCount + 1
                Else
                    failedCount = failedCount + 1
                    If detailCount < maxDetails Then
                        details = details & vbCrLf & "- Runtime failure: " & cmd
                        detailCount = detailCount + 1
                    End If
                End If
            Else
                rejectedCount = rejectedCount + 1
                If detailCount < maxDetails Then
                    details = details & vbCrLf & "- Rejected (" & validationError & "): " & cmd
                    detailCount = detailCount + 1
                End If
            End If
        End If
    Next i
    
    result = "[Commands executed: " & executedCount & ", rejected: " & rejectedCount & ", failed: " & failedCount & "]"
    If Len(details) > 0 Then
        result = result & vbCrLf & "Details:" & details
    End If
    ExecuteCommands = result
    Exit Function
    
ErrorHandler:
    ExecuteCommands = "Runtime error: " & Err.Description
End Function

Private Function ParseColor(colorStr As String) As Long
    On Error Resume Next
    
    Dim c As String
    Dim rgbParts() As String
    
    c = UCase(Trim(colorStr))
    Debug.Print "ParseColor: input=[" & colorStr & "] upper=[" & c & "]"
    
    ' Checking the RGB format
    If Len(c) >= 4 And Left(c, 4) = "RGB:" Then
        rgbParts = Split(Mid(c, 5), ",")
        If UBound(rgbParts) >= 2 Then
            ParseColor = RGB(CLng(Trim(rgbParts(0))), CLng(Trim(rgbParts(1))), CLng(Trim(rgbParts(2))))
            Debug.Print "ParseColor: RGB result=" & ParseColor
            Exit Function
        End If
    End If
    
    ' Predefined colors
    Select Case c
        Case "RED": ParseColor = RGB(255, 0, 0)
        Case "GREEN": ParseColor = RGB(0, 128, 0)
        Case "BLUE": ParseColor = RGB(0, 0, 255)
        Case "YELLOW": ParseColor = RGB(255, 255, 0)
        Case "ORANGE": ParseColor = RGB(255, 165, 0)
        Case "PURPLE": ParseColor = RGB(128, 0, 128)
        Case "PINK": ParseColor = RGB(255, 192, 203)
        Case "CYAN": ParseColor = RGB(0, 255, 255)
        Case "WHITE": ParseColor = RGB(255, 255, 255)
        Case "BLACK": ParseColor = RGB(0, 0, 0)
        Case "GRAY", "GREY": ParseColor = RGB(128, 128, 128)
        Case "LIGHTGRAY", "LIGHTGREY": ParseColor = RGB(192, 192, 192)
        Case "DARKGRAY", "DARKGREY": ParseColor = RGB(64, 64, 64)
        Case "BROWN": ParseColor = RGB(139, 69, 19)
        Case "LIME": ParseColor = RGB(0, 255, 0)
        Case "NAVY": ParseColor = RGB(0, 0, 128)
        Case "TEAL": ParseColor = RGB(0, 128, 128)
        Case "MAROON": ParseColor = RGB(128, 0, 0)
        Case "OLIVE": ParseColor = RGB(128, 128, 0)
        Case "GOLD": ParseColor = RGB(255, 215, 0)
        Case "SILVER": ParseColor = RGB(192, 192, 192)
        Case Else: ParseColor = RGB(0, 0, 0) ' Default black
    End Select
    
    Debug.Print "ParseColor: result=" & ParseColor
    If Err.Number <> 0 Then
        Debug.Print "ParseColor: ERROR " & Err.Number & " - " & Err.Description
        Err.Clear
    End If
End Function

Private Function LocalizeFormula(formula As String) As String
    Dim result As String
    result = formula
    
    ' ===MATHEMATICAL ===
    result = ReplaceFunc(result, "ABS", "ABS")
    result = ReplaceFunc(result, "ACOS", "ACOS")
    result = ReplaceFunc(result, "ACOSH", "ACOSH")
    result = ReplaceFunc(result, "ACOT", "ACOT")
    result = ReplaceFunc(result, "ACOTH", "ACOTH")
    result = ReplaceFunc(result, "AGGREGATE", "UNIT")
    result = ReplaceFunc(result, "ARABIC", "ARAB")
    result = ReplaceFunc(result, "ASIN", "ASIN")
    result = ReplaceFunc(result, "ASINH", "ASINH")
    result = ReplaceFunc(result, "ATAN", "ATAN")
    result = ReplaceFunc(result, "ATAN2", "ATAN2")
    result = ReplaceFunc(result, "ATANH", "ATANH")
    result = ReplaceFunc(result, "BASE", "BASE")
    result = ReplaceFunc(result, "CEILING.MATH", "OVERTOP.MAT")
    result = ReplaceFunc(result, "CEILING.PRECISE", "OKRUP.PRECISION")
    result = ReplaceFunc(result, "CEILING", "OKRVVERH")
    result = ReplaceFunc(result, "COMBIN", "NUMBERCOMB")
    result = ReplaceFunc(result, "COMBINA", "NUMBERCOMBA")
    result = ReplaceFunc(result, "COS", "COS")
    result = ReplaceFunc(result, "COSH", "COSH")
    result = ReplaceFunc(result, "COT", "COT")
    result = ReplaceFunc(result, "COTH", "COTH")
    result = ReplaceFunc(result, "CSC", "CSC")
    result = ReplaceFunc(result, "CSCH", "CSCH")
    result = ReplaceFunc(result, "DECIMAL", "DES")
    result = ReplaceFunc(result, "DEGREES", "DEGREES")
    result = ReplaceFunc(result, "EVEN", "EVEN")
    result = ReplaceFunc(result, "EXP", "EXP")
    result = ReplaceFunc(result, "FACT", "FACT")
    result = ReplaceFunc(result, "FACTDOUBLE", "DVFACTR")
    result = ReplaceFunc(result, "FLOOR.MATH", "OKRVNIZ.MAT")
    result = ReplaceFunc(result, "FLOOR.PRECISE", "OKRV.PRECISION")
    result = ReplaceFunc(result, "FLOOR", "OKRVNIZ")
    result = ReplaceFunc(result, "GCD", "GCD")
    result = ReplaceFunc(result, "INT", "WHOLE")
    result = ReplaceFunc(result, "ISO.CEILING", "ISO.OVERUP")
    result = ReplaceFunc(result, "LCM", "NOC")
    result = ReplaceFunc(result, "LN", "LN")
    result = ReplaceFunc(result, "LOG10", "LOG10")
    result = ReplaceFunc(result, "LOG", "LOG")
    result = ReplaceFunc(result, "MDETERM", "MOPRED")
    result = ReplaceFunc(result, "MINVERSE", "MOBR")
    result = ReplaceFunc(result, "MMULT", "MUMNIFE")
    result = ReplaceFunc(result, "MOD", "OSTAT")
    result = ReplaceFunc(result, "MROUND", "ROUND")
    result = ReplaceFunc(result, "MULTINOMIAL", "MULTINOM")
    result = ReplaceFunc(result, "MUNIT", "MEDIN")
    result = ReplaceFunc(result, "ODD", "ODD")
    result = ReplaceFunc(result, "PI", "PI")
    result = ReplaceFunc(result, "POWER", "DEGREE")
    result = ReplaceFunc(result, "PRODUCT", "PRODUCT")
    result = ReplaceFunc(result, "QUOTIENT", "PRIVATE")
    result = ReplaceFunc(result, "RADIANS", "RADIANS")
    result = ReplaceFunc(result, "RANDBETWEEN", "CASE BETWEEN")
    result = ReplaceFunc(result, "RAND", "RAND")
    result = ReplaceFunc(result, "ROMAN", "ROMAN")
    result = ReplaceFunc(result, "ROUNDDOWN", "ROUND BOTTOM")
    result = ReplaceFunc(result, "ROUNDUP", "ROUNDUP")
    result = ReplaceFunc(result, "ROUND", "ROUND")
    result = ReplaceFunc(result, "SEC", "SEC")
    result = ReplaceFunc(result, "SECH", "SECH")
    result = ReplaceFunc(result, "SERIESSUM", "SERIES.SUM")
    result = ReplaceFunc(result, "SIGN", "SIGN")
    result = ReplaceFunc(result, "SIN", "SIN")
    result = ReplaceFunc(result, "SINH", "SINH")
    result = ReplaceFunc(result, "SQRT", "ROOT")
    result = ReplaceFunc(result, "SQRTPI", "KORE|NY.PI")
    result = ReplaceFunc(result, "SUBTOTAL", "INTERMEDIATE.RESULTS")
    result = ReplaceFunc(result, "SUMIFS", "SUMIFS")
    result = ReplaceFunc(result, "SUMIF", "SUMIF")
    result = ReplaceFunc(result, "SUMPRODUCT", "SUMPRODUCT")
    result = ReplaceFunc(result, "SUMSQ", "SUMMKV")
    result = ReplaceFunc(result, "SUMX2MY2", "SUMMDISC")
    result = ReplaceFunc(result, "SUMX2PY2", "SUMMSUMMKV")
    result = ReplaceFunc(result, "SUMXMY2", "SUM DIFFERENCE")
    result = ReplaceFunc(result, "SUM", "SUM")
    result = ReplaceFunc(result, "TAN", "TAN")
    result = ReplaceFunc(result, "TANH", "TANH")
    result = ReplaceFunc(result, "TRUNC", "OTBR")
    
    ' === LOGICAL ===
    result = ReplaceFunc(result, "AND", "AND")
    result = ReplaceFunc(result, "FALSE", "LIE")
    result = ReplaceFunc(result, "IFERROR", "IFERROR")
    result = ReplaceFunc(result, "IFNA", "ESND")
    result = ReplaceFunc(result, "IFS", "CONDITIONS")
    result = ReplaceFunc(result, "IF", "IF")
    result = ReplaceFunc(result, "NOT", "NOT")
    result = ReplaceFunc(result, "OR", "OR")
    result = ReplaceFunc(result, "SWITCH", "SWITCH")
    result = ReplaceFunc(result, "TRUE", "TRUE")
    result = ReplaceFunc(result, "XOR", "EXCLUDED")
    
    ' === TEXT ===
    result = ReplaceFunc(result, "ASC", "ASC")
    result = ReplaceFunc(result, "BAHTTEXT", "BATT.TEXT")
    result = ReplaceFunc(result, "CHAR", "SYMBOL")
    result = ReplaceFunc(result, "CLEAN", "PECHSIMV")
    result = ReplaceFunc(result, "CODE", "CODSIM")
    result = ReplaceFunc(result, "CONCATENATE", "CONNECT")
    result = ReplaceFunc(result, "CONCAT", "SCENER")
    result = ReplaceFunc(result, "DOLLAR", "RUBLE")
    result = ReplaceFunc(result, "EXACT", "COINCIDENCE")
    result = ReplaceFunc(result, "FIND", "FIND")
    result = ReplaceFunc(result, "FIXED", "FIXED")
    result = ReplaceFunc(result, "LEFT", "LEVSIMV")
    result = ReplaceFunc(result, "LEN", "DLST")
    result = ReplaceFunc(result, "LOWER", "LOWER")
    result = ReplaceFunc(result, "MID", "PSTR")
    result = ReplaceFunc(result, "NUMBERVALUE", "VALUE")
    result = ReplaceFunc(result, "PHONETIC", "PHONETIC")
    result = ReplaceFunc(result, "PROPER", "PROPNACH")
    result = ReplaceFunc(result, "REPLACE", "REPLACE")
    result = ReplaceFunc(result, "REPT", "REPEAT")
    result = ReplaceFunc(result, "RIGHT", "RIGHT")
    result = ReplaceFunc(result, "SEARCH", "SEARCH")
    result = ReplaceFunc(result, "SUBSTITUTE", "SUBSTITUTE")
    result = ReplaceFunc(result, "TEXTJOIN", "COMBINE")
    result = ReplaceFunc(result, "TEXT", "TEXT")
    result = ReplaceFunc(result, "TRIM", "SPACE")
    result = ReplaceFunc(result, "UNICHAR", "UNISIM")
    result = ReplaceFunc(result, "UNICODE", "UNICODE")
    result = ReplaceFunc(result, "UPPER", "CAPITAL")
    result = ReplaceFunc(result, "VALUE", "SIGNIFICANT")
    
    ' === DATE AND TIME ===
    result = ReplaceFunc(result, "DATE", "DATE")
    result = ReplaceFunc(result, "DATEDIF", "RAZNDAT")
    result = ReplaceFunc(result, "DATEVALUE", "DATEVALUE")
    result = ReplaceFunc(result, "DAY", "DAY")
    result = ReplaceFunc(result, "DAYS360", "DAYS360")
    result = ReplaceFunc(result, "DAYS", "DAYS")
    result = ReplaceFunc(result, "EDATE", "DATAMES")
    result = ReplaceFunc(result, "EOMONTH", "EON-MONTH")
    result = ReplaceFunc(result, "HOUR", "HOUR")
    result = ReplaceFunc(result, "ISOWEEKNUM", "WEEK NUMBER.ISO")
    result = ReplaceFunc(result, "MINUTE", "MINUTES")
    result = ReplaceFunc(result, "MONTH", "MONTH")
    result = ReplaceFunc(result, "NETWORKDAYS.INTL", "NETWORKDAYS.INTL")
    result = ReplaceFunc(result, "NETWORKDAYS", "NETWORKDAYS")
    result = ReplaceFunc(result, "NOW", "TDATE")
    result = ReplaceFunc(result, "SECOND", "SECONDS")
    result = ReplaceFunc(result, "TIMEVALUE", "TIMEVALUE")
    result = ReplaceFunc(result, "TIME", "TIME")
    result = ReplaceFunc(result, "TODAY", "TODAY")
    result = ReplaceFunc(result, "WEEKDAY", "WEEKDAY")
    result = ReplaceFunc(result, "WEEKNUM", "WEEK NUMBER")
    result = ReplaceFunc(result, "WORKDAY.INTL", "WORKDAY INTERNATIONAL")
    result = ReplaceFunc(result, "WORKDAY", "WORKDAY")
    result = ReplaceFunc(result, "YEARFRAC", "PERCENTAGE OF THE YEAR")
    result = ReplaceFunc(result, "YEAR", "YEAR")
    
    ' === LINKS AND SEARCH ===
    result = ReplaceFunc(result, "ADDRESS", "ADDRESS")
    result = ReplaceFunc(result, "AREAS", "AREAS")
    result = ReplaceFunc(result, "CHOOSE", "CHOICE")
    result = ReplaceFunc(result, "COLUMNS", "NUMBERCOLUMN")
    result = ReplaceFunc(result, "COLUMN", "COLUMN")
    result = ReplaceFunc(result, "FORMULATEXT", "F.TEXT")
    result = ReplaceFunc(result, "GETPIVOTDATA", "GET.PICTTABLE.DATA")
    result = ReplaceFunc(result, "HLOOKUP", "GPR")
    result = ReplaceFunc(result, "HYPERLINK", "HYPERLINK")
    result = ReplaceFunc(result, "INDEX", "INDEX")
    result = ReplaceFunc(result, "INDIRECT", "INDIRECT")
    result = ReplaceFunc(result, "LOOKUP", "VIEW")
    result = ReplaceFunc(result, "MATCH", "SEARCH")
    result = ReplaceFunc(result, "OFFSET", "OFFSET")
    result = ReplaceFunc(result, "ROWS", "LINE")
    result = ReplaceFunc(result, "ROW", "LINE")
    result = ReplaceFunc(result, "RTD", "DRV")
    result = ReplaceFunc(result, "TRANSPOSE", "TRANSSP")
    result = ReplaceFunc(result, "VLOOKUP", "VLOOKUP")
    result = ReplaceFunc(result, "XLOOKUP", "VIEWX")
    result = ReplaceFunc(result, "XMATCH", "MATCHX")
    
    ' === STATISTICAL ===
    result = ReplaceFunc(result, "AVEDEV", "SROTCL")
    result = ReplaceFunc(result, "AVERAGEIFS", "AVERAGEIFS")
    result = ReplaceFunc(result, "AVERAGEIF", "AVERAGEIF")
    result = ReplaceFunc(result, "AVERAGEA", "AVERAGE")
    result = ReplaceFunc(result, "AVERAGE", "AVERAGE")
    result = ReplaceFunc(result, "BETA.DIST", "BETA.DIST")
    result = ReplaceFunc(result, "BETA.INV", "BETA.OBR")
    result = ReplaceFunc(result, "BINOM.DIST.RANGE", "BINOM.DIST.RANGE")
    result = ReplaceFunc(result, "BINOM.DIST", "BINOM.DIST")
    result = ReplaceFunc(result, "BINOM.INV", "BINOM.OBR")
    result = ReplaceFunc(result, "CHISQ.DIST.RT", "CH2.DIST.PH")
    result = ReplaceFunc(result, "CHISQ.DIST", "CH2.DIST")
    result = ReplaceFunc(result, "CHISQ.INV.RT", "CH2.OBR.PH")
    result = ReplaceFunc(result, "CHISQ.INV", "CH2.OBR")
    result = ReplaceFunc(result, "CHISQ.TEST", "CHI2.TEST")
    result = ReplaceFunc(result, "CONFIDENCE.NORM", "TRUST.NORM")
    result = ReplaceFunc(result, "CONFIDENCE.T", "TRUSTEE.STUDENT")
    result = ReplaceFunc(result, "CORREL", "CORREL")
    result = ReplaceFunc(result, "COUNTA", "COUNTING")
    result = ReplaceFunc(result, "COUNTBLANK", "COUNT VOIDS")
    result = ReplaceFunc(result, "COUNTIFS", "COUNTIFS")
    result = ReplaceFunc(result, "COUNTIF", "COUNTIF")
    result = ReplaceFunc(result, "COUNT", "CHECK")
    result = ReplaceFunc(result, "COVARIANCE.P", "COVARIANCE.G")
    result = ReplaceFunc(result, "COVARIANCE.S", "COVARIANCE.B")
    result = ReplaceFunc(result, "DEVSQ", "QUADROTCL")
    result = ReplaceFunc(result, "EXPON.DIST", "EXP.DIST.")
    result = ReplaceFunc(result, "F.DIST.RT", "F.DIST.PH")
    result = ReplaceFunc(result, "F.DIST", "F.DIST")
    result = ReplaceFunc(result, "F.INV.RT", "F.REV.PH")
    result = ReplaceFunc(result, "F.INV", "F.OBR")
    result = ReplaceFunc(result, "FISHER", "FISCHER")
    result = ReplaceFunc(result, "FISHERINV", "FISHEROBR")
    result = ReplaceFunc(result, "FORECAST.ETS.CONFINT", "FORECAST.ETS.DEVINTERVAL")
    result = ReplaceFunc(result, "FORECAST.ETS.SEASONALITY", "FORECAST.ETS.SEASONALITY")
    result = ReplaceFunc(result, "FORECAST.ETS.STAT", "FORECAST.ETS.STAT")
    result = ReplaceFunc(result, "FORECAST.ETS", "FORECAST.ETS")
    result = ReplaceFunc(result, "FORECAST.LINEAR", "PREDICTION")
    result = ReplaceFunc(result, "FORECAST", "PREDICTION")
    result = ReplaceFunc(result, "FREQUENCY", "FREQUENCY")
    result = ReplaceFunc(result, "F.TEST", "F.TEST")
    result = ReplaceFunc(result, "GAMMA.DIST", "GAMMA.DIST")
    result = ReplaceFunc(result, "GAMMA.INV", "GAMMA.OBR")
    result = ReplaceFunc(result, "GAMMALN.PRECISE", "GAMMAL ACCURACY")
    result = ReplaceFunc(result, "GAMMALN", "GAMMALN")
    result = ReplaceFunc(result, "GAMMA", "GAMMA")
    result = ReplaceFunc(result, "GAUSS", "GAUSS")
    result = ReplaceFunc(result, "GEOMEAN", "SRGEOM")
    result = ReplaceFunc(result, "GROWTH", "HEIGHT")
    result = ReplaceFunc(result, "HARMEAN", "SRGARM")
    result = ReplaceFunc(result, "HYPGEOM.DIST", "HYPERGEOM.DIST")
    result = ReplaceFunc(result, "INTERCEPT", "CUT")
    result = ReplaceFunc(result, "KURT", "EXCESS")
    result = ReplaceFunc(result, "LARGE", "BIGGEST")
    result = ReplaceFunc(result, "LINEST", "LINEST")
    result = ReplaceFunc(result, "LOGEST", "LGRFPRIBL")
    result = ReplaceFunc(result, "LOGNORM.DIST", "LOGNORM.DIST")
    result = ReplaceFunc(result, "LOGNORM.INV", "LOGNORM.REV")
    result = ReplaceFunc(result, "MAXA", "MAXA")
    result = ReplaceFunc(result, "MAXIFS", "MAXESLIMN")
    result = ReplaceFunc(result, "MAX", "MAX")
    result = ReplaceFunc(result, "MEDIAN", "MEDIAN")
    result = ReplaceFunc(result, "MINA", "MINE")
    result = ReplaceFunc(result, "MINIFS", "MINESLIMN")
    result = ReplaceFunc(result, "MIN", "MIN")
    result = ReplaceFunc(result, "MODE.MULT", "MODA.NSK")
    result = ReplaceFunc(result, "MODE.SNGL", "FASHION.ONE")
    result = ReplaceFunc(result, "MODE", "FASHION")
    result = ReplaceFunc(result, "NEGBINOM.DIST", "OTRBINOM.DIST.")
    result = ReplaceFunc(result, "NORM.DIST", "NORMAL DIST.")
    result = ReplaceFunc(result, "NORM.INV", "NORM.REV")
    result = ReplaceFunc(result, "NORM.S.DIST", "NORM.ST.DIST.")
    result = ReplaceFunc(result, "NORM.S.INV", "NORM.ST.REV")
    result = ReplaceFunc(result, "PEARSON", "PEARSON")
    result = ReplaceFunc(result, "PERCENTILE.EXC", "PERCENTILE.EXC.")
    result = ReplaceFunc(result, "PERCENTILE.INC", "PERCENTILE.ON")
    result = ReplaceFunc(result, "PERCENTILE", "PERCENTILE")
    result = ReplaceFunc(result, "PERCENTRANK.EXC", "PERCENTRANK.EXCL.")
    result = ReplaceFunc(result, "PERCENTRANK.INC", "PERCENTRANK.ON")
    result = ReplaceFunc(result, "PERCENTRANK", "PERCENTRANK")
    result = ReplaceFunc(result, "PERMUT", "STOP")
    result = ReplaceFunc(result, "PERMUTATIONA", "STOP")
    result = ReplaceFunc(result, "PHI", "FI")
    result = ReplaceFunc(result, "POISSON.DIST", "POISSON.DIST.")
    result = ReplaceFunc(result, "PROB", "PROBABILITY")
    result = ReplaceFunc(result, "QUARTILE.EXC", "QUARTILE.EXCL.")
    result = ReplaceFunc(result, "QUARTILE.INC", "QUARTILE.ON")
    result = ReplaceFunc(result, "QUARTILE", "QUARTILE")
    result = ReplaceFunc(result, "RANK.AVG", "RANK.SR")
    result = ReplaceFunc(result, "RANK.EQ", "RANK.RV")
    result = ReplaceFunc(result, "RANK", "RANK")
    result = ReplaceFunc(result, "RSQ", "KVPIERSON")
    result = ReplaceFunc(result, "SKEW.P", "SKOS.G")
    result = ReplaceFunc(result, "SKEW", "SKOS")
    result = ReplaceFunc(result, "SLOPE", "INCLINE")
    result = ReplaceFunc(result, "SMALL", "LEAST")
    result = ReplaceFunc(result, "STANDARDIZE", "NORMALIZATION")
    result = ReplaceFunc(result, "STDEV.P", "STDEV.G")
    result = ReplaceFunc(result, "STDEV.S", "STDEV.V")
    result = ReplaceFunc(result, "STDEVA", "STDEV")
    result = ReplaceFunc(result, "STDEVPA", "STDEV")
    result = ReplaceFunc(result, "STDEVP", "STDEV")
    result = ReplaceFunc(result, "STDEV", "STANDARD DEVIATION")
    result = ReplaceFunc(result, "STEYX", "STOSHYX")
    result = ReplaceFunc(result, "T.DIST.2T", "STUDENT.DIST.2X")
    result = ReplaceFunc(result, "T.DIST.RT", "STUDENT.DIST.PH")
    result = ReplaceFunc(result, "T.DIST", "STUDENT.DIST")
    result = ReplaceFunc(result, "TREND", "TREND")
    result = ReplaceFunc(result, "TRIMMEAN", "CURRENT AVERAGE")
    result = ReplaceFunc(result, "T.INV.2T", "STUDENT.OBR.2X")
    result = ReplaceFunc(result, "T.INV", "STUDENT.OBR")
    result = ReplaceFunc(result, "T.TEST", "STUDENT TEST")
    result = ReplaceFunc(result, "VAR.P", "DISP.G")
    result = ReplaceFunc(result, "VAR.S", "DISP.B")
    result = ReplaceFunc(result, "VARA", "DISPA")
    result = ReplaceFunc(result, "VARPA", "DISPRA")
    result = ReplaceFunc(result, "VARP", "DISPR")
    result = ReplaceFunc(result, "VAR", "DISP")
    result = ReplaceFunc(result, "WEIBULL.DIST", "WEIBULL.DIST")
    result = ReplaceFunc(result, "Z.TEST", "Z.TEST")
    
    ' === INFORMATIONAL ===
    result = ReplaceFunc(result, "CELL", "CELL")
    result = ReplaceFunc(result, "ERROR.TYPE", "ERROR TYPE")
    result = ReplaceFunc(result, "INFO", "INFORM")
    result = ReplaceFunc(result, "ISBLANK", "EMPTY")
    result = ReplaceFunc(result, "ISERR", "EOS")
    result = ReplaceFunc(result, "ISERROR", "ERROR")
    result = ReplaceFunc(result, "ISEVEN", "EVEN")
    result = ReplaceFunc(result, "ISFORMULA", "FORMULA")
    result = ReplaceFunc(result, "ISLOGICAL", "ELOGIC")
    result = ReplaceFunc(result, "ISNA", "UNM")
    result = ReplaceFunc(result, "ISNONTEXT", "ENETEXT")
    result = ReplaceFunc(result, "ISNUMBER", "ISNUMBER")
    result = ReplaceFunc(result, "ISODD", "NUTS")
    result = ReplaceFunc(result, "ISREF", "LINK")
    result = ReplaceFunc(result, "ISTEXT", "ETEXT")
    result = ReplaceFunc(result, "NA", "ND")
    result = ReplaceFunc(result, "SHEET", "SHEET")
    result = ReplaceFunc(result, "SHEETS", "SHEETS")
    result = ReplaceFunc(result, "TYPE", "TYPE")
    
    ' === FINANCIAL ===
    result = ReplaceFunc(result, "ACCRINT", "ACCUMULATED INCOME")
    result = ReplaceFunc(result, "ACCRINTM", "ACCUMULATED INCOME REDEMPTION")
    result = ReplaceFunc(result, "AMORDEGRC", "AMORUM")
    result = ReplaceFunc(result, "AMORLINC", "AMORUV")
    result = ReplaceFunc(result, "COUPDAYBS", "DAYSKUPONDO")
    result = ReplaceFunc(result, "COUPDAYS", "DAYSCOUPON")
    result = ReplaceFunc(result, "COUPDAYSNC", "DAYSCOUPONAFTER")
    result = ReplaceFunc(result, "COUPNCD", "DATECOUPONAFTER")
    result = ReplaceFunc(result, "COUPNUM", "COUPON NUMBER")
    result = ReplaceFunc(result, "COUPPCD", "DATACOUPONDO")
    result = ReplaceFunc(result, "CUMIPMT", "GENERAL PAYMENT")
    result = ReplaceFunc(result, "CUMPRINC", "TOTAL INCOME")
    result = ReplaceFunc(result, "DB", "FOO")
    result = ReplaceFunc(result, "DDB", "DDOB")
    result = ReplaceFunc(result, "DISC", "DISCOUNT")
    result = ReplaceFunc(result, "DOLLARDE", "RUBLE.DES")
    result = ReplaceFunc(result, "DOLLARFR", "RUBLE.FRACTION")
    result = ReplaceFunc(result, "DURATION", "DURATION")
    result = ReplaceFunc(result, "EFFECT", "EFFECT")
    result = ReplaceFunc(result, "FV", "BS")
    result = ReplaceFunc(result, "FVSCHEDULE", "BZSCHEDULE")
    result = ReplaceFunc(result, "INTRATE", "INORMA")
    result = ReplaceFunc(result, "IPMT", "PRPLT")
    result = ReplaceFunc(result, "IRR", "VSD")
    result = ReplaceFunc(result, "ISPMT", "PROCESS PAYMENT")
    result = ReplaceFunc(result, "MDURATION", "MDLIT")
    result = ReplaceFunc(result, "MIRR", "MVSD")
    result = ReplaceFunc(result, "NOMINAL", "RATING")
    result = ReplaceFunc(result, "NPER", "NPER")
    result = ReplaceFunc(result, "NPV", "NPV")
    result = ReplaceFunc(result, "ODDFPRICE", "PRICE UPON REGULAR")
    result = ReplaceFunc(result, "ODDFYIELD", "INCOMEPERVERNERG")
    result = ReplaceFunc(result, "ODDLPRICE", "PRICE REGULAR")
    result = ReplaceFunc(result, "ODDLYIELD", "INCOME AFTER REGULAR")
    result = ReplaceFunc(result, "PDURATION", "PDLIT")
    result = ReplaceFunc(result, "PMT", "PLT")
    result = ReplaceFunc(result, "PPMT", "OSPLT")
    result = ReplaceFunc(result, "PRICEDISC", "PRICE DISCOUNT")
    result = ReplaceFunc(result, "PRICEMAT", "PRICE CASH")
    result = ReplaceFunc(result, "PRICE", "PRICE")
    result = ReplaceFunc(result, "PV", "PS")
    result = ReplaceFunc(result, "RATE", "BID")
    result = ReplaceFunc(result, "RECEIVED", "RECEIVED")
    result = ReplaceFunc(result, "RRI", "EQ.RATE")
    result = ReplaceFunc(result, "SLN", "nuclear submarine")
    result = ReplaceFunc(result, "SYD", "ASCH")
    result = ReplaceFunc(result, "TBILLEQ", "RAVNOCHEK")
    result = ReplaceFunc(result, "TBILLPRICE", "PRICECHECK")
    result = ReplaceFunc(result, "TBILLYIELD", "INCOMECHECK")
    result = ReplaceFunc(result, "VDB", "POO")
    result = ReplaceFunc(result, "XIRR", "CHISTVNDOH")
    result = ReplaceFunc(result, "XNPV", "CHISTNZ")
    result = ReplaceFunc(result, "YIELDDISC", "INCOME DISCOUNT")
    result = ReplaceFunc(result, "YIELDMAT", "INCOME REDEMPTION")
    result = ReplaceFunc(result, "YIELD", "INCOME")
    
    ' === ENGINEERING ===
    result = ReplaceFunc(result, "BESSELI", "BESSEL.I")
    result = ReplaceFunc(result, "BESSELJ", "BESSEL.J")
    result = ReplaceFunc(result, "BESSELK", "BESSEL.K")
    result = ReplaceFunc(result, "BESSELY", "BESSEL.Y")
    result = ReplaceFunc(result, "BIN2DEC", "DV.V.DES")
    result = ReplaceFunc(result, "BIN2HEX", "DV.H.HEX")
    result = ReplaceFunc(result, "BIN2OCT", "DV.V.EIGHT")
    result = ReplaceFunc(result, "BITAND", "BIT.I")
    result = ReplaceFunc(result, "BITLSHIFT", "BIT.SHIFT")
    result = ReplaceFunc(result, "BITOR", "BIT.OR")
    result = ReplaceFunc(result, "BITRSHIFT", "BIT.SHIFT")
    result = ReplaceFunc(result, "BITXOR", "BIT.ORPED")
    result = ReplaceFunc(result, "COMPLEX", "COMPLEX")
    result = ReplaceFunc(result, "CONVERT", "CONVERT")
    result = ReplaceFunc(result, "DEC2BIN", "DES.V.DV")
    result = ReplaceFunc(result, "DEC2HEX", "DES.HEX")
    result = ReplaceFunc(result, "DEC2OCT", "DES.V.EIGHT")
    result = ReplaceFunc(result, "DELTA", "DELTA")
    result = ReplaceFunc(result, "ERF.PRECISE", "FOS.EXACT")
    result = ReplaceFunc(result, "ERFC.PRECISE", "DFOSH.PRECISION")
    result = ReplaceFunc(result, "ERFC", "DFOSH")
    result = ReplaceFunc(result, "ERF", "FOS")
    result = ReplaceFunc(result, "GESTEP", "THRESHOLD")
    result = ReplaceFunc(result, "HEX2BIN", "HEX H. DW")
    result = ReplaceFunc(result, "HEX2DEC", "HEX.V.DES")
    result = ReplaceFunc(result, "HEX2OCT", "HEX.EIGHT")
    result = ReplaceFunc(result, "IMABS", "IMAG.ABS")
    result = ReplaceFunc(result, "IMAGINARY", "IMAGINARY PART")
    result = ReplaceFunc(result, "IMARGUMENT", "IMAGINAL ARGUMENT")
    result = ReplaceFunc(result, "IMCONJUGATE", "IMAGINAL MATE")
    result = ReplaceFunc(result, "IMCOS", "IMIM.COS")
    result = ReplaceFunc(result, "IMCOSH", "INSIM.COSH")
    result = ReplaceFunc(result, "IMCOT", "INSIM.COT")
    result = ReplaceFunc(result, "IMCSC", "IMIM.CSC")
    result = ReplaceFunc(result, "IMCSCH", "IMIM.CSCH")
    result = ReplaceFunc(result, "IMDIV", "IMAGINAL CASE")
    result = ReplaceFunc(result, "IMEXP", "IMAG.EXP")
    result = ReplaceFunc(result, "IMLN", "IMAG.LN")
    result = ReplaceFunc(result, "IMLOG10", "IMAG.LOG10")
    result = ReplaceFunc(result, "IMLOG2", "IMAG.LOG2")
    result = ReplaceFunc(result, "IMPOWER", "IMAGINARY DEGREE")
    result = ReplaceFunc(result, "IMPRODUCT", "IMAGINAL PRODUCT")
    result = ReplaceFunc(result, "IMREAL", "IMAGINAL THINGS")
    result = ReplaceFunc(result, "IMSEC", "IMAG.SEC")
    result = ReplaceFunc(result, "IMSECH", "INSIM.SECH")
    result = ReplaceFunc(result, "IMSIN", "IMAG.SIN")
    result = ReplaceFunc(result, "IMSINH", "IMAG.SINH")
    result = ReplaceFunc(result, "IMSQRT", "IMAGINAL ROOT")
    result = ReplaceFunc(result, "IMSUB", "IMAGINAL DIFFERENCE")
    result = ReplaceFunc(result, "IMSUM", "IMAG.SUM")
    result = ReplaceFunc(result, "IMTAN", "IMAG.TAN")
    result = ReplaceFunc(result, "OCT2BIN", "EIGHT.H.DV")
    result = ReplaceFunc(result, "OCT2DEC", "EIGHT.V.DES")
    result = ReplaceFunc(result, "OCT2HEX", "EIGHT.HEX")
    
    ' === DATABASES ===
    result = ReplaceFunc(result, "DAVERAGE", "DSRVALUE")
    result = ReplaceFunc(result, "DCOUNT", "COUNT")
    result = ReplaceFunc(result, "DCOUNTA", "ACCOUNTS")
    result = ReplaceFunc(result, "DGET", "BIZVLECH")
    result = ReplaceFunc(result, "DMAX", "DMAX")
    result = ReplaceFunc(result, "DMIN", "DMIN")
    result = ReplaceFunc(result, "DPRODUCT", "BDPRODUCT")
    result = ReplaceFunc(result, "DSTDEVP", "DSTDEV")
    result = ReplaceFunc(result, "DSTDEV", "DSTANDOFF")
    result = ReplaceFunc(result, "DSUM", "BDSUMM")
    result = ReplaceFunc(result, "DVARP", "BDDISPP")
    result = ReplaceFunc(result, "DVAR", "BDDISP")
    
    ' === WEB ===
    result = ReplaceFunc(result, "ENCODEURL", "ENCODINGURL")
    result = ReplaceFunc(result, "FILTERXML", "FILTER.XML")
    result = ReplaceFunc(result, "WEBSERVICE", "WEB SERVICE")
    
    ' === DYNAMIC ARRAYS (Excel 365) ===
    result = ReplaceFunc(result, "FILTER", "FILTER")
    result = ReplaceFunc(result, "RANDARRAY", "RANDARY")
    result = ReplaceFunc(result, "SEQUENCE", "AFTERNOON")
    result = ReplaceFunc(result, "SORTBY", "SORTPO")
    result = ReplaceFunc(result, "SORT", "GRADE")
    result = ReplaceFunc(result, "UNIQUE", "UNIQ")
    
    ' === OTHER ===
    result = ReplaceFunc(result, "EUROCONVERT", "EURO")
    result = ReplaceFunc(result, "N", "H")
    result = ReplaceFunc(result, "T", "T")
    
    ' Replacing TRUE/FALSE constants
    result = ReplaceConstant(result, "TRUE", "TRUE")
    result = ReplaceConstant(result, "FALSE", "LIE")
    
    ' Replace the argument separator (comma -> semicolon for Russian locale)
    result = Replace(result, ",", Application.International(xlListSeparator))
    
    LocalizeFormula = result
End Function

Private Function GetColumnNumber(colRef As String, rng As Range, ws As Worksheet) As Long
    Dim col As String
    col = Trim(colRef)
    
    If IsNumeric(col) Then
        ' Already a number - return it as is
        GetColumnNumber = CLng(col)
    Else
        ' Column letter - calculate the number relative to the beginning of the range
        On Error Resume Next
        Dim absColNum As Long
        absColNum = ws.columns(col).Column ' Absolute column number (A=1, B=2...)
        
        If absColNum > 0 Then
            ' Calculate the relative number in a range
            GetColumnNumber = absColNum - rng.Column + 1
            If GetColumnNumber < 1 Then GetColumnNumber = 1
            If GetColumnNumber > rng.columns.Count Then GetColumnNumber = rng.columns.Count
        Else
            GetColumnNumber = 1 ' Default first column
        End If
        On Error GoTo 0
    End If
End Function

Private Function ReplaceFunc(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String
    
    result = formula
    pos = InStr(1, UCase(result), UCase(engName) & "(")
    
    Do While pos > 0
        ' We check that there is no letter before the function name (so as not to replace part of another function)
        If pos = 1 Then
            result = rusName & Mid(result, pos + Len(engName))
        Else
            before = Mid(result, pos - 1, 1)
            If Not (before >= "A" And before <= "Z") And Not (before >= "a" And before <= "z") And Not (before >= "A" And before <= "I") Then
                result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            End If
        End If
        pos = InStr(pos + Len(rusName), UCase(result), UCase(engName) & "(")
    Loop
    
    ReplaceFunc = result
End Function

Private Function ReplaceConstant(formula As String, engName As String, rusName As String) As String
    Dim result As String
    Dim pos As Long
    Dim before As String
    Dim after As String
    Dim isWordBoundary As Boolean
    
    result = formula
    pos = InStr(1, UCase(result), UCase(engName))
    
    Do While pos > 0
        isWordBoundary = True
        
        ' Checking the symbol before
        If pos > 1 Then
            before = Mid(result, pos - 1, 1)
            If (before >= "A" And before <= "Z") Or (before >= "a" And before <= "z") Or (before >= "A" And before <= "I") Or (before >= "0" And before <= "9") Then
                isWordBoundary = False
            End If
        End If
        
        ' Checking the symbol after
        If pos + Len(engName) <= Len(result) Then
            after = Mid(result, pos + Len(engName), 1)
            If (after >= "A" And after <= "Z") Or (after >= "a" And after <= "z") Or (after >= "A" And after <= "I") Or (after >= "0" And after <= "9") Then
                isWordBoundary = False
            End If
        End If
        
        If isWordBoundary Then
            result = Left(result, pos - 1) & rusName & Mid(result, pos + Len(engName))
            pos = InStr(pos + Len(rusName), UCase(result), UCase(engName))
        Else
            pos = InStr(pos + 1, UCase(result), UCase(engName))
        End If
    Loop
    
    ReplaceConstant = result
End Function

Private Function GetChartType(chartTypeName As String) As Long
    Select Case UCase(Trim(chartTypeName))
        Case "LINE": GetChartType = 4 ' xlLine
        Case "BAR": GetChartType = 57 ' xlBarClustered
        Case "COLUMN": GetChartType = 51 ' xlColumnClustered
        Case "PIE": GetChartType = 5 ' xlPie
        Case "AREA": GetChartType = 1 ' xlArea
        Case "SCATTER", "XY": GetChartType = -4169 ' xlXYScatter
        Case "DOUGHNUT": GetChartType = -4120 ' xlDoughnut
        Case "RADAR": GetChartType = -4151 ' xlRadar
        Case "SURFACE": GetChartType = 83 ' xlSurface
        Case "BUBBLE": GetChartType = 15 ' xlBubble
        Case "STOCK": GetChartType = 88 ' xlStockHLC
        Case "CYLINDER": GetChartType = 95 ' xlCylinderCol
        Case "CONE": GetChartType = 99 ' xlConeCol
        Case "PYRAMID": GetChartType = 103 ' xlPyramidCol
        Case "LINE_MARKERS": GetChartType = 65 ' xlLineMarkers
        Case "AREA_STACKED": GetChartType = 76 ' xlAreaStacked
        Case "BAR_STACKED": GetChartType = 58 ' xlBarStacked
        Case "COLUMN_STACKED": GetChartType = 52 ' xlColumnStacked
        Case Else: GetChartType = 4 ' xlLine by default
    End Select
End Function

Private Function FindPivotTable(pivotName As String) As pivotTable
    Dim wsSearch As Worksheet
    Dim ptSearch As pivotTable
    Dim searchName As String
    
    searchName = Trim(pivotName)
    Set FindPivotTable = Nothing
    
    On Error Resume Next
    ' Search all workbook sheets
    For Each wsSearch In ActiveWorkbook.Worksheets
        For Each ptSearch In wsSearch.PivotTables
            If ptSearch.Name = searchName Then
                Set FindPivotTable = ptSearch
                Exit Function
            End If
        Next ptSearch
    Next wsSearch
    On Error GoTo 0
End Function

Private Function GetChartIndex(ws As Worksheet, indexStr As String) As Long
    Dim idx As Long
    Dim s As String
    
    s = UCase(Trim(indexStr))
    
    ' If 0, LAST, NEW, or empty: return last chart
    If s = "0" Or s = "LAST" Or s = "NEW" Or s = "" Then
        GetChartIndex = ws.ChartObjects.Count
        Exit Function
    End If
    
    ' Let's try to convert it to a number
    On Error Resume Next
    idx = CLng(s)
    On Error GoTo 0
    
    If idx > 0 Then
        GetChartIndex = idx
    Else
        ' Default - last
        GetChartIndex = ws.ChartObjects.Count
    End If
End Function

Private Function MakeSafeTableName(rawName As String) As String
    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim outName As String
    
    s = Trim$(rawName)
    If Len(s) = 0 Then
        MakeSafeTableName = ""
        Exit Function
    End If
    
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            outName = outName & ch
        ElseIf ch = " " Or ch = "-" Then
            outName = outName & "_"
        End If
    Next i
    
    If Len(outName) = 0 Then outName = "Table_AI"
    If outName Like "[0-9]*" Then outName = "T_" & outName
    If Len(outName) > 255 Then outName = Left$(outName, 255)
    
    MakeSafeTableName = outName
End Function

Private Function ExecuteSingleCommand(cmd As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim parts() As String
    Dim action As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim i As Long
    
    parts = Split(cmd, "|")
    If UBound(parts) < 0 Then
        ExecuteSingleCommand = False
        Exit Function
    End If
    
    action = UCase(Trim(parts(0)))
    Set ws = ActiveSheet
    
    ' Debugging: showing the command and number of parts
    Debug.Print "CMD: " & action & " | Parts: " & (UBound(parts) + 1) & " | Full: " & cmd
    
    Select Case action
    
        ' ========== WORKING WITH CELLS ==========
        
        Case "SET_VALUE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).value = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "SET_FORMULA"
            If UBound(parts) >= 2 Then
                Dim localFormula As String
                localFormula = LocalizeFormula(parts(2))
                Debug.Print "SET_FORMULA: Original=" & parts(2)
                Debug.Print "SET_FORMULA: Localized=" & localFormula
                ' Using FormulaLocal for Russian formulas
                ws.Range(parts(1)).FormulaLocal = localFormula
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_DOWN"
            If UBound(parts) >= 2 Then
                Dim srcCell As Range, destRng As Range
                Set srcCell = ws.Range(parts(1))
                Set destRng = ws.Range(parts(1) & ":" & parts(2))
                srcCell.Copy
                destRng.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_RIGHT"
            If UBound(parts) >= 2 Then
                Dim srcCellR As Range, destRngR As Range
                Set srcCellR = ws.Range(parts(1))
                Set destRngR = ws.Range(parts(1) & ":" & parts(2))
                srcCellR.Copy
                destRngR.PasteSpecial xlPasteFormulas
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_SERIES"
            If UBound(parts) >= 2 Then
                Dim stepVal As Double
                stepVal = 1
                If UBound(parts) >= 2 Then stepVal = CDbl(parts(2))
                ws.Range(parts(1)).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=stepVal
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_CONTENTS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearContents
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).ClearFormats
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_ALL"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Clear
                ExecuteSingleCommand = True
            End If
            
        Case "COPY"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "CUT"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Cut Destination:=ws.Range(parts(2))
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "PASTE_VALUES"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteValues
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
            
        Case "TRANSPOSE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Copy
                ws.Range(parts(2)).PasteSpecial Paste:=xlPasteAll, Transpose:=True
                Application.CutCopyMode = False
                ExecuteSingleCommand = True
            End If
        
        ' ========== FORMATTING ==========
            
        Case "BOLD"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Bold = True
                ExecuteSingleCommand = True
            End If
            
        Case "ITALIC"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Italic = True
                ExecuteSingleCommand = True
            End If
            
        Case "UNDERLINE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Underline = xlUnderlineStyleSingle
                ExecuteSingleCommand = True
            End If
            
        Case "STRIKETHROUGH"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Font.Strikethrough = True
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_NAME"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_SIZE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Size = CLng(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "FONT_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Font.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "FILL_COLOR"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "BORDER"
            If UBound(parts) >= 2 Then
                Dim borderStyle As String
                borderStyle = UCase(Trim(parts(2)))
                With ws.Range(parts(1))
                    Select Case borderStyle
                        Case "ALL"
                            .Borders.LineStyle = xlContinuous
                        Case "TOP"
                            .Borders(xlEdgeTop).LineStyle = xlContinuous
                        Case "BOTTOM"
                            .Borders(xlEdgeBottom).LineStyle = xlContinuous
                        Case "LEFT"
                            .Borders(xlEdgeLeft).LineStyle = xlContinuous
                        Case "RIGHT"
                            .Borders(xlEdgeRight).LineStyle = xlContinuous
                        Case "NONE"
                            .Borders.LineStyle = xlNone
                    End Select
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "BORDER_THICK"
            If UBound(parts) >= 1 Then
                With ws.Range(parts(1)).Borders
                    .LineStyle = xlContinuous
                    .Weight = xlMedium
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "ALIGN_H"
            If UBound(parts) >= 2 Then
                Dim hAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "LEFT": hAlign = xlLeft
                    Case "CENTER": hAlign = xlCenter
                    Case "RIGHT": hAlign = xlRight
                    Case "JUSTIFY": hAlign = xlJustify
                    Case Else: hAlign = xlGeneral
                End Select
                ws.Range(parts(1)).HorizontalAlignment = hAlign
                ExecuteSingleCommand = True
            End If
            
        Case "ALIGN_V"
            If UBound(parts) >= 2 Then
                Dim vAlign As Long
                Select Case UCase(Trim(parts(2)))
                    Case "TOP": vAlign = xlTop
                    Case "CENTER": vAlign = xlCenter
                    Case "BOTTOM": vAlign = xlBottom
                    Case Else: vAlign = xlCenter
                End Select
                ws.Range(parts(1)).VerticalAlignment = vAlign
                ExecuteSingleCommand = True
            End If
            
        Case "WRAP_TEXT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).WrapText = True
                ExecuteSingleCommand = True
            End If
            
        Case "MERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Merge
                ExecuteSingleCommand = True
            End If
            
        Case "UNMERGE"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).UnMerge
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_NUMBER"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_DATE"
            If UBound(parts) >= 2 Then
                ws.Range(parts(1)).NumberFormat = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_PERCENT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).NumberFormat = "0.00%"
                ExecuteSingleCommand = True
            End If
            
        Case "FORMAT_CURRENCY"
            If UBound(parts) >= 1 Then
                Dim currSymbol As String
                currSymbol = "?"
                If UBound(parts) >= 2 Then currSymbol = parts(2)
                ws.Range(parts(1)).NumberFormat = "#,##0.00 " & currSymbol
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFIT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).columns.AutoFit
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFIT_ROWS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Rows.AutoFit
                ExecuteSingleCommand = True
            End If
            
        Case "COLUMN_WIDTH"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1)).ColumnWidth = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "ROW_HEIGHT"
            If UBound(parts) >= 2 Then
                ws.Rows(CLng(parts(1))).RowHeight = CDbl(parts(2))
                ExecuteSingleCommand = True
            End If
        
        ' ========== ROWS AND COLUMN ==========
            
        Case "INSERT_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_ROWS"
            If UBound(parts) >= 2 Then
                Dim rowNum As Long, rowCount As Long
                rowNum = CLng(parts(1))
                rowCount = CLng(parts(2))
                ws.Rows(rowNum & ":" & (rowNum + rowCount - 1)).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Insert
                ExecuteSingleCommand = True
            End If
            
        Case "INSERT_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim colNum As Long, colCount As Long
                colCount = CLng(parts(2))
                For i = 1 To colCount
                    ws.columns(parts(1)).Insert
                Next i
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Delete
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ROW"
            If UBound(parts) >= 1 Then
                ws.Rows(CLng(parts(1))).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = True
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_COLUMN"
            If UBound(parts) >= 1 Then
                ws.columns(parts(1)).Hidden = False
                ExecuteSingleCommand = True
            End If
            
        Case "GROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If
            
        Case "UNGROUP_ROWS"
            If UBound(parts) >= 2 Then
                ws.Rows(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If
            
        Case "GROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Group
                ExecuteSingleCommand = True
            End If
            
        Case "UNGROUP_COLUMNS"
            If UBound(parts) >= 2 Then
                ws.columns(parts(1) & ":" & parts(2)).Ungroup
                ExecuteSingleCommand = True
            End If
        
        ' ========== SORTING AND FILTERING ==========
            
        Case "SORT"
            If UBound(parts) >= 3 Then
                Dim SortRange As Range, sortCol As Long, sortOrder As Long
                Dim sortColStr As String, sortKeyRange As Range
                Set SortRange = ws.Range(parts(1))
                sortColStr = Trim(parts(2))
                
                ' Defining a column for sorting
                If IsNumeric(sortColStr) Then
                    ' Number - column number in the range (1, 2, 3...)
                    sortCol = CLng(sortColStr)
                    Set sortKeyRange = SortRange.columns(sortCol)
                Else
                    ' Column letter (A, B, C...) - use intersection with range
                    Set sortKeyRange = Intersect(SortRange, ws.columns(sortColStr))
                    If sortKeyRange Is Nothing Then
                        ' If the letter is out of range, take the first column
                        Set sortKeyRange = SortRange.columns(1)
                    End If
                End If
                
                sortOrder = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                SortRange.Sort Key1:=sortKeyRange, Order1:=sortOrder, Header:=xlGuess
                ExecuteSingleCommand = True
            End If
            
        Case "SORT_MULTI"
            If UBound(parts) >= 5 Then
                Dim sRng As Range
                Dim sCol1 As Long, sCol2 As Long
                Set sRng = ws.Range(parts(1))
                Dim o1 As Long, o2 As Long
                sCol1 = GetColumnNumber(parts(2), sRng, ws)
                sCol2 = GetColumnNumber(parts(4), sRng, ws)
                o1 = IIf(UCase(parts(3)) = "ASC", xlAscending, xlDescending)
                o2 = IIf(UCase(parts(5)) = "ASC", xlAscending, xlDescending)
                sRng.Sort Key1:=sRng.columns(sCol1), Order1:=o1, _
                          Key2:=sRng.columns(sCol2), Order2:=o2, Header:=xlGuess
                ExecuteSingleCommand = True
            End If
            
        Case "AUTOFILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then ws.AutoFilterMode = False
                ws.Range(parts(1)).AutoFilter
                ExecuteSingleCommand = True
            End If
            
        Case "FILTER"
            If UBound(parts) >= 3 Then
                Dim fRng As Range
                Dim fCol As Long
                Set fRng = ws.Range(parts(1))
                fCol = GetColumnNumber(parts(2), fRng, ws)
                If Not ws.AutoFilterMode Then fRng.AutoFilter
                fRng.AutoFilter Field:=fCol, Criteria1:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "FILTER_TOP"
            If UBound(parts) >= 3 Then
                Dim ftRng As Range
                Dim ftCol As Long
                Set ftRng = ws.Range(parts(1))
                ftCol = GetColumnNumber(parts(2), ftRng, ws)
                If Not ws.AutoFilterMode Then ftRng.AutoFilter
                ftRng.AutoFilter Field:=ftCol, Criteria1:=CLng(parts(3)), Operator:=xlTop10Items
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_FILTER"
            If UBound(parts) >= 1 Then
                If ws.AutoFilterMode Then
                    ws.Range(parts(1)).AutoFilter
                    ws.Range(parts(1)).AutoFilter
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_AUTOFILTER"
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ExecuteSingleCommand = True
            
        Case "REMOVE_DUPLICATES"
            If UBound(parts) >= 2 Then
                Dim dupRng As Range
                Dim colsArr() As Long
                Dim colsList() As String
                Set dupRng = ws.Range(parts(1))
                colsList = Split(parts(2), ",")
                ReDim colsArr(UBound(colsList))
                For i = 0 To UBound(colsList)
                    colsArr(i) = CLng(Trim(colsList(i)))
                Next i
                dupRng.RemoveDuplicates columns:=colsArr, Header:=xlYes
                ExecuteSingleCommand = True
            End If

        Case "CREATE_TABLE"
            If UBound(parts) >= 1 Then
                Dim tblRng As Range
                Dim tbl As ListObject
                Dim desiredName As String
                Dim candidateName As String
                Dim existingTbl As ListObject
                
                Set tblRng = ws.Range(parts(1))
                
                ' Reuse existing table if the range is already inside one.
                For Each existingTbl In ws.ListObjects
                    If Not Intersect(existingTbl.Range, tblRng) Is Nothing Then
                        existingTbl.TableStyle = "TableStyleMedium2"
                        If UBound(parts) >= 2 Then
                            desiredName = MakeSafeTableName(parts(2))
                            If Len(desiredName) > 0 Then
                                On Error Resume Next
                                existingTbl.Name = desiredName
                                On Error GoTo ErrorHandler
                            End If
                        End If
                        ExecuteSingleCommand = True
                        Exit Function
                    End If
                Next existingTbl
                
                Set tbl = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=tblRng, XlListObjectHasHeaders:=xlYes)
                tbl.TableStyle = "TableStyleMedium2"
                
                If UBound(parts) >= 2 Then
                    candidateName = MakeSafeTableName(parts(2))
                    If Len(candidateName) > 0 Then
                        On Error Resume Next
                        tbl.Name = candidateName
                        On Error GoTo ErrorHandler
                    End If
                End If
                
                ExecuteSingleCommand = True
            End If
             
        Case "FIND_REPLACE"
            If UBound(parts) >= 2 Then
                Dim replaceWith As String
                replaceWith = ""
                If UBound(parts) >= 2 Then replaceWith = parts(2)
                ws.UsedRange.Replace What:=parts(1), Replacement:=replaceWith, LookAt:=xlPart
                ExecuteSingleCommand = True
            End If
            
        Case "FIND_REPLACE_RANGE"
            If UBound(parts) >= 3 Then
                ws.Range(parts(1)).Replace What:=parts(2), Replacement:=parts(3), LookAt:=xlPart
                ExecuteSingleCommand = True
            End If
        
        ' ========== CHARTS ==========
            
        Case "CREATE_CHART"
            ' CREATE_CHART|range|type|name
            ' Supports non-adjacent ranges: A2:A5,B2:B5
            If UBound(parts) >= 2 Then
                Dim chartObj As ChartObject
                Dim chartType As Long
                Dim dataRange As Range
                Dim chartLeft As Double, chartTop As Double
                Dim rangeStr As String
                Dim rangeParts() As String
                Dim rngPart As Variant
                
                rangeStr = Trim(parts(1))
                
                ' Checking for non-adjacent ranges
                On Error Resume Next
                If InStr(rangeStr, ",") > 0 Then
                    ' Non-contiguous ranges
                    rangeParts = Split(rangeStr, ",")
                    Set dataRange = ws.Range(Trim(rangeParts(0)))
                    For i = 1 To UBound(rangeParts)
                        Set dataRange = Union(dataRange, ws.Range(Trim(rangeParts(i))))
                    Next i
                Else
                    Set dataRange = ws.Range(rangeStr)
                End If
                On Error GoTo ErrorHandler
                
                If dataRange Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If
                
                chartType = GetChartType(parts(2))
                
                ' Position the graph to the right of the data
                chartLeft = dataRange.Areas(1).Cells(1, dataRange.Areas(1).columns.Count).Offset(0, 2).Left
                chartTop = dataRange.Areas(1).Top
                
                Set chartObj = ws.ChartObjects.Add(Left:=chartLeft, Top:=chartTop, Width:=400, Height:=250)
                chartObj.Chart.SetSourceData Source:=dataRange
                chartObj.Chart.chartType = chartType
                
                If UBound(parts) >= 3 Then
                    If Len(Trim(parts(3))) > 0 Then
                        chartObj.Chart.HasTitle = True
                        chartObj.Chart.ChartTitle.text = parts(3)
                    End If
                End If
                
                ExecuteSingleCommand = True
            End If
            
        Case "CREATE_CHART_POS", "CREATE_CHART_AT"
            ' CREATE_CHART_POS|range|type|title|cell_or_left|top|width|height
            ' CREATE_CHART_AT|range|type|name|cell - simplified version
            ' Supports non-adjacent ranges: A2:A5,B2:B5
            If UBound(parts) >= 3 Then
                Dim chartObj2 As ChartObject
                Dim chartType2 As Long
                Dim dataRange2 As Range
                Dim posLeft As Double, posTop As Double
                Dim chWidth As Double, chHeight As Double
                Dim rangeStr2 As String
                Dim rangeParts2() As String
                
                rangeStr2 = Trim(parts(1))
                
                ' Checking for non-adjacent ranges
                On Error Resume Next
                If InStr(rangeStr2, ",") > 0 Then
                    rangeParts2 = Split(rangeStr2, ",")
                    Set dataRange2 = ws.Range(Trim(rangeParts2(0)))
                    For i = 1 To UBound(rangeParts2)
                        Set dataRange2 = Union(dataRange2, ws.Range(Trim(rangeParts2(i))))
                    Next i
                Else
                    Set dataRange2 = ws.Range(rangeStr2)
                End If
                On Error GoTo ErrorHandler
                
                If dataRange2 Is Nothing Then
                    ExecuteSingleCommand = False
                    Exit Function
                End If
                
                chartType2 = GetChartType(parts(2))
                
                ' Default values
                chWidth = 400
                chHeight = 250
                posLeft = 300
                posTop = 50
                
                ' Determining the position
                If UBound(parts) >= 4 Then
                    ' Check if this is a cell address or a number
                    On Error Resume Next
                    Dim posCell As Range
                    Set posCell = ws.Range(parts(4))
                    If Not posCell Is Nothing Then
                        ' This is the cell address
                        posLeft = posCell.Left
                        posTop = posCell.Top
                    Else
                        ' This number
                        posLeft = CDbl(parts(4))
                    End If
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 5 Then
                    On Error Resume Next
                    posTop = CDbl(parts(5))
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 6 Then
                    On Error Resume Next
                    chWidth = CDbl(parts(6))
                    On Error GoTo ErrorHandler
                End If
                
                If UBound(parts) >= 7 Then
                    On Error Resume Next
                    chHeight = CDbl(parts(7))
                    On Error GoTo ErrorHandler
                End If
                
                Set chartObj2 = ws.ChartObjects.Add(Left:=posLeft, Top:=posTop, Width:=chWidth, Height:=chHeight)
                chartObj2.Chart.SetSourceData Source:=dataRange2
                chartObj2.Chart.chartType = chartType2
                
                If Len(Trim(parts(3))) > 0 Then
                    chartObj2.Chart.HasTitle = True
                    chartObj2.Chart.ChartTitle.text = parts(3)
                End If
                
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_TITLE"
            ' CHART_TITLE|index|text (index: 1 = first chart, 0 or LAST = last)
            If UBound(parts) >= 2 Then
                Dim chIdx As Long
                chIdx = GetChartIndex(ws, parts(1))
                If chIdx > 0 And chIdx <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx).Chart.HasTitle = True
                    ws.ChartObjects(chIdx).Chart.ChartTitle.text = parts(2)
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_LEGEND"
            If UBound(parts) >= 2 Then
                Dim chIdx2 As Long
                chIdx2 = GetChartIndex(ws, parts(1))
                If chIdx2 > 0 And chIdx2 <= ws.ChartObjects.Count Then
                    Dim legPos As Long
                    Select Case UCase(Trim(parts(2)))
                        Case "TOP": legPos = xlLegendPositionTop
                        Case "BOTTOM": legPos = xlLegendPositionBottom
                        Case "LEFT": legPos = xlLegendPositionLeft
                        Case "RIGHT": legPos = xlLegendPositionRight
                        Case "NONE"
                            ws.ChartObjects(chIdx2).Chart.HasLegend = False
                            ExecuteSingleCommand = True
                            Exit Function
                        Case Else: legPos = xlLegendPositionBottom
                    End Select
                    ws.ChartObjects(chIdx2).Chart.HasLegend = True
                    ws.ChartObjects(chIdx2).Chart.Legend.Position = legPos
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_AXIS_TITLE"
            If UBound(parts) >= 3 Then
                Dim chIdx3 As Long
                chIdx3 = GetChartIndex(ws, parts(1))
                If chIdx3 > 0 And chIdx3 <= ws.ChartObjects.Count Then
                    On Error Resume Next
                    Dim ax As Object
                    If UCase(Trim(parts(2))) = "X" Then
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlCategory)
                    Else
                        Set ax = ws.ChartObjects(chIdx3).Chart.Axes(xlValue)
                    End If
                    If Not ax Is Nothing Then
                        ax.HasTitle = True
                        ax.AxisTitle.text = parts(3)
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_TYPE"
            If UBound(parts) >= 2 Then
                Dim chIdx4 As Long
                chIdx4 = GetChartIndex(ws, parts(1))
                If chIdx4 > 0 And chIdx4 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx4).Chart.chartType = GetChartType(parts(2))
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_MOVE"
            If UBound(parts) >= 2 Then
                Dim chIdx5 As Long
                chIdx5 = GetChartIndex(ws, parts(1))
                If chIdx5 > 0 And chIdx5 <= ws.ChartObjects.Count Then
                    ' Checking the cell address or coordinates
                    On Error Resume Next
                    Dim moveCell As Range
                    Set moveCell = ws.Range(parts(2))
                    If Not moveCell Is Nothing Then
                        ws.ChartObjects(chIdx5).Left = moveCell.Left
                        ws.ChartObjects(chIdx5).Top = moveCell.Top
                    ElseIf UBound(parts) >= 3 Then
                        ws.ChartObjects(chIdx5).Left = CLng(parts(2))
                        ws.ChartObjects(chIdx5).Top = CLng(parts(3))
                    End If
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_RESIZE"
            If UBound(parts) >= 3 Then
                Dim chIdx6 As Long
                chIdx6 = GetChartIndex(ws, parts(1))
                If chIdx6 > 0 And chIdx6 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx6).Width = CLng(parts(2))
                    ws.ChartObjects(chIdx6).Height = CLng(parts(3))
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_DELETE"
            If UBound(parts) >= 1 Then
                Dim chIdx7 As Long
                chIdx7 = GetChartIndex(ws, parts(1))
                If chIdx7 > 0 And chIdx7 <= ws.ChartObjects.Count Then
                    ws.ChartObjects(chIdx7).Delete
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "CHART_DELETE_ALL"
            Dim co As ChartObject
            For Each co In ws.ChartObjects
                co.Delete
            Next co
            ExecuteSingleCommand = True
            
        Case "MOVE_CHART"
            ' For compatibility - skip
            ExecuteSingleCommand = True
        
        ' ========== PIVOT TABLES ==========
            
        Case "CREATE_PIVOT"
            ' CREATE_PIVOT|source|destination|name
            ' Example: CREATE_PIVOT|A1:D10|F1|MyPivot
            ' Or: CREATE_PIVOT|Sheet1!A1:D10|Sheet2!A1|MyPivot
            If UBound(parts) >= 3 Then
                Dim pivotCache As pivotCache
                Dim pivotTable As pivotTable
                Dim srcRng As Range
                Dim destCell As Range
                Dim destSheet As Worksheet
                Dim pivotName As String
                
                ' Using Application.Range to support sheet name links
                On Error Resume Next
                Set srcRng = Application.Range(parts(1))
                If srcRng Is Nothing Then
                    ' Let's try without the sheet name
                    Set srcRng = ws.Range(parts(1))
                End If
                
                ' For assignment - check whether a new sheet needs to be created
                Set destCell = Application.Range(parts(2))
                If destCell Is Nothing Then
                    ' If the sheet does not exist, create it
                    Dim destParts() As String
                    If InStr(parts(2), "!") > 0 Then
                        destParts = Split(parts(2), "!")
                        Dim sheetName As String
                        sheetName = Replace(Replace(destParts(0), "'", ""), "!", "")
                        ' Checking the existence of a sheet
                        Dim sheetExists As Boolean
                        sheetExists = False
                        Dim wsCheck As Worksheet
                        For Each wsCheck In ActiveWorkbook.Worksheets
                            If wsCheck.Name = sheetName Then
                                sheetExists = True
                                Set destSheet = wsCheck
                                Exit For
                            End If
                        Next wsCheck
                        If Not sheetExists Then
                            Set destSheet = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count))
                            destSheet.Name = sheetName
                        End If
                        Set destCell = destSheet.Range(destParts(1))
                    Else
                        Set destCell = ws.Range(parts(2))
                    End If
                End If
                On Error GoTo ErrorHandler
                
                If Not srcRng Is Nothing And Not destCell Is Nothing Then
                    pivotName = Trim(parts(3))
                    
                    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=srcRng)
                        
                    Set pivotTable = pivotCache.CreatePivotTable( _
                        TableDestination:=destCell, _
                        tableName:=pivotName)
                        
                    ExecuteSingleCommand = True
                End If
            End If
            
        Case "PIVOT_ADD_ROW"
            If UBound(parts) >= 2 Then
                Dim pt As pivotTable
                Set pt = FindPivotTable(parts(1))
                If Not pt Is Nothing Then
                    On Error Resume Next
                    pt.PivotFields(parts(2)).Orientation = xlRowField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_COLUMN"
            If UBound(parts) >= 2 Then
                Dim pt2 As pivotTable
                Set pt2 = FindPivotTable(parts(1))
                If Not pt2 Is Nothing Then
                    On Error Resume Next
                    pt2.PivotFields(parts(2)).Orientation = xlColumnField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_VALUE"
            If UBound(parts) >= 3 Then
                Dim pt3 As pivotTable
                Dim pfFunc As Long
                Set pt3 = FindPivotTable(parts(1))
                If Not pt3 Is Nothing Then
                    Select Case UCase(Trim(parts(3)))
                        Case "SUM": pfFunc = xlSum
                        Case "COUNT": pfFunc = xlCount
                        Case "AVERAGE": pfFunc = xlAverage
                        Case "MAX": pfFunc = xlMax
                        Case "MIN": pfFunc = xlMin
                        Case Else: pfFunc = xlSum
                    End Select
                    On Error Resume Next
                    pt3.AddDataField pt3.PivotFields(parts(2)), , pfFunc
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_ADD_FILTER"
            If UBound(parts) >= 2 Then
                Dim pt4 As pivotTable
                Set pt4 = FindPivotTable(parts(1))
                If Not pt4 Is Nothing Then
                    On Error Resume Next
                    pt4.PivotFields(parts(2)).Orientation = xlPageField
                    On Error GoTo ErrorHandler
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_REFRESH"
            If UBound(parts) >= 1 Then
                Dim pt5 As pivotTable
                Set pt5 = FindPivotTable(parts(1))
                If Not pt5 Is Nothing Then
                    pt5.RefreshTable
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PIVOT_REFRESH_ALL"
            ActiveWorkbook.RefreshAll
            ExecuteSingleCommand = True
        
        ' ========== SHEETS ==========
            
        Case "ADD_SHEET"
            If UBound(parts) >= 1 Then
                Dim newSheet As Worksheet
                Set newSheet = ActiveWorkbook.Worksheets.Add
                newSheet.Name = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_SHEET_AFTER"
            If UBound(parts) >= 2 Then
                Dim newSheet2 As Worksheet
                Set newSheet2 = ActiveWorkbook.Worksheets.Add(after:=ActiveWorkbook.Worksheets(parts(2)))
                newSheet2.Name = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_SHEET"
            If UBound(parts) >= 1 Then
                Application.DisplayAlerts = False
                ActiveWorkbook.Worksheets(parts(1)).Delete
                Application.DisplayAlerts = True
                ExecuteSingleCommand = True
            End If
            
        Case "RENAME_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "COPY_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Copy after:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
                ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count).Name = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "MOVE_SHEET"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Move before:=ActiveWorkbook.Worksheets(CLng(parts(2)))
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetHidden
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Visible = xlSheetVisible
                ExecuteSingleCommand = True
            End If
            
        Case "ACTIVATE_SHEET"
            If UBound(parts) >= 1 Then
                ActiveWorkbook.Worksheets(parts(1)).Activate
                ExecuteSingleCommand = True
            End If
            
        Case "TAB_COLOR"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Worksheets(parts(1)).Tab.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "PROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd As String
                pwd = ""
                If UBound(parts) >= 2 Then pwd = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Protect Password:=pwd
                ExecuteSingleCommand = True
            End If
            
        Case "UNPROTECT_SHEET"
            If UBound(parts) >= 1 Then
                Dim pwd2 As String
                pwd2 = ""
                If UBound(parts) >= 2 Then pwd2 = parts(2)
                ActiveWorkbook.Worksheets(parts(1)).Unprotect Password:=pwd2
                ExecuteSingleCommand = True
            End If
        
        ' ========== NAMED RANGES ==========
            
        Case "CREATE_NAME"
            If UBound(parts) >= 2 Then
                ActiveWorkbook.Names.Add Name:=parts(1), RefersTo:=ws.Range(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_NAME"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ActiveWorkbook.Names(parts(1)).Delete
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If
        
        ' ========== CONDITIONAL FORMATTING ==========
            
        Case "COND_HIGHLIGHT"
            ' COND_HIGHLIGHT|range|formula|color (4 parts)
            ' COND_HIGHLIGHT|range|operator|value|color (5 parts)
            If UBound(parts) >= 3 Then
                Dim hlRng As Range
                Dim hlFormula As String
                Dim hlFirstCell As String
                Dim hlColor As String
                Dim hlOp As String
                Dim hlFC As Object
                
                Debug.Print "COND_HIGHLIGHT: Step 1 - parsing range: " & parts(1)
                Set hlRng = ws.Range(parts(1))
                Debug.Print "COND_HIGHLIGHT: Step 2 - range set OK: " & hlRng.address
                hlFirstCell = hlRng.Cells(1, 1).address(False, False)
                Debug.Print "COND_HIGHLIGHT: Step 3 - firstCell: " & hlFirstCell
                
                If UBound(parts) = 3 Then
                    ' 4 parts: range|formula|color
                    Debug.Print "COND_HIGHLIGHT: Step 4a - 4 parts mode"
                    hlFormula = Trim(parts(2))
                    hlColor = Trim(parts(3))
                Else
                    ' 5 parts: range|operator|value|color
                    Debug.Print "COND_HIGHLIGHT: Step 4b - 5 parts mode"
                    hlOp = Trim(parts(2))
                    hlColor = Trim(parts(4))
                    Debug.Print "COND_HIGHLIGHT: Step 5 - op=" & hlOp & " color=" & hlColor
                    
                    Select Case hlOp
                        Case ">", "<", ">=", "<=", "<>"
                            hlFormula = hlFirstCell & hlOp & Trim(parts(3))
                        Case "="
                            If IsNumeric(Trim(parts(3))) Then
                                hlFormula = hlFirstCell & "=" & Trim(parts(3))
                            Else
                                hlFormula = Trim(parts(3))
                            End If
                        Case Else
                            hlFormula = Trim(parts(3))
                    End Select
                End If
                
                Debug.Print "COND_HIGHLIGHT: Step 6 - formula before =: " & hlFormula
                
                ' Add = to the beginning if not
                If Left(hlFormula, 1) <> "=" Then hlFormula = "=" & hlFormula
                
                Debug.Print "COND_HIGHLIGHT: Step 7 - formula after =: " & hlFormula
                
                ' Localize the formula (MOD -> REMAIN, etc., comma -> semicolon)
                hlFormula = LocalizeFormula(hlFormula)
                Debug.Print "COND_HIGHLIGHT: Step 8 - formula localized: " & hlFormula
                Debug.Print "COND_HIGHLIGHT: Step 9 - adding FormatCondition..."
                
                ' IMPORTANT: Select the first cell of the range so that Excel correctly
                ' interpreted relative references in formula
                hlRng.Cells(1, 1).Select
                
                ' Adding conditional formatting
                Set hlFC = hlRng.FormatConditions.Add(Type:=xlExpression, Formula1:=hlFormula)
                
                Debug.Print "COND_HIGHLIGHT: Step 10 - hlColor=[" & hlColor & "]"
                
                Dim hlColorValue As Long
                On Error Resume Next
                hlColorValue = ParseColor(hlColor)
                If Err.Number <> 0 Then
                    Debug.Print "COND_HIGHLIGHT: ParseColor ERROR: " & Err.Number & " - " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
                
                Debug.Print "COND_HIGHLIGHT: Step 10a - colorValue=" & hlColorValue
                Debug.Print "COND_HIGHLIGHT: Step 10b - hlFC type=" & TypeName(hlFC)
                
                hlFC.Interior.Color = hlColorValue
                
                Debug.Print "COND_HIGHLIGHT: Step 11 - DONE"
                ExecuteSingleCommand = True
            End If
            
        Case "COND_TOP"
            If UBound(parts) >= 3 Then
                Dim cfRng2 As Range
                Set cfRng2 = ws.Range(parts(1))
                cfRng2.FormatConditions.AddTop10
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).TopBottom = xlTop10Top
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng2.FormatConditions(cfRng2.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_BOTTOM"
            If UBound(parts) >= 3 Then
                Dim cfRng3 As Range
                Set cfRng3 = ws.Range(parts(1))
                cfRng3.FormatConditions.AddTop10
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).TopBottom = xlTop10Bottom
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Rank = CLng(parts(2))
                cfRng3.FormatConditions(cfRng3.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_DUPLICATE"
            If UBound(parts) >= 2 Then
                Dim cfRng4 As Range
                Set cfRng4 = ws.Range(parts(1))
                cfRng4.FormatConditions.AddUniqueValues
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).DupeUnique = xlDuplicate
                cfRng4.FormatConditions(cfRng4.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_UNIQUE"
            If UBound(parts) >= 2 Then
                Dim cfRng5 As Range
                Set cfRng5 = ws.Range(parts(1))
                cfRng5.FormatConditions.AddUniqueValues
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).DupeUnique = xlUnique
                cfRng5.FormatConditions(cfRng5.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_TEXT"
            If UBound(parts) >= 3 Then
                Dim cfRng6 As Range
                Set cfRng6 = ws.Range(parts(1))
                cfRng6.FormatConditions.Add Type:=xlTextString, String:=parts(2), TextOperator:=xlContains
                cfRng6.FormatConditions(cfRng6.FormatConditions.Count).Interior.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng7 As Range
                Set cfRng7 = ws.Range(parts(1))
                cfRng7.FormatConditions.Add Type:=xlBlanksCondition
                cfRng7.FormatConditions(cfRng7.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COND_NOT_BLANK"
            If UBound(parts) >= 2 Then
                Dim cfRng8 As Range
                Set cfRng8 = ws.Range(parts(1))
                cfRng8.FormatConditions.Add Type:=xlNoBlanksCondition
                cfRng8.FormatConditions(cfRng8.FormatConditions.Count).Interior.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "DATA_BARS"
            If UBound(parts) >= 2 Then
                Dim cfRng9 As Range
                Set cfRng9 = ws.Range(parts(1))
                cfRng9.FormatConditions.AddDatabar
                cfRng9.FormatConditions(cfRng9.FormatConditions.Count).BarColor.Color = ParseColor(parts(2))
                ExecuteSingleCommand = True
            End If
            
        Case "COLOR_SCALE"
            If UBound(parts) >= 3 Then
                Dim cfRng10 As Range
                Set cfRng10 = ws.Range(parts(1))
                cfRng10.FormatConditions.AddColorScale ColorScaleType:=2
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng10.FormatConditions(cfRng10.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                ExecuteSingleCommand = True
            End If
            
        Case "COLOR_SCALE3"
            If UBound(parts) >= 4 Then
                Dim cfRng11 As Range
                Set cfRng11 = ws.Range(parts(1))
                cfRng11.FormatConditions.AddColorScale ColorScaleType:=3
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(1).FormatColor.Color = ParseColor(parts(2))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(2).FormatColor.Color = ParseColor(parts(3))
                cfRng11.FormatConditions(cfRng11.FormatConditions.Count).ColorScaleCriteria(3).FormatColor.Color = ParseColor(parts(4))
                ExecuteSingleCommand = True
            End If
            
        Case "ICON_SET"
            If UBound(parts) >= 2 Then
                Dim cfRng12 As Range
                Dim iconSetType As Long
                Set cfRng12 = ws.Range(parts(1))
                
                Select Case UCase(Trim(parts(2)))
                    Case "ARROWS": iconSetType = 1 ' xl3Arrows
                    Case "FLAGS": iconSetType = 7 ' xl3Flags
                    Case "STARS": iconSetType = 13 ' xl3Stars
                    Case "BARS": iconSetType = 14 ' xl4RedToBlack
                    Case Else: iconSetType = 1
                End Select
                
                cfRng12.FormatConditions.AddIconSetCondition
                cfRng12.FormatConditions(cfRng12.FormatConditions.Count).IconSet = ActiveWorkbook.IconSets(iconSetType)
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_COND_FORMAT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).FormatConditions.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== DATA CHECK ==========
            
        Case "VALIDATION_LIST"
            If UBound(parts) >= 2 Then
                Dim valRng As Range
                Set valRng = ws.Range(parts(1))
                valRng.Validation.Delete
                valRng.Validation.Add Type:=xlValidateList, Formula1:=Replace(parts(2), ";", ",")
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_NUMBER"
            If UBound(parts) >= 3 Then
                Dim valRng2 As Range
                Set valRng2 = ws.Range(parts(1))
                valRng2.Validation.Delete
                valRng2.Validation.Add Type:=xlValidateWholeNumber, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_DATE"
            If UBound(parts) >= 3 Then
                Dim valRng3 As Range
                Set valRng3 = ws.Range(parts(1))
                valRng3.Validation.Delete
                valRng3.Validation.Add Type:=xlValidateDate, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_TEXT_LENGTH"
            If UBound(parts) >= 3 Then
                Dim valRng4 As Range
                Set valRng4 = ws.Range(parts(1))
                valRng4.Validation.Delete
                valRng4.Validation.Add Type:=xlValidateTextLength, Operator:=xlBetween, Formula1:=parts(2), Formula2:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "VALIDATION_CUSTOM"
            If UBound(parts) >= 2 Then
                Dim valRng5 As Range
                Dim valFormula As String
                Set valRng5 = ws.Range(parts(1))
                valFormula = LocalizeFormula(parts(2))
                valRng5.Validation.Delete
                valRng5.Validation.Add Type:=xlValidateCustom, Formula1:=valFormula
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_VALIDATION"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Validation.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== COMMENTS ==========
            
        Case "ADD_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell As Range
                Set cmtCell = ws.Range(parts(1))
                If Not cmtCell.Comment Is Nothing Then cmtCell.Comment.Delete
                cmtCell.AddComment parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "EDIT_COMMENT"
            If UBound(parts) >= 2 Then
                Dim cmtCell2 As Range
                Set cmtCell2 = ws.Range(parts(1))
                If Not cmtCell2.Comment Is Nothing Then
                    cmtCell2.Comment.text text:=parts(2)
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell3 As Range
                Set cmtCell3 = ws.Range(parts(1))
                If Not cmtCell3.Comment Is Nothing Then cmtCell3.Comment.Delete
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell4 As Range
                Set cmtCell4 = ws.Range(parts(1))
                If Not cmtCell4.Comment Is Nothing Then cmtCell4.Comment.Visible = True
                ExecuteSingleCommand = True
            End If
            
        Case "HIDE_COMMENT"
            If UBound(parts) >= 1 Then
                Dim cmtCell5 As Range
                Set cmtCell5 = ws.Range(parts(1))
                If Not cmtCell5.Comment Is Nothing Then cmtCell5.Comment.Visible = False
                ExecuteSingleCommand = True
            End If
            
        Case "SHOW_ALL_COMMENTS"
            Dim cmt As Comment
            For Each cmt In ws.Comments
                cmt.Visible = True
            Next cmt
            ExecuteSingleCommand = True
            
        Case "HIDE_ALL_COMMENTS"
            Dim cmt2 As Comment
            For Each cmt2 In ws.Comments
                cmt2.Visible = False
            Next cmt2
            ExecuteSingleCommand = True
        
        ' ========== HYPERLINKS ==========
            
        Case "ADD_HYPERLINK"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_HYPERLINK_CELL"
            If UBound(parts) >= 3 Then
                ws.Hyperlinks.Add Anchor:=ws.Range(parts(1)), address:="", SubAddress:=parts(2), TextToDisplay:=parts(3)
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_HYPERLINK"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Hyperlinks.Delete
                ExecuteSingleCommand = True
            End If
        
        ' ========== PROTECTION ==========
            
        Case "LOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = True
                ExecuteSingleCommand = True
            End If
            
        Case "UNLOCK_CELLS"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Locked = False
                ExecuteSingleCommand = True
            End If
        
        ' ========== VIEWING AREA ==========
            
        Case "FREEZE_PANES"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ActiveWindow.FreezePanes = True
                ExecuteSingleCommand = True
            End If
            
        Case "FREEZE_TOP_ROW"
            ws.Range("A2").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True
            
        Case "FREEZE_FIRST_COLUMN"
            ws.Range("B1").Select
            ActiveWindow.FreezePanes = True
            ExecuteSingleCommand = True
            
        Case "UNFREEZE_PANES"
            ActiveWindow.FreezePanes = False
            ExecuteSingleCommand = True
            
        Case "ZOOM"
            If UBound(parts) >= 1 Then
                ActiveWindow.Zoom = CLng(parts(1))
                ExecuteSingleCommand = True
            End If
            
        Case "GOTO"
            If UBound(parts) >= 1 Then
                Application.Goto Reference:=ws.Range(parts(1)), Scroll:=True
                ExecuteSingleCommand = True
            End If
            
        Case "SELECT"
            If UBound(parts) >= 1 Then
                ws.Range(parts(1)).Select
                ExecuteSingleCommand = True
            End If
        
        ' ========== PRINT ==========
            
        Case "SET_PRINT_AREA"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintArea = parts(1)
                ExecuteSingleCommand = True
            End If
            
        Case "CLEAR_PRINT_AREA"
            ws.PageSetup.PrintArea = ""
            ExecuteSingleCommand = True
            
        Case "PAGE_ORIENTATION"
            If UBound(parts) >= 1 Then
                If UCase(Trim(parts(1))) = "LANDSCAPE" Then
                    ws.PageSetup.Orientation = xlLandscape
                Else
                    ws.PageSetup.Orientation = xlPortrait
                End If
                ExecuteSingleCommand = True
            End If
            
        Case "PAGE_MARGINS"
            If UBound(parts) >= 4 Then
                With ws.PageSetup
                    .LeftMargin = Application.CentimetersToPoints(CDbl(parts(1)))
                    .RightMargin = Application.CentimetersToPoints(CDbl(parts(2)))
                    .TopMargin = Application.CentimetersToPoints(CDbl(parts(3)))
                    .BottomMargin = Application.CentimetersToPoints(CDbl(parts(4)))
                End With
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_TITLES_ROWS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleRows = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_TITLES_COLS"
            If UBound(parts) >= 2 Then
                ws.PageSetup.PrintTitleColumns = "$" & parts(1) & ":$" & parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "PRINT_GRIDLINES"
            If UBound(parts) >= 1 Then
                ws.PageSetup.PrintGridlines = (UCase(Trim(parts(1))) = "TRUE")
                ExecuteSingleCommand = True
            End If
            
        Case "FIT_TO_PAGE"
            If UBound(parts) >= 2 Then
                With ws.PageSetup
                    .Zoom = False
                    .FitToPagesWide = CLng(parts(1))
                    .FitToPagesTall = CLng(parts(2))
                End With
                ExecuteSingleCommand = True
            End If
        
        ' ========== IMAGES ==========
            
        Case "INSERT_PICTURE"
            If UBound(parts) >= 5 Then
                Dim pic As Object
                Set pic = ws.Shapes.AddPicture(parts(1), msoFalse, msoTrue, _
                    CLng(parts(2)), CLng(parts(3)), CLng(parts(4)), CLng(parts(5)))
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_PICTURES"
            Dim shp As Shape
            For Each shp In ws.Shapes
                If shp.Type = msoPicture Then shp.Delete
            Next shp
            ExecuteSingleCommand = True
        
        ' ========== FORMS ==========
            
        Case "ADD_BUTTON"
            If UBound(parts) >= 5 Then
                Dim btn As Object
                Set btn = ws.Buttons.Add(CLng(parts(1)), CLng(parts(2)), CLng(parts(3)), CLng(parts(4)))
                btn.Caption = parts(5)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_CHECKBOX"
            If UBound(parts) >= 2 Then
                Dim chk As Object
                Dim chkCell As Range
                Set chkCell = ws.Range(parts(1))
                Set chk = ws.CheckBoxes.Add(chkCell.Left, chkCell.Top, 100, 15)
                chk.Caption = parts(2)
                ExecuteSingleCommand = True
            End If
            
        Case "ADD_DROPDOWN"
            If UBound(parts) >= 2 Then
                Dim dd As Object
                Dim ddCell As Range
                Set ddCell = ws.Range(parts(1))
                Set dd = ws.DropDowns.Add(ddCell.Left, ddCell.Top, 100, 15)
                dd.List = Split(parts(2), ";")
                ExecuteSingleCommand = True
            End If
            
        Case "DELETE_SHAPES"
            Dim shp2 As Shape
            For Each shp2 In ws.Shapes
                shp2.Delete
            Next shp2
            ExecuteSingleCommand = True
        
        ' ========== SPECIAL ==========
            
        Case "CALCULATE"
            Application.Calculate
            ExecuteSingleCommand = True
            
        Case "CALCULATE_SHEET"
            ws.Calculate
            ExecuteSingleCommand = True
            
        Case "TEXT_TO_COLUMNS"
            If UBound(parts) >= 2 Then
                Dim ttcRng As Range
                Dim delim As String
                Set ttcRng = ws.Range(parts(1))
                delim = parts(2)
                
                Dim delimTab As Boolean, delimSemi As Boolean, delimComma As Boolean, delimSpace As Boolean, delimOther As Boolean
                Dim otherChar As String
                
                Select Case UCase(delim)
                    Case "TAB": delimTab = True
                    Case "SEMICOLON", ";": delimSemi = True
                    Case "COMMA", ",": delimComma = True
                    Case "SPACE", " ": delimSpace = True
                    Case Else
                        delimOther = True
                        otherChar = delim
                End Select
                
                ttcRng.TextToColumns Destination:=ttcRng, DataType:=xlDelimited, _
                    Tab:=delimTab, Semicolon:=delimSemi, Comma:=delimComma, _
                    Space:=delimSpace, Other:=delimOther, otherChar:=otherChar
                    
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_SPACES"
            If UBound(parts) >= 1 Then
                Dim spRng As Range, spCell As Range
                Set spRng = ws.Range(parts(1))
                For Each spCell In spRng
                    If Not IsEmpty(spCell.value) Then
                        spCell.value = Application.WorksheetFunction.Trim(spCell.value)
                    End If
                Next spCell
                ExecuteSingleCommand = True
            End If
            
        Case "UPPER_CASE"
            If UBound(parts) >= 1 Then
                Dim ucRng As Range, ucCell As Range
                Set ucRng = ws.Range(parts(1))
                For Each ucCell In ucRng
                    If Not IsEmpty(ucCell.value) Then
                        ucCell.value = UCase(ucCell.value)
                    End If
                Next ucCell
                ExecuteSingleCommand = True
            End If
            
        Case "LOWER_CASE"
            If UBound(parts) >= 1 Then
                Dim lcRng As Range, lcCell As Range
                Set lcRng = ws.Range(parts(1))
                For Each lcCell In lcRng
                    If Not IsEmpty(lcCell.value) Then
                        lcCell.value = LCase(lcCell.value)
                    End If
                Next lcCell
                ExecuteSingleCommand = True
            End If
            
        Case "PROPER_CASE"
            If UBound(parts) >= 1 Then
                Dim pcRng As Range, pcCell As Range
                Set pcRng = ws.Range(parts(1))
                For Each pcCell In pcRng
                    If Not IsEmpty(pcCell.value) Then
                        pcCell.value = Application.WorksheetFunction.Proper(pcCell.value)
                    End If
                Next pcCell
                ExecuteSingleCommand = True
            End If
            
        Case "FLASH_FILL"
            If UBound(parts) >= 1 Then
                On Error Resume Next
                ws.Range(parts(1)).FlashFill
                On Error GoTo ErrorHandler
                ExecuteSingleCommand = True
            End If
            
        Case "SUBTOTAL"
            If UBound(parts) >= 3 Then
                Dim stRng As Range
                Dim stFunc As Long
                Dim stCol As Long
                Set stRng = ws.Range(parts(1))
                
                Select Case UCase(Trim(parts(2)))
                    Case "SUM": stFunc = xlSum
                    Case "COUNT": stFunc = xlCount
                    Case "AVERAGE": stFunc = xlAverage
                    Case "MAX": stFunc = xlMax
                    Case "MIN": stFunc = xlMin
                    Case Else: stFunc = xlSum
                End Select
                
                stCol = GetColumnNumber(parts(3), stRng, ws)
                stRng.Subtotal GroupBy:=1, Function:=stFunc, TotalList:=Array(stCol)
                ExecuteSingleCommand = True
            End If
            
        Case "REMOVE_SUBTOTALS"
            ws.UsedRange.RemoveSubtotal
            ExecuteSingleCommand = True
            
        Case Else
            ExecuteSingleCommand = False
    End Select
    
    Exit Function
    
ErrorHandler:
    Debug.Print "ExecuteSingleCommand Error: " & Err.Number & " - " & Err.Description & " | CMD: " & cmd
    ExecuteSingleCommand = False
End Function


