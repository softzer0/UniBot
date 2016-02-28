VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMathParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'* CLASS: clsMathParser                                                 v.3.3.0 *
'*                                                        Leonardo Volpi        *
'*                                                        Michael Ruder         *
'*                                                        Thomas Zeutschler     *
'*                                                        Lieven Dossche        *
'*                                                        Arnaud d.Grammont     *
'*  Math-Physical Expression Evaluation                                         *
'*  for VB 6, VBA 97/2000                                                       *
'********************************************************************************
Option Explicit
'-------------------------------------------------------------------------------
#Const DEBUG_MODE = 0
'-------------------------------------------------------------------------------
' CONSTANTS
'-------------------------------------------------------------------------------
Const HiOPER     As Long = 100
Const HiVT       As Long = 100
Const HiET       As Long = 100
Const HiARG      As Long = 2
Const PI_        As Double = 3.14159265358979
Const Fun1V As String = "Abs Atn Cos Exp Fix Int Ln Log Rnd Sgn Sin Sqr Cbr Tan Acos Asin " & _
    "Cosh Sinh Tanh Acosh Asinh Atanh Fact Not Erf Gamma Gammaln Digamma Zeta Ei " & _
    "csc sec cot acsc asec acot csch sech coth acsch asech acoth Dec Rad Deg Grad"
Const Fun2V As String = "Comb Max Min Mcm Mcd Lcm Gcd Mod And Or Xor Beta Root Round Nand Nor NXor"
'-------------------------------------------------------------------------------
' ENUMERATIONS
'-------------------------------------------------------------------------------
Const symARGUMENT = -1      'An Argument
Const symPlus = 0           '"+"
Const symMinus = 1          '"-"
Const symMul = 2            '"*"
Const symDiv = 3            '"/"
Const symPercent = 4        '% percentage
Const symDivInt = 5         '"\" 'integer division, added MR 20-06-02
Const symPov = 6            '"^"
Const symAbs = 7            '"abs", "|.|"
Const symAtn = 8            '"atn"
Const symCos = 9            '"cos"
Const symSin = 11           '"sin"
Const symExp = 12           '"exp"
Const symFix = 13           '"fix"
Const symInt = 14           '"int"
 Const symLn = 15           '"ln"
 Const symLog = 16          '"log"
Const symRnd = 17           '"rnd"
Const symSgn = 18           '"sgn"
Const symSqr = 19           '"sqr"
Const symTan = 20           '"tan"
Const symAcos = 21          '"acos"
Const symAsin = 22          '"asin"
Const symCosh = 23          '"cosh"
Const symSinh = 24          '"sinh"
Const symTanh = 25          '"tanh"
Const symAcosh = 26         '"acosh"
Const symAsinh = 27         '"asinh"
Const symAtanh = 28         '"atanh"
Const symmod = 29           '"mod"
Const symFact = 30          '"fact", "!"
Const symComb = 31          '"comb"
Const symMin = 32           '"min"
Const symMax = 33           '"max"
Const symMcd = 34           '"mcd"
Const symMcm = 35           '"mcm"
Const symGT = 36            '">"
Const symGE = 37            '">="
Const symLT = 38            '"<"
Const symLE = 39            '"<="
Const symEQ = 40            '"="
Const symNE = 41            '"<>"
Const symAnd = 42           '"and"
Const symOr = 43            '"or"
Const symNot = 44           '"not"
Const symXor = 45           '"xor"
Const symErf = 46           '"erf"
Const symGamma = 47         '"gamma"
Const symGammaln = 48       '"gammaln"
Const symDigamma = 49       '"digamma"
Const symBeta = 50          '"beta"
Const symZeta = 51          '"zeta"
Const symEi = 52            '"ei"
Const symCsc = 53           '"csc  cosecant"
Const symSec = 54           '"sec  secant"
Const symCot = 55           '"cot  cotangent"
Const symACsc = 56          '"acsc  inverse cosecant"
Const symASec = 57          '"asec  inverse secant"
Const symACot = 58          '"acot  inverse cotangent"
Const symCsch = 59          '"csch  hyperbolic cosecant"
Const symSech = 60          '"sech  hyperbolic secant"
Const symCoth = 61          '"coth  hyperbolic cotangent"
Const symACsch = 62         '"acsch inverse hyperbolic cosecant"
Const symASech = 63         '"asech inverse hyperbolic secant"
Const symACoth = 64         '"acoth inverse hyperbolic cotangent"
Const symCbr = 65           '"cbr cube root"
Const symRoot = 66          '"root n-th root"
Const symDec = 67           '"dec  decimal part"
Const symRad = 68           '"rad  convert radiant to current angle unit"
Const symDeg = 69           '"deg  convert degree 360 to current angle unit"
Const symRound = 70         '"round"
Const symGrad = 71          '"grad  convert degree 400 to current angle unit"
Const symNAnd = 72          '"nand"
Const symNOr = 73           '"nor"
Const symNXor = 74          '"NXor"

'-------------------------------------------------------------------------------
' TYPE DECLARATIONS
'-------------------------------------------------------------------------------
Private Type T_VTREC           'Variable Table Record
  Idx              As Long
  Nome             As String
  Value            As Double
End Type

Private Type T_ETREC           'Expression Table record
  Fun              As String
  FunTok           As Long
  Arg(1 To HiARG)  As T_VTREC
  ArgTop           As Long
  ArgOf            As Long
  ArgIdx           As Long
  Value            As Double
  PosInExpr        As Long
  PriLvl           As Long
  PriIdx           As Long
End Type
'-------------------------------------------------------------------------------
' LOCALS
'-------------------------------------------------------------------------------
Dim Expr            As String
Dim ExprOK          As Boolean  'expression OK
Dim ExprNoOK        As Boolean  'expr OK, no variables
Dim Expr1OK         As Boolean  'expr OK, exact 1 var
Dim VT()            As T_VTREC
Dim ET()            As T_ETREC
Dim VTtop           As Long
Dim ETtop           As Long
Dim ErrMsg          As String   'error message  11-11-02 VL
Dim Angle           As String   'RAD GRAD DEG
Dim DecRound        As Integer
Dim CvAngleCoeff    As Double

'-------------------------------------------------------------------------------
' PUBLIC FUCTIONS
'-------------------------------------------------------------------------------
' store expression as array of records; check syntax
Public Function StoreExpression(ByVal strExpr As String) As Boolean
  Expr = Trim$(strExpr)
  ExprOK = Parse(Expr)
  ExprNoOK = ExprOK And (VTtop = 0)
  Expr1OK = ExprOK And (VTtop = 1)
  StoreExpression = ExprOK
End Function
'-------------------------------------------------------------------------------
' get the expression
Public Property Get Expression() As String
  Expression = Expr
End Property
'-------------------------------------------------------------------------------
' get the top of the var array (=N-1 bacause starts on 0)
Public Property Get VarTop() As Long
  VarTop = VTtop
End Property
'-------------------------------------------------------------------------------
' get name of a variable. VL
Public Property Get VarName(ByVal Index As Long) As String
  If Index <= VTtop Then
    VarName = VT(Index).Nome
  End If
End Property
'-------------------------------------------------------------------------------
' get value assigned to a variable
Public Property Get VarValue(ByVal Index As Long) As Double
  If Index <= VTtop Then
    VarValue = VT(Index).Value
  End If
End Property
'-------------------------------------------------------------------------------
' assign a value to a certain variable
Public Property Let VarValue(ByVal Index As Long, ByVal VarVal As Double)
  If Index <= VTtop Then
    VT(Index).Value = VarVal
  End If
End Property
'-------------------------------------------------------------------------------
' get current setting for angle computing (RAD (default), DEG or GRAD)
Public Property Get AngleUnit() As String
    If Angle = vbNullString Then Angle = "RAD"
    AngleUnit = Angle
End Property
'-------------------------------------------------------------------------------
' set the unit of measure for angle computing (RAD (default), DEG or GRAD)
Public Property Let AngleUnit(ByVal AngleUM As String)
    Select Case UCase$(AngleUM)
        Case "DEG"
            Angle = "DEG"
            CvAngleCoeff = PI_ / 180
        Case "GRAD"
            Angle = "GRAD"
            CvAngleCoeff = PI_ / 200
        Case Else
            Angle = "RAD"
            CvAngleCoeff = 1
    End Select
End Property
'-------------------------------------------------------------------------------
' get the error message  11-11-02 VL
Public Property Get ErrorDescription() As String
  ErrorDescription = ErrMsg
End Property
'-------------------------------------------------------------------------------
' evaluate expression
Public Function Eval() As Double
Dim ExprVal As Double

If Not ExprOK Then GoTo Error_Handler
If VTtop > 0 Then SubstVars
If Not Eval_(ExprVal) Then GoTo Error_Handler
Eval = ExprVal
Exit Function
'
Error_Handler:
On Error Resume Next
Err.Raise 1001, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' evaluate an expression with exactly 1 var
Public Function Eval1(ByVal x As Double) As Double
Dim i  As Long
Dim J  As Long
Dim Id As Long
Dim ExprVal As Double
'
If VTtop > 1 Then
    ErrMsg = "too many variables"
    GoTo Error_Handler
End If

For i = 1 To ETtop
  For J = 1 To HiARG
    Id = ET(i).Arg(J).Idx
    If Id <> 0 Then
      ET(i).Arg(J).Value = x
    End If
  Next
Next
If Not Eval_(ExprVal) Then GoTo Error_Handler
Eval1 = ExprVal
Exit Function
'
Error_Handler:
On Error Resume Next
Err.Raise 1002, "MathParser", ErrMsg
End Function
'-------------------------------------------------------------------------------
' Math Parser Routine
'         rev 30-08-02 Leonardo Volpi;  rev 20-10-02 L.Dos
Private Function Parse(ByVal strExpr As String) As Boolean
  Dim lExpr          As String
  Dim char           As String '*1
  Dim SubExpr        As String
  Dim lenExpr        As Long
  Dim FunName        As String
  Dim GetNextArg     As Boolean
  Dim SaveArg        As String
  Dim Npart          As Long
  Dim Nabs           As Long
  Dim Idx            As Long
  Dim retval         As Double
  Dim arrPriLvl()    As Long
  Dim srtLo          As Long      'vars for sort algoritme
  Dim srtHi          As Long
  Dim Tmp            As Long
  Dim flag_exchanged As Boolean
  Dim i              As Long
  Dim J              As Long
  Dim j1             As Long
  Dim j2             As Long
  Dim LogicSymb      As String
  Dim PowerSign      As Boolean
  Dim LastArg        As Boolean
  
  ReDim ET(HiET)
  ReDim VT(HiVT)
  ETtop = 0
  VTtop = 0
  Angle = "RAD"
  DecRound = 15
  CvAngleCoeff = 1
  ErrMsg = vbNullString   'VL
  lExpr = Trim$(strExpr)
  PowerSign = False      'fix power sign bug. ex: 3^-2 - VL 2.6.2003
  '***** abs |.| function counter
  i = NabsCount(lExpr)
  Nabs = i / 2
  If (2 * Nabs <> i) Then
    ErrMsg = "abs symbols |.| mismatch"  'VL
    GoTo ErrHandler
  End If
  '***** begin parse process
  lenExpr = Len(lExpr)
  For i = 1 To lenExpr
    char = Mid$(lExpr, i, 1)
    Select Case char
      Case " "                                    '***** skip spaces
      Case "(", "[", "{"                          '***** open parentheses
        Npart = Npart + 1                         'inc # open parentheses
        If SubExpr <> vbNullString Then                     'eval preceding text
          If InList(SubExpr, Fun1V) Then          'monovariable function
            ETtop = ETtop + 1                     '   store in ET
            With ET(ETtop)
              .PosInExpr = i                      'position in expr
              .Fun = SubExpr                      'function name
              .FunTok = GetFunTok(SubExpr)        'function Token (enum)
              .PriLvl = Npart * 10                'priority level=open parenth*10
              .ArgTop = 1                         'ntal Args=1
            End With
          ElseIf InList(SubExpr, Fun2V) Then      'bivariable function
            FunName = SubExpr                     '   store only the name
          Else                                    'unknown expression preceding parenth
            ErrMsg = "Function <" + SubExpr + "> unknown at pos " + Str(i)
            GoTo ErrHandler
          End If
          SubExpr = vbNullString                            'start parsing for new subexpr
        End If
      Case ")", "]", "}"                          '***** open parentheses
        Npart = Npart - 1                         'dec # open parentheses
        If Npart < 0 Then                         'want to close to many brackets
          ErrMsg = "Too many closing brackets at pos " + Str(i)
          GoTo ErrHandler
        End If
      Case "+", "-"                               '*****
        'check the exponential sign (preceding was eg 1.23E of 1.23E-2)
        If CheckExpo(SubExpr) Then             'fix bug 18-1-03  thanks to Michael Ruder
          SubExpr = SubExpr + char                  'continue parsing number
        ElseIf PowerSign Then
          SubExpr = SubExpr + char    'fix power sign bug. ex: 3^-2 - VL 2.6.2003  ' thanks to Javier M. Montalban 14.2.03
        Else
          ETtop = ETtop + 1                       'store in ET
          With ET(ETtop)                          '
            .PosInExpr = i
            .Fun = char
            .FunTok = GetFunTok(char)
            .PriLvl = 2 + Npart * 10
            .ArgTop = 2                           'two arguments
          End With
          If SubExpr = vbNullString Then                    'if nothing precedes: implicit 0
            SubExpr = "0"
          End If
          GetNextArg = True                       'get second argument
        End If
      Case "*", "/", "\"                      '*****
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = char
          .FunTok = GetFunTok(char)
          .PriLvl = 3 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
      Case "^"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "^"
          .FunTok = GetFunTok(char)
          .PriLvl = 4 + Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        PowerSign = True    'fix power sign bug. ex: 3^-2 - VL 2.6.2003
      Case "!"
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "!"
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        GetNextArg = True
        SaveArg = SubExpr
      Case "%"  'percentage
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = "%"
          .FunTok = GetFunTok(char)
          .PriLvl = 9 + Npart * 10
          .ArgTop = 1                             'one argument
        End With
        GetNextArg = True
        SaveArg = SubExpr
      Case ";"                                    'comes from bivariate function f(x;y)
        If FunName = vbNullString Then
          ErrMsg = "syntax error at pos:" + Str(i)
          GoTo ErrHandler
        End If
        ETtop = ETtop + 1
        With ET(ETtop)
          .PosInExpr = i
          .Fun = FunName                          'previous stored
          .FunTok = GetFunTok(FunName)
          .PriLvl = Npart * 10
          .ArgTop = 2                             'two arguments
        End With
        GetNextArg = True
        FunName = vbNullString                              'reset function
      Case "|"                                    '***** absolute symbol |.|
        If SubExpr = vbNullString Then
          Npart = Npart + 1                       'increment brackets PriLvl
          ETtop = ETtop + 1
          With ET(ETtop)
            .PosInExpr = i
            .Fun = "abs"                          'symbols |.| is similar to  abs(.)
            .FunTok = GetFunTok("abs")
            .PriLvl = Npart * 10
            .ArgTop = 1                           'one argument
          End With
        Else
          Npart = Npart - 1
          If Npart < 0 Then                       'too many closing brackets
            ErrMsg = "Too many closing brackets at pos " + Str(i) ' MR 16-01-03
            GoTo ErrHandler
          End If
        End If
      Case "=", "<", ">"                          'Logical operators
        If LogicSymb = vbNullString Then
            ETtop = ETtop + 1
            GetNextArg = True
        End If
        LogicSymb = LogicSymb + char
        With ET(ETtop)
          .PosInExpr = i
          .Fun = LogicSymb                        'logic symbol
          .FunTok = GetFunTok(LogicSymb)
          .PriLvl = 1 + Npart * 10
          .ArgTop = 2                             'two argument
        End With
      Case "x", "y", "z", "X", "Y", "Z"           ''monomial coeff.
        If IsNumeric_(SubExpr) Then               'fix 2.3.2003 thanks to Michael Ruder
            ETtop = ETtop + 1                     'Ex: 7x  is converted into product 7*x
            With ET(ETtop)
              .PosInExpr = i
              .Fun = "*"
              .FunTok = GetFunTok("*")
              .PriLvl = 3 + Npart * 10
              .ArgTop = 2                             'two argument
            End With
            GetNextArg = True
            i = i - 1  'one step back
        Else
            SubExpr = SubExpr + char
        End If
      Case Else                                   '***** continue parsing
        SubExpr = SubExpr & char
    End Select
    
    If char <> " " And char <> "^" And PowerSign = True Then PowerSign = False  'reset powersign flag
    If GetNextArg Then                            'search for next argument
      If SubExpr = vbNullString Then                        'no next argument found
        ErrMsg = "missing argument"
        GoTo ErrHandler
      End If
        J = 1
        LastArg = False
        GoSub Catch_Argument
      SubExpr = vbNullString
      GetNextArg = False
      If ET(ETtop).Fun = "!" Or ET(ETtop).Fun = "%" Then  'restore the previous argument for "!" and "%" operator
        SubExpr = SaveArg
      End If
    Else
      LogicSymb = vbNullString
    End If
  Next
  
  If Npart > 0 Then                               'parentheses
    ErrMsg = "Not enough closing brackets"
    GoTo ErrHandler
  End If
  If ETtop < 1 Then                               'no operation detected
    ETtop = 1
    With ET(ETtop)
      .PosInExpr = 1
      .Fun = "+"
      .FunTok = GetFunTok("+")
      .PriLvl = 1
      .ArgTop = 2
    End With
  End If
  For i = 1 To ETtop                              'init 2e argument
    ET(i).Arg(ET(i).ArgTop) = ET(i + 1).Arg(1)
  Next
  
  If SubExpr <> vbNullString Then                           'catch last argument or Vars
    J = ET(ETtop).ArgTop
    LastArg = True
    GoSub Catch_Argument
  Else
    'bug 7.10.03 last argument missing 3+ or sin() ...  thanks to Rodigro Farinha
    ErrMsg = "missing argument"
    GoTo ErrHandler
  End If
  
  If ETtop > 0 Then
    ReDim Preserve ET(ETtop)
  Else
    ReDim Preserve ET(0)
  End If
  If VTtop > 0 Then
    ReDim Preserve VT(VTtop)
  Else
    ReDim Preserve VT(0)
  End If
  
  ReDim arrPriLvl(ETtop)                     'create array with priority levels
  For i = 1 To ETtop                              'and copy then from main array
    arrPriLvl(i) = ET(i).PriLvl
  Next
  For i = 1 To ETtop                              'fill sort order default 0 to ETtop
    ET(i).PriIdx = i
  Next
  srtLo = 1                                       '***** start sort algorithm
  srtHi = ETtop - 1
  Do
    flag_exchanged = False
    For i = srtLo To srtHi Step 2
      J = i + 1
      If arrPriLvl(i) < arrPriLvl(J) Then
        Tmp = arrPriLvl(J)
        arrPriLvl(J) = arrPriLvl(i)
        arrPriLvl(i) = Tmp
        Tmp = ET(J).PriIdx
        ET(J).PriIdx = ET(i).PriIdx
        ET(i).PriIdx = Tmp
        flag_exchanged = True
      End If
    Next
    If srtLo = 1 Then
      srtLo = 2
    Else
      srtLo = 1
    End If
  Loop Until (srtLo = 1) And Not flag_exchanged
  
  For i = 1 To ETtop                              'build relations
    J = ET(i).PriIdx
    j1 = J - 1
    Do While j1 >= 0
      If ET(j1).ArgOf = 0 Then
        Exit Do
      End If
      j1 = j1 - 1
    Loop
    j2 = J + 1
    Do While j2 <= ETtop
      If ET(j2).ArgOf = 0 Then
        Exit Do
      End If
      j2 = j2 + 1
    Loop
    If (j1 < 1) And (j2 <= ETtop) Then            '
      ET(J).ArgOf = j2
      ET(J).ArgIdx = 1
    ElseIf (j1 > 0) And (j2 > ETtop) Then        '
      ET(J).ArgOf = j1
      ET(J).ArgIdx = ET(j1).ArgTop
    ElseIf (j1 > 0) And (j2 <= ETtop) Then       '
      If (ET(j1).PriLvl) >= (ET(j2).PriLvl) Then  'take that one with the upper PriLvl
        ET(J).ArgOf = j1
        ET(J).ArgIdx = ET(j1).ArgTop
      Else                                        '
        ET(J).ArgOf = j2
        ET(J).ArgIdx = 1
      End If
    Else
      Exit For
    End If
  Next
  
  For i = 1 To ETtop                              'eliminate dependent arguments
    J = ET(i).ArgOf
    If J > 0 Then
      With ET(J).Arg(ET(i).ArgIdx)
       .Idx = 0
       .Nome = vbNullString
      End With
    End If
  Next
                            
#If DEBUG_MODE > 0 Then
  ET_Dump 1
#End If
                            
  Parse = True
  Exit Function
'internal routines -----------------------
Catch_Argument:
    If convSymbConst(SubExpr, retval) Then      'check if argument is a symbolic constant
      ET(ETtop).Arg(J).Value = retval
    ElseIf convEGU(SubExpr, retval) Then        'check if argument is Eng Units
      ET(ETtop).Arg(J).Value = retval
    ElseIf IsNumeric_(SubExpr) Then             'check if argument is number
      ET(ETtop).Arg(J).Value = Val(SubExpr)
    ElseIf cvDegree(SubExpr, retval) Then       'check if argument is ddmmss format degree
      ET(ETtop).Arg(J).Value = retval
    Else
      If Not IsLetter(Left$(SubExpr, 1)) Then
        ErrMsg = "variable name must start with a letter: " + SubExpr
        GoTo ErrHandler
      End If
      StoreVar SubExpr, LastArg
      If VTtop > HiVT Then
        ErrMsg = "too many Vars"
        GoTo ErrHandler
      End If
    End If
Return

ErrHandler:
  ETtop = ETtop
  Parse = False
End Function
'---------------------------------------------------------------------------------
Private Function CheckExpo(ByVal SubExpr As String) As Boolean
Dim s_1 As String, s_2 As String, ls As Long
'detect if SubExpr is the mantissa of an expo format number 1.2E+01 , 4E-12, 1.0E-6
    CheckExpo = False
    ls = Len(SubExpr)
    If ls < 2 Then Exit Function
    s_1 = Right$(SubExpr, 1)
    s_2 = Left$(SubExpr, ls - 1)
    If (UCase$(s_1) = "E") And IsNumeric_(s_2) Then CheckExpo = True
End Function
'-------------------------------------------------------------------------------
'[modified 10/02 by Thomas Zeutschler]
Private Function Eval_(EvalValue As Double) As Boolean
  Dim a     As Double
  Dim b     As Double
  Dim ris   As Double
  Dim J     As Long
  Dim k     As Long
  Dim pos   As Long
  Dim m     As Long
  Dim N     As Long
  
  On Error GoTo ErrHandler  '<<< comment for debug  VL 30-8-02
  For J = 1 To ETtop    'Evaluation procedure begins
    k = ET(J).PriIdx
    With ET(k)
      a = .Arg(1).Value
      b = .Arg(2).Value
      Select Case .FunTok
        Case symPlus:            ris = a + b
        Case symMinus:           ris = a - b
        Case symMul:             ris = a * b
        Case symDiv:             ris = a / b
        Case symPercent:         ris = percent(k)
        Case symDivInt:          ris = a \ b
        Case symPov:             ris = a ^ b
        Case symAbs:             ris = Abs(a)
        Case symAtn:             ris = Atn(a) / CvAngleCoeff
        Case symCos:             ris = Cos(CvAngleCoeff * a)
        Case symSin:             ris = Sin(CvAngleCoeff * a)
        Case symExp:             ris = Exp(a)
        Case symFix:             ris = Fix(a)
        Case symInt:             ris = Int(a)
        Case symDec:             ris = a - Fix(a)
        Case symLn:              ris = Log(a)
        Case symLog:             ris = Log(a) / Log(10)
        Case symRnd:             ris = a * Rnd(1)
        Case symSgn:             ris = Sgn(a)
        Case symSqr:             ris = Sqr(a)
        Case symCbr:             ris = Sgn(a) * Abs(a) ^ (1 / 3)
        Case symTan:             ris = Tan(CvAngleCoeff * a)
        Case symAcos:            ris = Acos_(a) / CvAngleCoeff
        Case symAsin:            ris = Asin_(a) / CvAngleCoeff
        Case symCosh:            ris = Cosh_(a)
        Case symSinh:            ris = Sinh_(a)
        Case symTanh:            ris = Tanh_(a)
        Case symAcosh:           ris = Acosh_(a)   '7.10.2003 fix bug (thank to Rodrigo Farinha)
        Case symAsinh:           ris = Asinh_(a)
        Case symAtanh:           ris = Atanh_(a)
        Case symRoot:            ris = MiRoot_(a, b)
        Case symmod:             ris = a Mod b
        Case symFact:            ris = fact(a)
        Case symComb:            ris = Comb(a, b)
        Case symMin:             ris = min_(a, b)
        Case symMax:             ris = max_(a, b)
        Case symMcd:             ris = MCD_(a, b)
        Case symMcm:             ris = MCM_(a, b)
        Case symGT:              ris = -CDbl((a > b))
        Case symGE:              ris = -CDbl(a >= b)
        Case symLT:              ris = -CDbl(a < b)
        Case symLE:              ris = -CDbl(a <= b)
        Case symEQ:              ris = -CDbl(a = b)
        Case symNE:              ris = -CDbl(a <> b)
        Case symAnd:             ris = -CDbl((a <> 0) And (b <> 0))
        Case symOr:              ris = -CDbl((a <> 0) Or (b <> 0))
        Case symNot:             ris = -CDbl(a = 0)
        Case symXor:             ris = -CDbl((a <> 0) Xor (b <> 0)) ' MR 16-01-03 XOR corrected
        Case symNAnd:            ris = -CDbl((a = 0) Or (b = 0))     '
        Case symNOr:             ris = -CDbl((a = 0) And (b = 0))    '
        Case symNXor:            ris = -CDbl((a <> 0) = (b <> 0))    'MR 16-01-03 NXor        '
        Case symErf:             ris = erf_(a)
        Case symGamma:           ris = gamma_(a)
        Case symGammaln:         ris = gammaln_(a)
        Case symDigamma:         ris = digamma_(a)
        Case symBeta:            ris = beta_(a, b)
        Case symZeta:            ris = Zeta_(a)
        Case symEi:              ris = exp_integr(a)
        Case symCsc:             ris = 1 / Sin(CvAngleCoeff * a)
        Case symSec:             ris = 1 / Cos(CvAngleCoeff * a)
        Case symCot:             ris = 1 / Tan(CvAngleCoeff * a)
        Case symACsc:            ris = Asin_(1 / a) / CvAngleCoeff
        Case symASec:            ris = Acos_(1 / a) / CvAngleCoeff
        Case symACot:            ris = PI_ / 2 - Atn(a) / CvAngleCoeff
        Case symCsch:            ris = 1 / Sinh_(a)
        Case symSech:            ris = 1 / Cosh_(a)
        Case symCoth:            ris = 1 / Tanh_(a)
        Case symACsch:           ris = Asinh_(1 / a)
        Case symASech:           ris = Acosh_(1 / a)
        Case symACoth:           ris = Atanh_(1 / a)
        Case symRad:             ris = a / CvAngleCoeff
        Case symDeg:             ris = a / CvAngleCoeff * PI_ / 180
        Case symGrad:            ris = a / CvAngleCoeff * PI_ / 200
        Case symRound:           ris = round_(a, b)
        Case Else
          ErrMsg = "Function <" & CStr(.FunTok) & "> missing?"  'VL
          GoTo ErrHandler
      End Select
      .Value = ris
      m = .ArgOf
      N = .ArgIdx
      If m = 0 Or N = 0 Then Exit For
      ET(m).Arg(N).Value = ris
    End With
  Next
  EvalValue = ET(k).Value
  Eval_ = True
  Exit Function
ErrHandler:
  If ET(k).ArgTop = 1 Then
      ErrMsg = "Evaluation error <" + ET(k).Fun + "(" + CStr(a) + ")" + "> at pos:" + Str(ET(k).PosInExpr)
  Else
      If InList(ET(k).Fun, Fun2V) Then
        ErrMsg = "Evaluation error <" + ET(k).Fun + "(" + CStr(a) + " , " + CStr(b) + ")" + "> at pos:" + Str(ET(k).PosInExpr)
      Else
        ErrMsg = "Evaluation error <" + CStr(a) + " " + ET(k).Fun + " " + CStr(b) + "> at pos:" + Str(ET(k).PosInExpr)
      End If
  End If
  EvalValue = 0
  Eval_ = False
End Function
'-------------------------------------------------------------------------------
' Assignes a value to symbolic Vars
Private Sub SubstVars()
  Dim i  As Long
  Dim J  As Long
  Dim Id As Long

  For i = 1 To ETtop
    For J = 1 To HiARG
      Id = ET(i).Arg(J).Idx
      If Id <> 0 Then
        ET(i).Arg(J).Value = VT(Id).Value
      End If
    Next
  Next
End Sub
'-------------------------------------------------------------------------------
' search if var already exists in table, if not add it
Private Sub StoreVar(ByVal SubExpr As String, ByVal LastArg As Boolean)
  Dim VTIdx  As Long
  Dim ArgIdx As Long
  Dim Found  As Boolean
    
  Found = False
  For VTIdx = 1 To VTtop
    If VT(VTIdx).Nome = SubExpr Then
      Found = True
      Exit For
    End If
  Next
  If Not Found Then
    VTtop = VTtop + 1     'new variable
    If VTtop > HiVT Then  'to many Vars
      Exit Sub
    End If
    VT(VTtop).Nome = SubExpr
  End If
  If LastArg Then
    ArgIdx = ET(ETtop).ArgTop
  Else
    ArgIdx = 1
  End If
  With ET(ETtop).Arg(ArgIdx)
    .Nome = SubExpr
    .Idx = VTIdx
  End With
End Sub
'-------------------------------------------------------------------------------
' get function token '[added 10/02 by Thomas Zeutschler]
Private Function GetFunTok(ByVal FunTok As String) As Long
  Select Case LCase$(FunTok)
    Case "+":       GetFunTok = symPlus
    Case "-":       GetFunTok = symMinus
    Case "*":       GetFunTok = symMul
    Case "/":       GetFunTok = symDiv
    Case "%":       GetFunTok = symPercent
    Case "\":       GetFunTok = symDivInt
    Case "^":       GetFunTok = symPov
    Case "abs":     GetFunTok = symAbs
    Case "atn":     GetFunTok = symAtn
    Case "cos":     GetFunTok = symCos
    Case "sin":     GetFunTok = symSin
    Case "exp":     GetFunTok = symExp
    Case "fix":     GetFunTok = symFix
    Case "int":     GetFunTok = symInt
    Case "dec":     GetFunTok = symDec
    Case "ln":      GetFunTok = symLn
    Case "log":     GetFunTok = symLog
    Case "rnd":     GetFunTok = symRnd
    Case "sgn":     GetFunTok = symSgn
    Case "sqr":     GetFunTok = symSqr
    Case "cbr":     GetFunTok = symCbr
    Case "tan":     GetFunTok = symTan
    Case "acos":    GetFunTok = symAcos
    Case "asin":    GetFunTok = symAsin
    Case "cosh":    GetFunTok = symCosh
    Case "sinh":    GetFunTok = symSinh
    Case "tanh":    GetFunTok = symTanh
    Case "acosh":   GetFunTok = symAcosh
    Case "asinh":   GetFunTok = symAsinh
    Case "atanh":   GetFunTok = symAtanh
    Case "root":    GetFunTok = symRoot
    Case "mod":     GetFunTok = symmod
    Case "fact", "!": GetFunTok = symFact
    Case "comb":    GetFunTok = symComb
    Case "min":     GetFunTok = symMin
    Case "max":     GetFunTok = symMax
    Case "mcd", "gcd":    GetFunTok = symMcd
    Case "mcm", "lcm":    GetFunTok = symMcm
    Case ">":       GetFunTok = symGT
    Case ">=", "=>": GetFunTok = symGE
    Case "<":       GetFunTok = symLT
    Case "<=", "=<": GetFunTok = symLE
    Case "=":       GetFunTok = symEQ
    Case "<>":      GetFunTok = symNE
    Case "and":     GetFunTok = symAnd
    Case "or":      GetFunTok = symOr
    Case "not":     GetFunTok = symNot
    Case "xor":     GetFunTok = symXor
    Case "nand":    GetFunTok = symNAnd
    Case "nor":     GetFunTok = symNOr
    Case "nxor":    GetFunTok = symNXor
    Case "erf":     GetFunTok = symErf
    Case "gamma":   GetFunTok = symGamma
    Case "gammaln": GetFunTok = symGammaln
    Case "digamma": GetFunTok = symDigamma
    Case "beta":    GetFunTok = symBeta
    Case "zeta":    GetFunTok = symZeta
    Case "ei":      GetFunTok = symEi
    Case "csc":     GetFunTok = symCsc
    Case "sec":     GetFunTok = symSec
    Case "cot":     GetFunTok = symCot
    Case "acsc":    GetFunTok = symACsc
    Case "asec":    GetFunTok = symASec
    Case "acot":    GetFunTok = symACot
    Case "csch":    GetFunTok = symCsch
    Case "sech":    GetFunTok = symSech
    Case "coth":    GetFunTok = symCoth
    Case "acsch":   GetFunTok = symACsch
    Case "asech":   GetFunTok = symASech
    Case "acoth":   GetFunTok = symACoth
    Case "rad":     GetFunTok = symRad
    Case "deg":     GetFunTok = symDeg
    Case "grad":     GetFunTok = symGrad
    Case "round":   GetFunTok = symRound
    Case Else
        GetFunTok = symARGUMENT
    End Select
End Function
'-------------------------------------------------------------------------------
' translate egu to multiplication factor
'  accepts a string contains a measure like "2ms" ,"234.12Mhz", "0.1uF" , 12Km , etc
'  [relaxed parsing: allow space between number and unit and allow numbers without units]
Private Function convEGU(ByVal strSource As String, retval As Double) As Boolean
  Dim EguStr   As String
  Dim EguChar  As String
  Dim EguStart As Long
  Dim EguLen   As Long
  Dim EguMult  As String
  Dim EguCoeff As Double
  Dim EguFact  As Long
  Dim EguSym   As String
  Dim EguBase  As Double
    
  EguStr = strSource      'trim niet nodig; alle spaties zijn weg
  EguLen = Len(EguStr)
  For EguStart = 1 To EguLen
    EguChar = Mid$(EguStr, EguStart, 1)  'fix Expo number bug. 25.1.03 VL
    If Not IsNumeric_(EguChar) Then
        Select Case EguChar
            Case "+", "-", "."
                'continue
            Case "E", "e"
                EguChar = Mid$(EguStr, EguStart + 1, 1) 'check next char
                If Not (EguChar = "+" Or EguChar = "-" Or IsNumeric_(EguChar)) Then Exit For
            Case Else
                If IsLetter(EguChar) Then
                    Exit For
                Else
                    convEGU = False: Exit Function
                End If
        End Select
    End If
  Next
  If EguStart = 1 Then
'    Debug.Print "missing coefficient"
    convEGU = False
    Exit Function
  ElseIf EguStart > EguLen Then
'    Debug.Print "missing unit of measure"
    convEGU = False
    Exit Function
  End If
  EguCoeff = Val(Left$(EguStr, EguStart - 1))   'get number
  EguStr = Mid$(EguStr, EguStart)          'extract literal substring
  EguLen = Len(EguStr)
  If EguLen > 1 Then                      'extract multiply factor
    EguMult = Left$(EguStr, 1)
    Select Case EguMult
      Case "p":  EguFact = -12
      Case "n":  EguFact = -9
      Case "µ", "u":  EguFact = -6      '14.2.03 VL
      Case "m":  EguFact = -3
      Case "c":  EguFact = -2
      'Case "d":  EguFact = -1
      'Case "D":  EguFact = 1
      'Case "H":  EguFact = 2           'fix bug. thanks to Javier M. Montalban 14.2.03
      Case "k":  EguFact = 3       '14.2.03 VL
      Case "M":  EguFact = 6
      Case "G":  EguFact = 9
      Case "T":  EguFact = 12
      Case Else: EguFact = 0
    End Select
  Else
    EguFact = 0
  End If
  
  If EguFact <> 0 Then       'extract um symbol
    EguSym = Mid$(EguStr, 2)
  Else
    EguSym = EguStr          ' MR 16-01-03 enable units without factors
  End If
  
  Select Case EguSym         'check if um exists and compute numeric value
    Case "s":   EguBase = 1                 'second
    Case "Hz":  EguBase = 1                 'frequency
    Case "m":   EguBase = 1                 'meter
    Case "g":   EguBase = 0.001             'gramme
    Case "rad", "Rad", "RAD": EguBase = 1   'radiant  '18-10-02 VL
    Case "S":   EguBase = 1                 'siemens
    Case "V":   EguBase = 1                 'volt
    Case "A":   EguBase = 1                 'ampere
    Case "W":   EguBase = 1                 'watt
    Case "F":   EguBase = 1                 'farad
    Case "bar": EguBase = 1                 'bar
    Case "Pa":  EguBase = 1                 'pascal
    Case "Nm":  EguBase = 1                 'newtonmeter
    Case "Ohm", "ohm": EguBase = 1          'ohm     '18-10-02 VL
    Case Else
      ErrMsg = "unknown unit of measure: " + EguSym
      convEGU = False
      Exit Function
  End Select
  retval = EguCoeff * EguBase * 10 ^ EguFact   'finally compute the numeric value
  convEGU = True
End Function
'check if it is a letter
Private Function IsLetter(ByVal char As String) As Boolean
  Dim code As Long
    
  code = Asc(char)
  IsLetter = (65 <= code And code <= 90) Or (97 <= code And code <= 122) Or char = "_"
End Function
'-------------------------------------------------------------------------------
'check for an expression to occur in a list
Private Function InList(ByVal strElem As String, ByVal strList As String) As Boolean
  Dim lstrElem As String
  Dim lstrList As String
    
  lstrList = " " & strList & " "
  lstrElem = " " & strElem & " "
  InList = InStr(1, lstrList, lstrElem, vbTextCompare) > 0
End Function
'-------------------------------------------------------------------------------
' translate a symbolic Constant to its double value
Private Function convSymbConst(ByVal strSource As String, retval As Double) As Boolean
  Dim CostToken   As String
  Dim SymbConst   As String
  CostToken = "#"
  convSymbConst = False
  'check if string is "pi" only for compatibility with previous release.
  If LCase$(strSource) = "pi" Then strSource = strSource + CostToken
  If Right$(strSource, 1) <> CostToken Then Exit Function
  retval = 0
  SymbConst = Left$(strSource, Len(strSource) - 1)
  Select Case SymbConst
    Case "pi", "PI":  retval = PI_              'pi-greek
    Case "pi2": retval = PI_ / 2                'pi-greek/2
    Case "pi3": retval = PI_ / 3                'pi-greek/3
    Case "pi4": retval = PI_ / 4                'pi-greek/4
    Case "e":   retval = 2.71828182845905       'Euler-Napier
    Case "eu":  retval = 0.577215664901533      'Euler-Mascheroni
    Case "phi": retval = 1.61803398874989       'golden ratio
    Case "g":   retval = 9.80665                'Acceleration due to gravity
    Case "G":   retval = 6.672 * 10 ^ -11       'Gravitational constant
    Case "R":   retval = 8.31451                'Gas constant
    Case "eps": retval = 8.854187817 * 10 ^ -12 'Permittivity of vacuum
    Case "mu":  retval = 12.566370614 * 10 ^ -7 'Permeability of vacuum
    Case "c":   retval = 2.99792458 * 10 ^ 8    'Speed of light
    Case "q":   retval = 1.60217733 * 10 ^ -19  'Elementary charge
    Case "me":  retval = 9.1093897 * 10 ^ -31   'Electron rest mass
    Case "mp":  retval = 1.6726231 * 10 ^ -27   'Proton rest mass
    Case "K":   retval = 1.380658 * 10 ^ -23    'Boltzmann constant
    Case "h":   retval = 6.6260755 * 10 ^ -34   'Planck constant
    Case Else:  Err.Raise 2100, , "Constant unknown: " + SymbConst
  End Select
  convSymbConst = True
End Function
'----------------------------------------------------------------------------------
Private Function Acos_(ByVal a As Double) As Double
  If a = 1 Then
    Acos_ = 0
  ElseIf a = -1 Then
    Acos_ = PI_
  Else
    Acos_ = Atn(-a / Sqr(-a * a + 1)) + 2 * Atn(1)
  End If
End Function
Private Function Asin_(ByVal a As Double) As Double
  If Abs(a) = 1 Then
    Asin_ = Sgn(a) * PI_ / 2
  Else
    Asin_ = Atn(a / Sqr(-a * a + 1))
  End If
End Function
Private Function Cosh_(ByVal a As Double) As Double
  Cosh_ = (Exp(a) + Exp(-a)) / 2
End Function
Private Function Sinh_(ByVal a As Double) As Double
  Sinh_ = (Exp(a) - Exp(-a)) / 2
End Function
Private Function Tanh_(ByVal a As Double) As Double
  Tanh_ = (Exp(a) - Exp(-a)) / (Exp(a) + Exp(-a))
End Function
Private Function Acosh_(ByVal a As Double) As Double
  Acosh_ = Log(a + Sqr(a * a - 1))
End Function
Private Function Asinh_(ByVal a As Double) As Double
  Asinh_ = Log(a + Sqr(a * a + 1))
End Function
Private Function Atanh_(ByVal a As Double) As Double
  Atanh_ = Log((1 + a) / (1 - a)) / 2  'bug 3-1-2003 VL
End Function
Private Function round_(ByVal x As Double, ByVal d As Integer) As Double
Dim xi As Double, xd As Double, b As Double
b = 10 ^ d
x = x * b
xi = Int(x)
xd = x - xi
If xd > 0.5 Then xi = xi + 1
round_ = xi / b
End Function
'-------------------------------------------------------------------------------
' calculate Factorial  (bug overflow for n > 12, 8-7-02 VL)
Private Function fact(ByVal N As Double) As Double
  Dim p As Double
  Dim i As Long
  '7.10.2003 thanks to Rodrigo Farinha
  If N < 0 Then
    fact = vbNullString  'raise an error
  Else
    p = 1
    For i = 1 To Int(N)
      p = p * i
    Next
    fact = p
  End If
End Function
'-------------------------------------------------------------------------------
'FIdx the mcm between two integer numbers
Private Function MCM_(ByVal a As Double, ByVal b As Double) As Double
  MCM_ = a * b / MCD_(a, b)
End Function
'-------------------------------------------------------------------------------
' FIdx the MCD between two integer numbers
Private Function MCD_(ByVal a As Double, ByVal b As Double) As Double
  Dim x As Long
  Dim y As Long
  Dim r As Long
  
  y = a
  x = b
  Do Until x = 0
    r = y Mod x
    y = x
    x = r
  Loop
  MCD_ = y
End Function
'-------------------------------------------------------------------------------
' combination n objects, k classes
Private Function Comb(ByVal a As Double, ByVal b As Double) As Double
    Dim N As Long
    Dim k As Long
    Dim y As Long
    Dim i As Long
    
    N = Int(a)
    k = Int(b)
    If N < 0 Then Comb = 0
    If k < 1 Or k > N Then k = 0
    If N = 0 Or k = 0 Or k = N Then Comb = 1: Exit Function
    y = N
    If k > Int(N / 2) Then k = N - k
    For i = 2 To k
        y = y * (N + 1 - i) / i
    Next i
    Comb = y
End Function
'-------------------------------------------------------------------------------
' max value of 2 numbers
Private Function max_(ByVal a As Double, ByVal b As Double) As Double
  If a > b Then max_ = a Else max_ = b
End Function
'-------------------------------------------------------------------------------
' min value of 2 numbers
Private Function min_(ByVal a As Double, ByVal b As Double) As Double
  If a < b Then min_ = a Else min_ = b
End Function
'-------------------------------------------------------------------------------
' percentage
Private Function percent(ByVal i As Long) As Double
Dim ris As Double, a As Double, b As Double
    b = ET(i).Arg(1).Value
    ris = b / 100
If i > 1 Then
    Select Case ET(i - 1).FunTok
        Case symPlus, symMinus, symMul, symDiv
            a = ET(i - 1).Arg(1).Value
            If a <> 0 Then ris = a * ris
        Case Else
            'nothing to do
    End Select
End If
percent = ris
End Function
'-------------------------------------------------------------------------------
' count number of abs sybol sets in formula
Private Function NabsCount(ByVal s As String) As Long
  Dim N As Long
  Dim p As Long

  N = 0
  p = InStr(s, "|")
  Do While p > 0
    p = p + 1
    N = N + 1
    p = InStr(p, s, "|")
  Loop
  NabsCount = N
End Function
'-------------------------------------------------------------------------------
' error distribution function
Private Function erf_(ByVal x As Double) As Double
Const MAXLOOP = 400
Const TINY = 0.000000000000001
Dim y As Double, t As Double, p As Double, s As Double, i As Long
Dim A0 As Double, B0 As Double, A1 As Double, b1 As Double, A2 As Double, b2 As Double
Dim f1 As Double, f2 As Double, g As Double, d As Double
If x <= 2 Then
    t = 2 * x * x
    p = 1
    s = 1
    For i = 3 To MAXLOOP Step 2
        p = p * t / i
        s = s + p
        If p < TINY Then Exit For
    Next
    y = 2 * s * x * Exp(-x * x) / Sqr(PI_)
Else
    A0 = 1: B0 = 0
    A1 = 0: b1 = 1
    f1 = 0
    For i = 1 To MAXLOOP
        g = 2 - (i Mod 2)
        A2 = g * x * A1 + i * A0
        b2 = g * x * b1 + i * B0
        f2 = A2 / b2
        d = Abs(f2 - f1)
        If d < TINY Then Exit For
        A0 = A1: B0 = b1
        A1 = A2: b1 = b2
        f1 = f2
    Next
    y = 1 - 2 * Exp(-x * x) / (2 * x + f2) / Sqr(PI_)

End If
erf_ = y
End Function
'-------------------------------------------------------------------------------
' gamma  - Lanczos approximation algorithm for gamma function
Private Sub gamma_split(x As Double, mantissa As Double, expo As Long)
Dim z As Double, Cf(14) As Double, w As Double, i As Long, s As Double, p As Double
Const DOUBLEPI = 6.28318530717959
Const G_ = 4.7421875  '607/128
z = x - 1

Cf(0) = 0.999999999999997
Cf(1) = 57.1562356658629
Cf(2) = -59.5979603554755
Cf(3) = 14.1360979747417
Cf(4) = -0.49191381609762
Cf(5) = 3.39946499848119E-05
Cf(6) = 4.65236289270486E-05
Cf(7) = -9.83744753048796E-05
Cf(8) = 1.58088703224912E-04
Cf(9) = -2.10264441724105E-04
Cf(10) = 2.17439618115213E-04
Cf(11) = -1.64318106536764E-04
Cf(12) = 8.44182239838528E-05
Cf(13) = -2.61908384015814E-05
Cf(14) = 3.68991826595316E-06

w = Exp(G_) / Sqr(DOUBLEPI)
s = Cf(0)
For i = 1 To 14
    s = s + Cf(i) / (z + i)
Next
s = s / w
p = Log((z + G_ + 0.5) / Exp(1)) * (z + 0.5) / Log(10)
'split in mantissa and exponent to avoid overflow
expo = Int(p)
p = p - Int(p)
mantissa = 10 ^ p * s
'rescaling
p = Int(Log(mantissa) / Log(10))
mantissa = mantissa * 10 ^ -p
expo = expo + p
End Sub
'-------------------------------------------------------------------------------
' gamma function
Private Function gamma_(ByVal x As Double) As Double
Dim mantissa As Double, expo As Long
gamma_split x, mantissa, expo
gamma_ = mantissa * 10 ^ expo
End Function
'-------------------------------------------------------------------------------
' logarithm gamma function
Private Function gammaln_(ByVal x As Double) As Double
Dim mantissa As Double, expo As Long
gamma_split x, mantissa, expo
gammaln_ = Log(mantissa) + expo * Log(10)
End Function
'-------------------------------------------------------------------------------
' digamma function
Private Function digamma_(ByVal x As Double) As Double
Dim b1(11) As Double, b2(11) As Double
Dim z As Double, s As Double, k As Long, Tmp As Double, y As Double
Const LIM_LOW = 8
'Bernoulli's numbers
b1(0) = 1: b2(0) = 1
b1(1) = 1: b2(1) = 6
b1(2) = -1: b2(2) = 30
b1(3) = 1: b2(3) = 42
b1(4) = -1: b2(4) = 30
b1(5) = 5: b2(5) = 66
b1(6) = -691: b2(6) = 2730
b1(7) = 7: b2(7) = 6
b1(8) = -3617: b2(8) = 360
b1(9) = 43867: b2(9) = 798
b1(10) = -174611: b2(10) = 330
b1(11) = 854513: b2(11) = 138
If x <= LIM_LOW Then
    z = x - 1 + LIM_LOW
Else
    z = x - 1
End If
s = 0
For k = 1 To 11
    Tmp = b1(k) / b2(k) / k / z ^ (2 * k)
    s = s + Tmp
Next
y = Log(z) + 0.5 * (1 / z - s)

If x <= LIM_LOW Then
    s = 0
    For k = 0 To LIM_LOW - 1
        s = s + 1 / (x + k)
    Next
    y = y - s
End If

digamma_ = y
End Function
'-------------------------------------------------------------------------------
' beta function
Private Function beta_(ByVal z As Double, ByVal w As Double) As Double
beta_ = Exp(gammaln_(z) + gammaln_(w) - gammaln_(z + w))
End Function
'-------------------------------------------------------------------------------
' Riemman's zeta function
Private Function Zeta_(ByVal x As Double) As Double
Dim Cnk As Double, k As Long, N As Long, s As Double, s1 As Double, coeff As Double
Const N_MAX = 1000
Const TINY = 1E-16
N = 0: s = 0
Do
    s1 = 0: Cnk = 1
    For k = 0 To N
        If k > 0 Then Cnk = Cnk * (N - k + 1) / k
        s1 = s1 + (-1) ^ k * Cnk / (k + 1) ^ x
    Next k
    coeff = s1 / 2 ^ (1 + N)
    s = s + coeff
    N = N + 1
Loop Until Abs(coeff) < TINY Or N > N_MAX
Zeta_ = s / (1 - 2 ^ (1 - x))
End Function
'-------------------------------------------------------------------------------
' exponential integral Ei(x) for x >0.
Private Function exp_integr(ByVal x As Double) As Double
'original from NUMERICAL RECIPES in FORTRAN 77, by Numerical Recipes Software
Const EPS = 0.000000000000001, EULER = 0.577215664901532, MAXIT = 100, FPMIN = 1E-30
Dim k As Long, fact As Double, prev As Double, sum As Double, term As Double
If (x <= 0) Then Exit Function '
If (x < FPMIN) Then          'Special case: avoid failure of convergence test be-
      exp_integr = Log(x) + EULER            'cause of under ow.
ElseIf (x <= -Log(EPS)) Then 'Use power series.
     sum = 0
     fact = 1
     For k = 1 To MAXIT
        fact = fact * x / k
        term = fact / k
        sum = sum + term
        If (term < EPS * sum) Then Exit For
    Next
    exp_integr = sum + Log(x) + EULER
Else 'Use asymptotic series.
    sum = 0 'Start with second term.
    term = 1
    For k = 1 To MAXIT
      prev = term
      term = term * k / x
      If (term < EPS) Then Exit For 'Since al sum is greater than one, term itself ap-
      If (term < prev) Then
         sum = sum + term 'Still converging: add new term.
      Else
         sum = sum - prev 'Diverging: subtract previous term and exit.
         Exit For
      End If
    Next
    exp_integr = Exp(x) * (1 + sum) / x
End If
End Function

Private Function cvDegree(ByVal DMS As String, ByRef Angle As Double) As Boolean
'converts a string dd°mm'ss" (degrees, minutes, seconds) into a decimal-degree angle
Dim p1 As Long, p2 As Long, p3 As Long
Dim dd As Integer, mm As Integer, ss As Long
cvDegree = False
p1 = InStr(DMS, Chr$(176)) ' degrees °
p2 = InStr(DMS, Chr$(39))  ' minutes '
p3 = InStr(DMS, Chr$(34))  ' seconds "
If p1 = 0 And p2 = 0 And p3 = 0 Then Exit Function
On Error Resume Next
dd = CInt(Mid$(DMS, 1, p1 - 1))
mm = CInt(Mid$(DMS, p1 + 1, p2 - p1 - 1))
ss = CInt(Mid$(DMS, p2 + 1, p3 - p2 - 1))
On Error GoTo 0
If mm > 60 Or ss > 60 Then Exit Function
Angle = dd + (mm + ss / 60) / 60
cvDegree = True
End Function

'Comment this subroutine for MobileVB
Sub ET_Dump(ByVal level As Integer)
  Dim i As Long

  Debug.Print "--------------------------------", "Dump Level " & CStr(level)
  Debug.Print "Id", "Fun", "ArgTop", _
              "A0 Idx", "Arg0 Name", "Arg0 Value", _
              "A1 Idx", "Arg1 Name", "Arg1 Value", _
              "ArgOf", "ArgIdx", "Value", "PriLvl", "PosInExpr", "PriIdx"
  For i = 1 To ETtop
    With ET(i)
      Debug.Print i, .Fun, .ArgTop, _
                  .Arg(1).Idx, .Arg(1).Nome, .Arg(1).Value, _
                  .Arg(2).Idx, .Arg(2).Nome, .Arg(2).Value, _
                  .ArgOf, .ArgIdx, .Value, .PriLvl, .PosInExpr, .PriIdx
    End With
  Next
End Sub

Private Function IsNumeric_(ByVal x As String) As Boolean
  'independent check function from international system setting
  'x must have always the decimal point "123.756", ".0056", "1.3455E-12"
  'by LV, 27.2.2003. Thanks to Michael Ruder
  Dim p As Integer
  IsNumeric_ = False
  If InStr(x, ",") Then Exit Function
  IsNumeric_ = IsNumeric(x)
  If Not IsNumeric_ Then
  'try with decimal comma
  p = InStr(x, ".")
  If p > 0 Then
  Mid$(x, p, 1) = ","
  IsNumeric_ = IsNumeric(x)
  End If
  End If
End Function

Private Function MiRoot_(a As Double, N As Double) As Double
Dim m As Long
'7.10.2003 thanks to Rodigro Farinha
'algebric extension of root for a<0
    m = Int(N)  'only integer here
    If m Mod 2 = 0 Then    'b is even => root in a<0 doesn´t exist
        If a < 0 Then
            MiRoot_ = vbNullString 'raise an error
        Else
            MiRoot_ = a ^ (1 / m)
        End If
    Else    'b is odd => root in a<0 exists
        MiRoot_ = Sgn(a) * Abs(a) ^ (1 / m)
    End If
End Function

