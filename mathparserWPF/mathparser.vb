
' 
' * Copyright (C) 2012-2014 Artem Los, Ljubomir Blanusa
' * All rights reserved.
' * 
' * Please see the license file in the project folder,
' * 
' * Please feel free to ask me directly at my email!
' *  artem@artemlos.net
' 


Imports System.Collections.Generic
Imports System.Linq
Imports System.Globalization
Imports System.Collections.ObjectModel

Namespace mparserNonIL
    ''' <summary>
    ''' This is a mathematical expression parser that allows you to parser a string value,
    ''' perform the required calculations, and return a value in form of a Double.
    ''' </summary>
    Public Class MathParser

        ''' <summary>
        ''' Basic constructor.
        ''' </summary>
        ''' <remarks>Use boolean flags to load basic operators, functions, and variables</remarks>
        Public Sub New()

        End Sub

        ''' <summary>
        ''' This constructor will add some basic operators, functions, and variables
        ''' to the parser. Please note that you are able to change that using
        ''' boolean flags
        ''' </summary>
        ''' <param name="loadPreDefinedFunctions">This will load "abs", "cos"...</param>
        ''' <param name="loadPreDefinedOperators">This will load "%", "*", ":", "/", "+", "-", ">", "&lt;", "="</param>
        ''' <param name="loadPreDefinedVariables">This will load "pi" and "e"</param>
        Public Sub New(loadPreDefinedFunctions As Boolean, loadPreDefinedOperators As Boolean, loadPreDefinedVariables As Boolean)
            If loadPreDefinedOperators Then
                ' by default, we will load basic arithmetic operators.
                ' please note, its possible to do it either inside the constructor,
                ' or outside the class. the lowest value will be executed first!
                OperatorList.Add("%")
                ' modulo
                OperatorList.Add("^")
                ' to the power of
                OperatorList.Add(":")
                ' division 1
                OperatorList.Add("/")
                ' division 2
                OperatorList.Add("*")
                ' multiplication
                OperatorList.Add("-")
                ' subtraction
                OperatorList.Add("+")
                ' addition
                OperatorList.Add(">")
                ' greater than
                OperatorList.Add("<")
                ' less than
                OperatorList.Add("=")
                ' are equal

                ' when an operator is executed, the parser needs to know how.
                ' this is how you can add your own operators. note, the order
                ' in this list does not matter.
                _operatorAction.Add("%", Function(numberA, numberB) numberA Mod numberB)
                _operatorAction.Add("^", Function(numberA, numberB) Math.Pow(numberA, numberB))
                _operatorAction.Add(":", Function(numberA, numberB) numberA / numberB)
                _operatorAction.Add("/", Function(numberA, numberB) numberA / numberB)
                _operatorAction.Add("*", Function(numberA, numberB) numberA * numberB)
                _operatorAction.Add("+", Function(numberA, numberB) numberA + numberB)
                _operatorAction.Add("-", Function(numberA, numberB) numberA - numberB)

                _operatorAction.Add(">", Function(numberA, numberB) If(numberA > numberB, 1, 0))
                _operatorAction.Add("<", Function(numberA, numberB) If(numberA < numberB, 1, 0))
                _operatorAction.Add("=", Function(numberA, numberB) If(numberA = numberB, 1, 0))
            End If


            If loadPreDefinedFunctions Then
                ' these are the basic functions you might be able to use.
                LocalFunctions.Add("abs", Function(x) Math.Abs(x(0)))
                LocalFunctions.Add("cos", Function(x) Math.Cos(x(0)))
                LocalFunctions.Add("cosh", Function(x) Math.Cosh(x(0)))
                LocalFunctions.Add("arccos", Function(x) Math.Acos(x(0)))
                LocalFunctions.Add("sin", Function(x) Math.Sin(x(0)))
                LocalFunctions.Add("sinh", Function(x) Math.Sinh(x(0)))
                LocalFunctions.Add("arcsin", Function(x) Math.Asin(x(0)))
                LocalFunctions.Add("tan", Function(x) Math.Tan(x(0)))
                LocalFunctions.Add("tanh", Function(x) Math.Tanh(x(0)))
                LocalFunctions.Add("arctan", Function(x) Math.Atan(x(0)))

                LocalFunctions.Add("sqrt", Function(x) Math.Sqrt(x(0)))
                LocalFunctions.Add("rem", Function(x) Math.IEEERemainder(x(0), x(1)))
                LocalFunctions.Add("root", Function(x) Math.Pow(x(0), 1 / x(1)))

                LocalFunctions.Add("pow", Function(x) Math.Pow(x(0), x(1)))
                LocalFunctions.Add("exp", Function(x) Math.Exp(x(0)))

                LocalFunctions.Add("ln", Function(x) Math.Log(x(0), Math.E))
                ' input[0] is the number
                ' input[1] is the base
                LocalFunctions.Add("log", Function(x)
                                              Select Case x.Length
                                                  Case 1
                                                      Return Math.Log10(x(0))
                                                  Case 2
                                                      Return Math.Log(x(0), x(1))
                                                  Case Else
                                                      Return 0
                                                      ' false
                                              End Select
                                          End Function)

                LocalFunctions.Add("round", Function(x) Math.Round(x(0)))
                LocalFunctions.Add("truncate", Function(x) If(x(0) < 0, -Math.Floor(-x(0)), Math.Floor(x(0))))
                LocalFunctions.Add("floor", Function(x) Math.Floor(x(0)))
                LocalFunctions.Add("ceiling", Function(x) Math.Ceiling(x(0)))

                LocalFunctions.Add("sign", Function(x) Math.Sign(x(0)))
            End If

            If loadPreDefinedVariables Then
                ' local variables such as pi can also be added into the parser.
                LocalVariables.Add("pi", Math.PI) '3.14159265358979
                LocalVariables.Add("pi2", Math.PI * 2) ' 6.28318530717959
                LocalVariables.Add("pi05", Math.PI / 2) '1.5707963267949
                LocalVariables.Add("e", Math.E) '2.71828182845905
            End If
        End Sub


        ' THE LOCAL VARIABLES 

        Private _operatorList As New List(Of String)()
        ''' <summary>
        ''' All operators should be inside this property.
        ''' The first operator is executed first, et cetera.
        ''' An operator may only be ONE character.
        ''' </summary>
        Public Property OperatorList() As List(Of String)
            Get
                Return _operatorList
            End Get
            Set(value As List(Of String))
                _operatorList = value
            End Set
        End Property
        Private _operatorAction As New Dictionary(Of String, Func(Of Double, Double, Double))()
        ''' <summary>
        ''' When adding a variable in the OperatorList property, you need to assign how that operator should work.
        ''' </summary>
        Public Property OperatorAction() As Dictionary(Of String, Func(Of Double, Double, Double))
            Get
                Return _operatorAction
            End Get
            Set(value As Dictionary(Of String, Func(Of Double, Double, Double)))
                _operatorAction = value
            End Set
        End Property
        Private _localFunctions As New Dictionary(Of String, Func(Of Double(), Double))()
        ''' <summary>
        ''' All functions that you want to define should be inside this property.
        ''' </summary>
        Public Property LocalFunctions() As Dictionary(Of String, Func(Of Double(), Double))
            Get
                Return _localFunctions
            End Get
            Set(value As Dictionary(Of String, Func(Of Double(), Double)))
                _localFunctions = value
            End Set
        End Property
        Private _localVariables As New Dictionary(Of String, Double)()
        ''' <summary>
        ''' All variables that you want to define should be inside this property.
        ''' </summary>
        Public Property LocalVariables() As Dictionary(Of String, Double)
            Get
                Return _localVariables
            End Get
            Set(value As Dictionary(Of String, Double))
                _localVariables = value
            End Set
        End Property

        Private _cultureInfo As CultureInfo = cultureInfo.InvariantCulture
        ''' <summary>
        ''' Use Culture info to correctly parse numbers in different formats.
        ''' </summary>
        ''' <value>Valid CultureInfo</value>
        ''' <returns>Current CultureInfo used for parsing numbers and separating numbers</returns>
        ''' <remarks>User should use decimal numbers without number grouping mark.</remarks>
        Public Property cultureInfo() As CultureInfo
            Get
                Return _cultureInfo
            End Get
            Set(value As CultureInfo)
                _cultureInfo = value
            End Set
        End Property


        ' PARSER FUNCTIONS, PUBLIC 

        ''' <summary>
        ''' Enter the math expression in form of a string, numbers must be in the same culture as parser, and without grouping mark.
        ''' </summary>
        ''' <param name="mathExpression"></param>
        ''' <returns></returns>
        Public Function Parse(mathExpression As String) As Double
            Return MathParserLogic(Scanner(mathExpression))
        End Function

        ''' <summary>
        ''' Enter the math expression in form of a list of tokens.
        ''' </summary>
        ''' <param name="mathExpression"></param>
        ''' <returns></returns>
        Public Function Parse(mathExpression As ReadOnlyCollection(Of String)) As Double
            Return MathParserLogic(New List(Of String)(mathExpression))
        End Function

        ''' <summary>
        ''' This will convert a string expression into a list of tokens that can be later executed by Parse
        ''' </summary>
        ''' <param name="mathExpression"></param>
        ''' <returns>A ReadOnlyCollection</returns>
        Public Function GetTokens(mathExpression As String) As ReadOnlyCollection(Of String)
            Return New ReadOnlyCollection(Of String)(Scanner(mathExpression))
        End Function

        ''' <summary>
        ''' This will correct sqrt() and arctan() written in different ways only.
        ''' </summary>
        ''' <param name="input"></param>
        ''' <returns></returns>
        Private Function Correction(input As String) As String
            ' Word corrections

            input = System.Text.RegularExpressions.Regex.Replace(input, "\b(sqr|sqrt)\b", "sqrt", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            input = System.Text.RegularExpressions.Regex.Replace(input, "\b(atan2|arctan2)\b", "arctan2", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            '... and more

            Return input
        End Function


        ' UNDER THE HOOD - THE CORE OF THE PARSER 

        Private Function Scanner(expr As String) As List(Of String)
            ' SCANNING THE INPUT STRING AND CONVERT IT INTO TOKENS

            Dim tokens = New List(Of String)()
            Dim vector = ""

            '_expr = _expr.Replace(" ", ""); // remove white space
            expr = expr.Replace("+-", "-")
            ' some basic arithmetical rules
            expr = expr.Replace("-+", "-")
            expr = expr.Replace("--", "+")

            For i As Integer = 0 To expr.Length - 1
                Dim ch As Char = expr(i)

                If Char.IsWhiteSpace(ch) Then
                    ' could also be used to remove white spaces.
                ElseIf [Char].IsLetter(ch) Then
                    If i <> 0 AndAlso ([Char].IsDigit(expr(i - 1)) OrElse [Char].IsDigit(expr(i - 1)) OrElse expr(i - 1) = ")") Then
                        tokens.Add("*")
                    End If

                    vector = vector + ch

                    While (i + 1) < expr.Length AndAlso [Char].IsLetterOrDigit(expr(i + 1))
                        ' here is it is possible to choose whether you want variables that only contain letters with or without digits.
                        i += 1
                        vector = vector + expr(i)
                    End While

                    tokens.Add(vector)
                    vector = ""
                ElseIf [Char].IsDigit(ch) Then
                    vector = vector + ch
                    While (i + 1) < expr.Length AndAlso ([Char].IsDigit(expr(i + 1)) OrElse expr(i + 1) = cultureInfo.NumberFormat.NumberDecimalSeparator)
                        ' removed || _expr[i + 1] == ','
                        i += 1
                        vector = vector + expr(i)
                    End While

                    tokens.Add(vector)
                    vector = ""
                ElseIf (i + 1) < expr.Length AndAlso (ch = "-" OrElse ch = "+") AndAlso [Char].IsDigit(expr(i + 1)) AndAlso (i = 0 OrElse OperatorList.IndexOf(expr(i - 1).ToString()) <> -1 OrElse ((i - 1) > 0 AndAlso expr(i - 1) = "(")) Then
                    ' removed CultureInfo.InvariantCulture from expr[i-1].ToString(...)
                    ' if the above is true, then, the token for that negative number will be "-1", not "-","1".
                    ' to sum up, the above will be true if the minus sign is in front of the number, but
                    ' at the beginning, for example, -1+2, or, when it is inside the brakets (-1).
                    ' NOTE: this works for + sign as well!
                    vector = vector + ch

                    While (i + 1) < expr.Length AndAlso ([Char].IsDigit(expr(i + 1)) OrElse expr(i + 1) = cultureInfo.NumberFormat.NumberDecimalSeparator)
                        ' removed || _expr[i + 1] == ','
                        i += 1
                        vector = vector + expr(i)
                    End While

                    tokens.Add(vector)
                    vector = ""
                ElseIf ch = "(" Then
                    If i <> 0 AndAlso ([Char].IsDigit(expr(i - 1)) OrElse [Char].IsDigit(expr(i - 1)) OrElse expr(i - 1) = ")") Then
                        tokens.Add("*")
                        ' if we remove this line(above), we would be able to have numbers in function names. however, then we can't parser 3(2+2)
                        tokens.Add("(")
                    Else
                        tokens.Add("(")
                    End If
                Else
                    tokens.Add(ch.ToString())
                End If
            Next

            Return tokens
            'return MathParserLogic(_tokens);
        End Function

        Private Function MathParserLogic(tokens As List(Of String)) As Double
            ' CALCULATING THE EXPRESSIONS INSIDE THE BRACKETS
            ' IF NEEDED, EXECUTE A FUNCTION

            ' Variables replacement
            For i As Integer = 0 To tokens.Count - 1
                If LocalVariables.Keys.Contains(tokens(i)) Then
                    tokens(i) = LocalVariables(tokens(i)).ToString(cultureInfo)
                End If
            Next
            While tokens.IndexOf("(") <> -1
                ' getting data between "(", ")"
                Dim open = tokens.LastIndexOf("(")
                Dim close = tokens.IndexOf(")", open)
                ' in case open is -1, i.e. no "(" // , open == 0 ? 0 : open - 1
                If open >= close Then
                    ' if there is no closing bracket, throw a new exception
                    Throw New ArithmeticException("No closing bracket/parenthesis! tkn: " + open.ToString(cultureInfo))
                End If

                Dim roughExpr = New List(Of String)()

                For i As Integer = open + 1 To close - 1
                    roughExpr.Add(tokens(i))
                Next

                Dim result As Double
                ' the temporary result is stored here
                Dim functioName = tokens(If(open = 0, 0, open - 1))
                Dim args = New Double(-1) {}

                If LocalFunctions.Keys.Contains(functioName) Then
                    If roughExpr.Contains(cultureInfo.TextInfo.ListSeparator) Then
                        ' converting all arguments into a Double array
                        For i As Integer = 0 To roughExpr.Count - 1
                            Dim firstCommaOrEndOfExpression = If(roughExpr.IndexOf(cultureInfo.TextInfo.ListSeparator, i) <> -1, roughExpr.IndexOf(cultureInfo.TextInfo.ListSeparator, i), roughExpr.Count)
                            Dim defaultExpr = New List(Of String)()

                            While i < firstCommaOrEndOfExpression
                                defaultExpr.Add(roughExpr(i))
                                i += 1
                            End While

                            ' changing the size of the array of arguments
                            Array.Resize(args, args.Length + 1)

                            If defaultExpr.Count = 0 Then
                                args(args.Length - 1) = 0
                            Else
                                args(args.Length - 1) = BasicArithmeticalExpression(defaultExpr)
                            End If
                        Next

                        ' finnaly, passing the arguments to the given function
                        result = Double.Parse(LocalFunctions(functioName)(args).ToString(cultureInfo), cultureInfo)
                    Else
                        ' but if we only have one argument, then we pass it directly to the function
                        Array.Resize(args, 1)
                        args(0) = BasicArithmeticalExpression(roughExpr)
                        result = Double.Parse(LocalFunctions(functioName)(args).ToString(cultureInfo), cultureInfo)
                    End If
                Else
                    ' if no function is need to execute following expression, pass it
                    ' to the "BasicArithmeticalExpression" method.
                    result = BasicArithmeticalExpression(roughExpr)
                End If

                ' when all the calculations have been done
                ' we replace the "opening bracket with the result"
                ' and removing the rest.
                tokens(open) = result.ToString(cultureInfo)
                tokens.RemoveRange(open + 1, close - open)
                If LocalFunctions.Keys.Contains(functioName) Then
                    ' if we also executed a function, removing
                    ' the function name as well.
                    tokens.RemoveAt(open - 1)
                End If
            End While

            ' at this point, we should have replaced all brackets
            ' with the appropriate values, so we can simply
            ' calculate the expression. it's not so complex
            ' any more!
            Return BasicArithmeticalExpression(tokens)
        End Function

        Private Function BasicArithmeticalExpression(tokens As List(Of String)) As Double
            ' PERFORMING A BASIC ARITHMETICAL EXPRESSION CALCULATION
            ' THIS METHOD CAN ONLY OPERATE WITH NUMBERS AND OPERATORS
            ' AND WILL NOT UNDERSTAND ANYTHING BEYOND THAT.

            If tokens.Count = 1 Then
                Return Double.Parse(tokens(0), cultureInfo)
            Else
                For Each op As String In OperatorList
                    While tokens.IndexOf(op) <> -1
                        'this covers situations where sign is a token before any function
                        Dim result As Double
                        Dim opPlace = tokens.IndexOf(op)
                        If opPlace = 0 AndAlso tokens(0) = "-" Then
                            result = -Double.Parse(tokens(1), cultureInfo)
                            tokens(0) = result.ToString(cultureInfo)
                            tokens.RemoveAt(1)
                        Else
                            Dim numberA = Double.Parse(tokens(opPlace - 1), cultureInfo)
                            Dim numberB As Double

                            'this covers situations like: -sin(5)*-pi   (x * - 3.14)
                            'it would be better if sign is rembembered before function is evaluated, and then applied to the result of evaluated function after
                            'this way, this method would only deal with raw numbers and single operators in between
                            Dim removeplaces As Integer = 2
                            If tokens(opPlace + 1) = "-" Then
                                numberB = -Double.Parse(tokens(opPlace + 2), cultureInfo)
                                removeplaces = removeplaces + 1
                            ElseIf tokens(opPlace + 1) = "+" Then
                                numberB = Double.Parse(tokens(opPlace + 2), cultureInfo)
                                removeplaces = removeplaces + 1
                            Else
                                numberB = Double.Parse(tokens(opPlace + 1), cultureInfo)
                            End If

                            result = OperatorAction(op)(numberA, numberB)
                            tokens(opPlace - 1) = result.ToString(cultureInfo)
                            tokens.RemoveRange(opPlace, removeplaces)
                        End If
                    End While
                Next
                Return Double.Parse(tokens(0), cultureInfo)
            End If
        End Function

    End Class
End Namespace


