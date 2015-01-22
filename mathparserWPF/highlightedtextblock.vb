Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports System.Globalization

Namespace htb
    Public Class HighlightedTextBlock
        Inherits TextBlock

#Region "colors"
        'definitions of basic colors used for coloring functions, variables and numbers.

        Public Property HighlightFuncColor() As Brush
            Get
                If DirectCast(GetValue(HighlightFuncColorProperty), Brush) Is Nothing Then
                    SetValue(HighlightFuncColorProperty, Brushes.Black)
                End If
                Return DirectCast(GetValue(HighlightFuncColorProperty), Brush)
            End Get
            Set(value As Brush)
                SetValue(HighlightFuncColorProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightFuncColorProperty As DependencyProperty = DependencyProperty.Register("HighlightFuncColor", GetType(Brush), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightFuncColorChangedCallback)))

        Private Shared Sub HighlightFuncColorChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub

        Public Property HighlightVarColor() As Brush
            Get
                If DirectCast(GetValue(HighlightVarColorProperty), Brush) Is Nothing Then
                    SetValue(HighlightVarColorProperty, Brushes.Black)
                End If
                Return DirectCast(GetValue(HighlightVarColorProperty), Brush)
            End Get
            Set(value As Brush)
                SetValue(HighlightVarColorProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightVarColorProperty As DependencyProperty = DependencyProperty.Register("HighlightVarColor", GetType(Brush), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightVarColorChangedCallback)))


        Private Shared Sub HighlightVarColorChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub

        Public Property HighlightNumColor() As Brush
            Get
                If DirectCast(GetValue(HighlightNumColorProperty), Brush) Is Nothing Then
                    SetValue(HighlightNumColorProperty, Brushes.Black)
                End If
                Return DirectCast(GetValue(HighlightNumColorProperty), Brush)
            End Get
            Set(value As Brush)
                SetValue(HighlightNumColorProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightNumColorProperty As DependencyProperty = DependencyProperty.Register("HighlightNumColor", GetType(Brush), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightNumColorChangedCallback)))

        Private Shared Sub HighlightNumColorChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub
#End Region

#Region "text"
        'texts to highlight or to control highlight
        Public Property HighlightText() As String
            Get
                Return DirectCast(GetValue(HighlightTextProperty), String)
            End Get
            Set(value As String)
                SetValue(HighlightTextProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightTextProperty As DependencyProperty = DependencyProperty.Register("HighlightText", GetType(String), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightTextChangedCallback)))

        Private Shared Sub HighlightTextChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub

        Public Property HighlightFunctions() As List(Of String)
            Get
                If DirectCast(GetValue(HighlightFunctionsProperty), List(Of String)) Is Nothing Then
                    SetValue(HighlightFunctionsProperty, New List(Of String)())
                End If
                Return DirectCast(GetValue(HighlightFunctionsProperty), List(Of String))
            End Get
            Set(value As List(Of String))
                SetValue(HighlightFunctionsProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightFunctionsProperty As DependencyProperty = DependencyProperty.Register("HighlightFunctions", GetType(List(Of String)), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightFunctionsChangedCallback)))

        Private Shared Sub HighlightFunctionsChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub

        Public Property HighlightVars() As List(Of String)
            Get
                If DirectCast(GetValue(HighlightVarsProperty), List(Of String)) Is Nothing Then
                    SetValue(HighlightVarsProperty, New List(Of String)())
                End If
                Return DirectCast(GetValue(HighlightVarsProperty), List(Of String))
            End Get
            Set(value As List(Of String))
                SetValue(HighlightVarsProperty, value)
            End Set
        End Property

        Public Shared ReadOnly HighlightVarsProperty As DependencyProperty = DependencyProperty.Register("HighlightVars", GetType(List(Of String)), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf HighlightVarsChangedCallback)))

        Private Shared Sub HighlightVarsChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            If stb IsNot Nothing Then
                stb.Text = stb.HighlightText
            End If
            stb.update()
        End Sub


        Public Property Culture() As CultureInfo
            Get
                If DirectCast(GetValue(CultureProperty), CultureInfo) Is Nothing Then
                    SetValue(CultureProperty, CultureInfo.InvariantCulture)
                End If
                Return DirectCast(GetValue(CultureProperty), CultureInfo)
            End Get
            Set(value As CultureInfo)
                SetValue(CultureProperty, value)
            End Set
        End Property

        Public Shared ReadOnly CultureProperty As DependencyProperty = DependencyProperty.Register("Culture", GetType(CultureInfo), GetType(HighlightedTextBlock), New PropertyMetadata(New PropertyChangedCallback(AddressOf CultureChangedCallback)))

        Private Shared Sub CultureChangedCallback(obj As DependencyObject, e As DependencyPropertyChangedEventArgs)
            Dim stb As HighlightedTextBlock = TryCast(obj, HighlightedTextBlock)
            stb.update()
        End Sub


#End Region

#Region "highlight control"
        Private Sub update()
            Dim originaltext As String = MyBase.Text
            If originaltext = "" Then Exit Sub
            Dim barr(originaltext.Length - 1) As Byte
            barr = fillNums(barr, originaltext, Culture.NumberFormat.NumberDecimalSeparator, 1) 'First! This simplifies regex if numbers are in names of vars.
            barr = fillarray(barr, originaltext, HighlightFunctions, 2)
            barr = fillarray(barr, originaltext, HighlightVars, 3)

            Inlines.Clear()
            For i = 0 To originaltext.Length - 1
                Dim run As New Documents.Run(originaltext(i))
                Select Case barr(i)
                    Case 1
                        'color numbers first
                        run.Foreground = HighlightNumColor
                    Case 2
                        'color functions
                        run.Foreground = HighlightFuncColor
                    Case 3
                        'color variables
                        run.Foreground = HighlightVarColor
                    Case Else
                        'default color
                        run.Foreground = Foreground
                End Select
                Inlines.Add(run)
            Next
        End Sub

        Private Function fillarray(barr() As Byte, text As String, list As List(Of String), id As Byte) As Byte()
            For Each s As String In list
                Dim matchColl As MatchCollection = Regex.Matches(text, "\b" & s & "\b") 'string is a word boundary
                For Each match As Match In matchColl
                    For i As Integer = match.Index To match.Index + match.Length - 1
                        barr(i) = id
                    Next
                Next
            Next
            Return barr
        End Function

        Private Function fillNums(barr() As Byte, text As String, decimalmark As String, id As Byte)
            Dim matchColl As MatchCollection = Regex.Matches(text, "[-+]?\d+\" & decimalmark & "*\d*")
            '[-+]?\d+\.*\d* match +- followed by digit with maybe . and maybe more digits
            For Each match As Match In matchColl
                For i As Integer = match.Index To match.Index + match.Length - 1
                    barr(i) = id
                Next
            Next
            Return barr
        End Function
#End Region

    End Class

End Namespace