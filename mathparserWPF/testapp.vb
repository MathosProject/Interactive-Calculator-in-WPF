Imports System.ComponentModel
Imports System.Collections.ObjectModel
Imports System.Globalization

'Simple WPF application for evaluate math using modified Mathosparser.
'Released under MIT Lincence
'Ljubomir Blanusa 2015.

Namespace testapp

    Public Class main
        Implements INotifyPropertyChanged
        Private mparser As mathparserWPF.mparserNonIL.MathParser
        Private _culture As CultureInfo = CultureInfo.CurrentCulture

        Sub New()
            mparser = New mathparserWPF.mparserNonIL.MathParser(True, True, True)
            mparser.cultureInfo = selectedCulture
        End Sub

        Private _mytext As String
        Public Property ToEvaluate As String
            Get
                Return _mytext
            End Get
            Set(value As String)
                _mytext = value
                recalc()
                OnPropertyChanged(New PropertyChangedEventArgs("ToEvaluate"))
            End Set
        End Property

        Private Property _result As String
        Public Property Result As String
            Get
                Return _result
            End Get
            Set(value As String)
                _result = value
                OnPropertyChanged(New PropertyChangedEventArgs("Result"))
            End Set
        End Property

        Private _hasError As Boolean
        Public Property hasError As Boolean
            Get
                Return _hasError
            End Get
            Set(value As Boolean)
                _hasError = value
                OnPropertyChanged(New PropertyChangedEventArgs("hasError"))
            End Set
        End Property

        Public ReadOnly Property ListOfFunctions As List(Of String)
            Get
                Return mparser.LocalFunctions.Keys.ToList
            End Get
        End Property

        Public ReadOnly Property ListOfVars As List(Of String)
            Get
                Return mparser.LocalVariables.Keys.ToList
            End Get
        End Property

        Private Sub recalc()
            Try
                hasError = False
                Result = mparser.Parse(ToEvaluate)
            Catch
                hasError = True
            End Try
        End Sub

#Region "international"
        Public ReadOnly Property cultures As List(Of String)
            Get
                Dim cinfo As New List(Of String)
                For Each ci As CultureInfo In CultureInfo.GetCultures(CultureTypes.SpecificCultures)
                    cinfo.Add(ci.Name)
                Next
                Return cinfo
            End Get
        End Property

        Public Property selectedCultureName As String
            Get
                Return _culture.Name
            End Get
            Set(value As String)
                _culture = New CultureInfo(value)
                mparser.cultureInfo = _culture
                recalc()
                OnPropertyChanged(New PropertyChangedEventArgs("selectedCultureName"))
                OnPropertyChanged(New PropertyChangedEventArgs("selectedCulture"))
            End Set
        End Property

        Public ReadOnly Property selectedCulture As CultureInfo
            Get
                Return _culture
            End Get
        End Property
#End Region

        Private Sub OnPropertyChanged(ByVal e As PropertyChangedEventArgs)
            If Not PropertyChangedEvent Is Nothing Then
                RaiseEvent PropertyChanged(Me, e)
            End If
        End Sub
        Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged
    End Class


End Namespace
