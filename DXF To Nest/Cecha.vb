Public Class Cecha                                                                            'Cecha to np. linia, łuk, okrąg 

    'Każda linia i łuk składa się z dwóch punktów: początkowego i końcowego
    Public PunktPoczatkowyX As Double
    Public PunktPoczatkowyY As Double

    Public PunktKoncowyX As Double
    Public PunktKoncowyY As Double

    Public CzyJestToLiniaProsta As Boolean                                                    'Linia prosta - true, łuk - false
    Public CzyJestToPelenOkrag As Boolean                                                     'Pełny okrąg - true, łuk - false

    Public Identyfikator As String                                                            'Identyfikator z Radana
    Public LiczbaPorzadkowa As Integer

    Public LiniaSasiadujaca1 As String = "0"
    Public LiniaSasiadujaca2 As String


    'Jeżeli punkt końcowy jest taki sam jak początkowy to jest to pełen okrąg
    Sub SprawdzCzyToOkrag()
        If PunktPoczatkowyX = PunktKoncowyX And PunktPoczatkowyY = PunktKoncowyY Then
            CzyJestToPelenOkrag = True
        Else
            CzyJestToPelenOkrag = False
        End If
    End Sub


End Class
