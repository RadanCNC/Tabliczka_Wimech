Imports System.Drawing

Public Class Tabliczka_Wimech
    Private Sub DXFtoNest_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    ' linia kodu odpowiadajaca za okno makra zawsze na wierzchu
    Private Sub frm_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Me.TopMost = True                                                                 'Zawsze na wierzchu
    End Sub

    Private Sub btnNie_Click(sender As Object, e As EventArgs) Handles btnNie.Click
        Me.Dispose()                                                                      'Zamknij okno
    End Sub

    Dim listaWarstw As New List(Of Pattern)

    'Poniżej funkcje skopiowane z API Radana

    '
    ' Check whether the given pattern exists.
    ' Sprawdź czy warstwa wogóle istnieje
    Public Function PatternExists(pattern As String) As Boolean


        ' Update all MAC variables by executing an empty command
        myMac.mac2("")

        ' Record the currently open pattern
        Dim oldOpenPattern As String
        oldOpenPattern = myMac.COP

        ' Try to open the pattern we want to check. If the
        ' pattern can be opened, then it exists. Patterns
        ' that do not exist cannot be opened.
        myMac.SetString("stringVal", pattern)
        PatternExists = myMac.mac2("\?\P,stringVal?o")

        ' Re-open the original open pattern.
        myMac.SetString("stringVal", oldOpenPattern)
        myMac.mac2("\?\P,stringVal?o")
    End Function

    'Funkcja otwórz warstwę
    Public Function PatternOpen(pattern As String) As Boolean

        ' Note: You can only open a pattern if it exists
        If PatternExists(pattern) Then
            myMac.SetString("stringVal", pattern)
            PatternOpen = myMac.mac2("\?\P,stringVal?o")
            ' DodajKomunikat(" otworzono warstwę " & pattern)
        Else
            PatternOpen = False
        End If
    End Function



    Public Function PobierzNazwyWarstw() As Boolean


        myMac.scan("/", "p", 0)
        Dim Ilosc As Integer = 0

        While (myMac.next())
            listaWarstw.Add(New Pattern() With {
                .Name = myMac.FT0,
                .Numer = Ilosc
                })
            Ilosc = Ilosc + 1
        End While
        For Each warstwa As Pattern In listaWarstw
            ' MsgBox(warstwa.Name)
        Next

        'MsgBox(myMac.FT0)

    End Function

    Private Sub BtnZapiszCzesci_Click(sender As Object, e As EventArgs) Handles BtnZapiszCzesci.Click


        'Do zapisu warstwy do pliku .sym potrzebne są poniższe dane
        Dim material As String = "MS"
        Dim grubosc As Double = 1
        Dim jednostki As String = "mm"
        Dim orientacja As Integer = 8

        myMac.SetString("material", material)
        myMac.SetNumber("grubosc", grubosc)
        myMac.SetString("jednostki", jednostki)
        myMac.SetNumber("orientacja", orientacja)

        Dim IleTabliczek As Integer = 0

        Try
            IleTabliczek = CInt(TBLinia4Do.Text) - CInt(TBLinia4Od.Text)
            If IleTabliczek <= 0 Then
                MsgBox("Wpisz poprawną ilość od do")
                Return
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        '  MsgBox(IleTabliczek)

        Dim SciezkaPliku As String

        Try
            '  myApp.Application.ActiveDocument.Close(True)


            For index As Integer = 1 To IleTabliczek + 1
                'Nowa część
                myApp.Application.ActiveDocument.Application.NewSymbol(True)

                myMac.PRS = ""
                myMac.MES = False

                'Rysowanie prostokąta
                myMac.mac2("'\!'")
                myMac.SetNumber("kolor", 7)
                myMac.mac2("e\?P,7?")
                myMac.UX = 100
                myMac.UY = 100
                myMac.mac2("s")
                myMac.UX = myMac.UX + 100
                myMac.UY = myMac.UY + 60
                myMac.mac2("\")
                myMac.rfmac(ControlChars.Quote)

                ZaokraglenieWierzcholka(101, 101, 5)
                ZaokraglenieWierzcholka(101, 159, 5)
                ZaokraglenieWierzcholka(199, 101, 5)
                ZaokraglenieWierzcholka(199, 159, 5)


                'Rysowanie tekstu
                myMac.mac2("'\!'")
                Dim NumerSeryjny As String = TBLinia4Prefix.Text + " " + TBLinia4Rok.Text + (CInt(TBLinia4Od.Text) + index - 1).ToString("D5")


                Dim NumerPiora As Integer = CInt(TBNumerPiora.Text)

                ' myMac.UX = 100
                ' myMac.UY = 100

                Dim PolozenieTekstuX As Integer = -100 + 24.4 + 4.6 + 25.624
                Dim PolozenieTekstuY As Integer = -60

                Dim IleZnakowWLinii As Integer

                Dim OilePrzesunac2 As Integer = TBLinia2.Text.Length - 10

                radDrawText(-100 + 24.4 + 4.6 + 25.624, -60 + 50 - 4 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia1.Text, 8)

                Dim text1 As String = TBLinia2.Text
                Dim arialBold As New Font(TextCzcionka.Text, 8)
                Dim textSize As Size = Windows.Forms.TextRenderer.MeasureText(text1, arialBold)
                'Dim textsize As Size = TextRenderer.MeasureText(cbx_Email.Text, cbx_Email.Font)
                ' MsgBox(textSize.Width.ToString)

                'PolozenieTekstuX = -100 + (100 - CInt(textSize.Width))
                'radDrawText(-100 + 14.9 + 4.6, -60 + 40 - 4, NumerPiora, "\ELaser sans Serif black\N" + TBLinia2.Text, 7)






                Select Case TBLinia2.Text.Length
                    Case < 10
                        radDrawText(PolozenieTekstuX, -60 + 40 - 4 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia2.Text, 8)
                    Case = 10
                        radDrawText(PolozenieTekstuX, -60 + 40 - 4 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia2.Text, 8)
                    Case < 14
                        radDrawText(PolozenieTekstuX, -60 + 40 - 4 + 15.5, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia2.Text, 7)
                    Case Else
                        radDrawText(PolozenieTekstuX, -60 + 40 - 4 + 14, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia2.Text, 6)
                End Select

                OilePrzesunac2 = TBLinia3.Text.Length - 10
                ' PolozenieTekstuX = -100 + 10.8 + 4.6

                Select Case TBLinia3.Text.Length
                    Case < 10
                        radDrawText(PolozenieTekstuX, -60 + 30 - 4 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia3.Text, 8)
                    Case = 10
                        radDrawText(PolozenieTekstuX, -60 + 30 - 4, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia3.Text, 8)
                    Case < 14
                        radDrawText(PolozenieTekstuX, -60 + 30 - 4 + 15.5, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia3.Text, 7)
                    Case Else
                        radDrawText(PolozenieTekstuX, -60 + 30 - 4 + 14, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia3.Text, 6)
                End Select


                ' radDrawText(-100 + 10.8 + 4.6, -60 + 30 - 4, NumerPiora, "\ELaser sans Serif black\N" + TBLinia3.Text, 7)
                radDrawText(PolozenieTekstuX, -60 + 20 - 4 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + NumerSeryjny, 8)
                radDrawText(PolozenieTekstuX, -60 + 6 + 17, NumerPiora, "\E" + TextCzcionka.Text + "\N\|" + TBLinia5.Text, 8)

                ' & index.ToString("D5")

                '  myMac.mac2("s")
                '  myMac.mac2("D")                     'Aby zapisać warstwę do pliku .sym trzeba nadać Datum
                '  myMac.mac2("\p")                    'Wchodzimy w tryb edycji warstwy

                SciezkaPliku = TBSciezkaDoZapisu.Text & "SN" & TBLinia4Rok.Text & (CInt(TBLinia4Od.Text) + index - 1).ToString("D5") & ComboBox1.Text
                'myMac.SetString("sciezka", AktualnaSciezkaPliku) 'sciezka do zapisu
                'myMac.mac2("\?s,sciezka,material,grubosc,jednostki,orientacja?")
                myApp.Application.ActiveDocument.SaveCopyAs(SciezkaPliku, "Radquote")
                ' MsgBox(index)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ZaokraglenieWierzcholka(ByVal punktX As Double, ByVal punktY As Double, ByVal promien As Integer)

        myMac.UX = punktX
        myMac.UY = punktY
        myMac.mac2("F")
        myMac.SetNumber("radius", promien)
        myMac.mac2("\?5,radius?")

    End Sub


    Public Sub radDrawText(ByVal punktX As Double, ByVal punktY As Double, ByVal pen As Integer, ByVal Tekst As String, ByVal charHeight As Integer)

        myMac.SetNumber("p1x", punktX)
        myMac.SetNumber("p1y", punktY)
        myMac.SetString("tem", Tekst)
        myMac.SetNumber("kolor", pen)
        myMac.SetNumber("cH", charHeight)
        myMac.SetString("font", "Stencil")
        myMac.rfmac("\?=,font?")
        myMac.mac2("\?-,cH?") ' zmiana rozmiaru czcionki
        myMac.mac2("\?T,tem?\?3,p1x,p1y? ") 'wpisz tekst
        myApp.Mac.scan(myApp.Mac.PCC_PATTERN_LAYOUT_TEXT, "t", 0) ' znajdź wstawiony tekst
        myMac.mac2("e\?P,kolor?") ' zmień kolor
        myMac.mac2("e\?)?") ' rozbij na linie i łuki
        myMac.profile_healing("/", True, TextBox1.Text, True, True, True, True, True)



        'and for other text settings : 
        'rfmac('\?-,aspectratio,charheight?') /* set the text aspect ratio &height with keystroke - */ 
        ' rfmac('\?=,font?') /* set the text font with keystroke =*/ 
        'rfmac('\?+,slant?') /* set the text slant with keystroke +*/ 
        'rfmac('\?_,txt_angle?') /* set the text orientation angle with keystroke _ */ 

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TBLinia2_TextChanged(sender As Object, e As EventArgs) Handles TBLinia2.TextChanged

    End Sub

    Private Sub Form1_Paint(ByVal sender As Object,
        ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint

        Dim myFont As New Font("Courier New", 8)
        Dim myFontBold As New Font("Microsoft Sans Serif", 10, FontStyle.Bold)
        Dim StringSize As New SizeF

        StringSize = e.Graphics.MeasureString("How wide is this string?", myFont)
        Debug.WriteLine("Height: " & StringSize.Height)
        Debug.WriteLine("Width: " & StringSize.Width)

        StringSize = e.Graphics.MeasureString("How wide is this string?", myFontBold)
        Debug.WriteLine("Height: " & StringSize.Height)
        Debug.WriteLine("Width: " & StringSize.Width)
    End Sub

    Private Sub TBLinia4Prefix_TextChanged(sender As Object, e As EventArgs) Handles TBLinia4Prefix.TextChanged

    End Sub

    Private Sub TBLinia5_TextChanged(sender As Object, e As EventArgs) Handles TBLinia5.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub TextCzcionka_TextChanged(sender As Object, e As EventArgs) Handles TextCzcionka.TextChanged

    End Sub

    Private Sub TBSciezkaDoZapisu_TextChanged(sender As Object, e As EventArgs) Handles TBSciezkaDoZapisu.TextChanged

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class

