Option Explicit
Sub WebScrolling()

    Dim Trash As Variant
    Dim IE As New InternetExplorer
    Dim Elts As IHTMLElementCollection
    Dim HtmlDoc As HTMLDocument

   'Chargement d'une page web Google
   'IE.Navigate "https://www.google.fr"
   IE.Navigate "https://fr.finance.yahoo.com/q/hp?s=AI.PA&a=00&b=01&c=1990&d=02&e=21&f=2016&g=d&z=66&y=0"
   'http://real-chart.finance.yahoo.com/table.csv?s=AI.PA&a=00&b=01&c=1990&d=02&e=21&f=2016&g=d&ignore=.csv

   'Affichage de la fenêtre IE
   IE.Visible = True
   WaitIE IE
   
   'IE.document.getElementById ("rightcoll")
   
   Set Elts = LooKingFor(IE.document)
   
   
   Dim L, C As Long: L = 1: C = 1
   Dim EltL, EltC As HTMLGenericElement
   For Each EltC In Elts.Item.all
    ThisWorkbook.Worksheets("Sortie").Cells(L, C).Value = EltC.innerText
    
    If EltC.all.Length > 0 Then
        For Each EltL In EltC.all
            ThisWorkbook.Worksheets("Sortie").Cells(L, C).Value = EltC.innerText
            C = C + 1
        Next EltL
        C = 1
    End If
    
    L = L + 1
   Next EltC
   
   
   'On libère la variable IE
   IE.Quit
   Set IE = Nothing

End Sub
Function LooKingFor(ByRef HtmlDoc As HTMLDocument) As Variant




End Function
