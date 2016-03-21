Sub WaitIE(IE As InternetExplorer)
   'On boucle tant que la page n'est exploitable
   Do Until IE.readyState = READYSTATE_COMPLETE
      DoEvents
   Loop
End Sub
