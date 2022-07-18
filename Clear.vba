Sub clear() 
Selection.Copy 
Selection.PasteSpecial Paste:=xlValues, _ 
Operation:=xlNone, SkipBlanks:=False, Transpose:=False 
Application.CutCopyMode = False 
End Sub

'Source: https://www.poradykomputerowe.pl/aplikacje-biurowe/szybka-zamiana-formul-na-wartosci.html?cid=K000OG'