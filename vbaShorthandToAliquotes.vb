Sub shorthandToAliquots()


    ' shorthand_to_Aliquots Macro

    Dim sel As Range
    Set sel = Selection.Range
    
    ' Define the dictionary mapping shorthand to verbose
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    ' all the straight quarters
    dict.Add "NE", "NE¼"
    dict.Add "SE", "SE¼"
    dict.Add "NW", "NW¼"
    dict.Add "SW", "SW¼"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' all the straight halfs
    dict.Add "N2", "N½"
    dict.Add "S2", "S½"
    dict.Add "E2", "E½"
    dict.Add "W2", "W½"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' all the halfs in the NE4
    dict.Add "N2NE", "N½NE¼"
    dict.Add "S2NE", "S½NE¼"
    dict.Add "E2NE", "E½NE¼"
    dict.Add "W2NE", "W½NE¼"
    
    ' all the quarters_quarters in the NE4
    dict.Add "NENE", "NE¼NE¼"
    dict.Add "NWNE", "NW¼NE¼"
    dict.Add "SWNE", "SW¼NE¼"
    dict.Add "SENE", "SE¼NE¼"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' all the halfs in the NW4
    dict.Add "N2NW", "N½NW¼"
    dict.Add "S2NW", "S½NW¼"
    dict.Add "E2NW", "E½NW¼"
    dict.Add "W2NW", "W½NW¼"
    
    ' all the quarters_quarters in the NW4
    dict.Add "NENW", "NE¼NW¼"
    dict.Add "NWNW", "NW¼NW¼"
    dict.Add "SWNW", "SW¼NW¼"
    dict.Add "SENW", "SE¼NW¼"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' all the halfs in the SW4
    dict.Add "N2SW", "N½SW¼"
    dict.Add "S2SW", "S½SW¼"
    dict.Add "E2SW", "E½SW¼"
    dict.Add "W2SW", "W½SW¼"
                
    ' all the quarters_quarters in the SW4
    dict.Add "NESW", "NE¼SW¼"
    dict.Add "NWSW", "NW¼SW¼"
    dict.Add "SWSW", "SW¼SW¼"
    dict.Add "SESW", "SE¼SW¼"
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' all the halfs in the SE4
    dict.Add "N2SE", "N½SE¼"
    dict.Add "S2SE", "S½SE¼"
    dict.Add "E2SE", "E½SE¼"
    dict.Add "W2SE", "W½SE¼"
                 
    ' all the quarters_quarters in the SE4
    dict.Add "NESE", "NE¼SE¼"
    dict.Add "NWSE", "NW¼SE¼"
    dict.Add "SWSE", "SW¼SE¼"
    dict.Add "SESE", "SE¼SE¼"
    
    ' Iterate through the selected text
    Dim word As Variant
    Dim textArray As Variant
    Dim result As String
    
    textArray = Split(sel.Text, " ")
    result = ""
    
    For Each word In textArray
        If dict.Exists(word) Then
            result = result & dict(word) & " "
        Else
            result = result & word & " "
        End If
    Next word
    
    ' Replace the selected text with the converted text
    sel.Text = Trim(result)

'
End Sub
