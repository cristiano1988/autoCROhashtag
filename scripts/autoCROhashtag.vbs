Set wshshell = wscript.CreateObject("WScript.Shell")
strtimes = inputbox ("Unesite broj komentara.")
strtimeneed = inputbox ("Unesite broj sekundi.")
strspeed = 300

If not isnumeric (strtimes & strtimeneed) then
    msgbox "Unesena neispravna vrijednost broja sekundi ili komentara. Gasenje."
    wscript.quit
End If

strtimeneed2 = strtimeneed * 1000

do
    msgbox "Imate " & strtimeneed & " sekundi za postavljanje kursora u prostor za objavu."
    wscript.sleep strtimeneed2
    
    for i=0 to strtimes
    wshshell.sendkeys ("#CRO #OrangeSponsorsYou" & "{enter}")
    wscript.sleep strspeed
    Next

    wscript.sleep strspeed * strtimes / 10
    returnvalue = MsgBox ("Zelite li ponoviti komentiranje?", 36)

    If returnvalue = 6 Then
    Msgbox "autoCROhashtag ce se aktivirati ponovno."
    End If

    If returnvalue = 7 Then
        msgbox "Gasenje."
        wscript.quit
    End IF
loop