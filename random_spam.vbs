Randomize

answer = MsgBox("Ez egy poén script, ami ablakokat fog nyitogatni és a creator  és élvezd hogy belle hal a géped és semi felelöséget nem válalok semi damage ért . Indítsam el?", vbYesNo + vbExclamation, "Figyelmeztetés")

If answer = vbYes Then
    Set shell = CreateObject("WScript.Shell")
    
    Do
        r = Int((3 * Rnd) + 1) ' 1-től 3-ig véletlen szám

        Select Case r
            Case 1
                shell.Run "notepad"
            Case 2
                shell.Run "cmd"
            Case 3
                shell.Run "mspaint" ' Paint is vicces :D
        End Select

        WScript.Sleep Int((300 + Rnd * 700)) ' Vár 300-1000 ms között

    Loop
Else
    MsgBox "A program leállt program stoped.", vbInformation, "Kilépés"
End If
