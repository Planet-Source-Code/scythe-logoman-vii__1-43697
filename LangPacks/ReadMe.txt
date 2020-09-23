How to crate your own LanguagePack
-------------------------------------------------------------
Fisrt of all find out what id Ur language has.
U can use the LngTable.txt or u take a look at the About dialog of LogoMan
Now open a Languagefile and edit it.
Here is how.
The first line tells us what we have
# = Combobox
~ = Tabstrip
+ = ListView
^ = Control with Only Caption
° = Control with Only Tooltip
* = Control with Caption & ToolTipText 
The next line gives us the name of the control
The third line holds the index of this control

Never Change these 3 lines !!!!!!!!!!!!!!

Text to change comes from line 4 to ???

Combobox:
After the # comes a number 
This number show how many lines this combo has
#2 = 2 List Items in this combo (2 Text´s to translate)
Now change line 4 and 5
Sample
 #2
 Combo1
 0
 This is line 1
 This is line 2

German Version
 #2
 Combo1
 0
 Dies ist Zeile 1
 Dies ist Zeile 2

Tabstrip:
 Like Combobox but here we have 2 lines for each number
 Line1 Caption (Visible Text)
 Line2 TootTipText (The text u see after placing the Mouse over)

Listview:
 Same like Combobox

All others are easy (think so)

And dont forget to send me a copy of your Language Pack
Scythe@cablenet.de



