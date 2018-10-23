<h3>EXCEL (LOPSLAG)</h3>
Vi har anvendt Microsoft Excel med funktionen LOPSLAG til at samle vores data i et samlet regneark. Der er andre programmer og andre måder at opbevare data på (database). Vi viser dog alligevel her, hvordan du hurtigt kan kombinere data via en nøgle (DOI).

Du lavet et LOPSLAG ved at skrive følgende i formellinjen i Excel: =LOPSLAG(

Hvis du er nybegynder i Excel, kan det være en fordel at trykke på indsæt funktion. Du skal derefter udfylde følgende elementer i vinduet.
1. Opslagsværdi - dette er den værdi du vil bruge som nøgle i kombineringen af de to regneark (fællesnøglen). Du skal finde værdien i dit primære regneark, dvs. det regneark du vil arbejde med fremadrettet. Vi anbefaler at benytte værdier som DOI som fællesnøglen.

2. Tabelmatrix - dette er det dataområde, du vil gennemsøge for opslagsværdien. Dette er det regneark, som du vil indhente data fra til primærkarket. Det er vigtigt at den værdi du vil søge står i kolonne A.

3. Kolonneindeks_nr - dette er angivelsen af hvilket kolonnenummer i sekundærarket, der indeholder den værdu du er interesseret i at overføre.

4. Lig_med - dette er en angivelse af, om Excel såfremt en værdi ikke findes skal forsøge at finde en tilstødende værdi. Dette skal i vores tilfælde være FALSK, da vi ikke ønsker tilnærmelser af DOI, men kun præcise matches.

Herunder et eksempel, hvor der skabes et match og overføres data fra et regneark til et andet. Vi matcher på DOI, som findes i kolonne 'E' i primær regnearket (DOI skal altid findes i kolonne A i det regneark vi matcher med, ellers virker det ikke). I tabelmatrix definerer vi hele det sekundære ark som vores søgested. Vi er interesseret i at overføre værdien i 4-kolonne, dette angiver vi med tallet '4'. Sidst angiver vi, at vi kun ønsker et præcist match mellem DOI i kolonne E og A, dette gøres ved at angive FALSK i feltet.
