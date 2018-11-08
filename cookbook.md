<h1>Cookbook: Excel Carpentry for systematiske reviews.</h1>

**BEMÆRK!** Du kan åbne billeder med kommandoer i fuld størrelse, ved at klikke højre museknap og vælge 'Vis billede'. Så er det lettere at læse!

<h2>Formål:</h2>
Formålet med denne cookbok, er at du som læser skal kunne foretage simpel datamanipulation, for derigennem at sammenkøre store datasæt til brug i forbindelse med systematiske reviews, og lave en simpel visualisering af dette.

Trin 1: Sammenfald mellem de to datasæt
Trin 2: Differencen mellem de to datasæt
Trin 3: Visualisering ( VENN Diagram illustration )

De udleverede datasæt indeholder ca. 2000 poster fra Scopus og Web of Science, vi kender på nuværende tidspunkt ikke til eventuelle sammenfald og differencer mellem de to datasæt, og såfremt datasættene inspiceres vil det også være tydeligt, at formatet ikke er ens. Disse ting ønsker vi at ændre.

<h2>EXCEL (LOPSLAG) SÅDAN GØR DU!</h2>
Vi har anvendt Microsoft Excel med funktionen LOPSLAG til at samle vores data i et samlet regneark. Der er andre programmer og andre måder at opbevare data på (database). Vi viser dog alligevel her, hvordan du hurtigt kan kombinere data via en nøgle (DOI).

Du lavet et LOPSLAG ved at skrive følgende i formellinjen i Excel: =LOPSLAG(

Hvis du er nybegynder i Excel, kan det være en fordel at trykke på indsæt funktion. Du skal derefter udfylde følgende elementer i vinduet.
1. Opslagsværdi - dette er den værdi du vil bruge som nøgle i kombineringen af de to regneark (fællesnøglen). Du skal finde værdien i dit primære regneark, dvs. det regneark du vil arbejde med fremadrettet. Vi anbefaler at benytte værdier som DOI som fællesnøglen.

2. Tabelmatrix - dette er det dataområde, du vil gennemsøge for opslagsværdien. Dette er det regneark, som du vil indhente data fra til primærkarket. Det er vigtigt at den værdi du vil søge står i kolonne A.

3. Kolonneindeks_nr - dette er angivelsen af hvilket kolonnenummer i sekundærarket, der indeholder den værdu du er interesseret i at overføre.

4. Lig_med - dette er en angivelse af, om Excel såfremt en værdi ikke findes skal forsøge at finde en tilstødende værdi. Dette skal i vores tilfælde være FALSK, da vi ikke ønsker tilnærmelser af DOI, men kun præcise matches.

<h2>FREMGANGSMÅDE for fællesmængde og differens</h2>
Herunder et eksempel, hvor der skabes et match og overføres data fra et regneark til et andet.
Vi matcher på DOI, men der kunne også matches på ISSN for at se dækningen imellem de to datasæt på forlag. Såfremt dette ønskes, skal der dog foretages en eleminering af dubletter først, da det samme forlag ellers vil optræde mange gange.

1. Start med at importere wos.tsv i er Excel regneark. Dette gøres under fanebladet 'Data' -> 'Hent eksterne data' -> 'Fra tekst'. Naviger herefter til den mappe, hvor du har downloadet de to datasæt. Da endelsen på wos-filen er .tsv, er det nødvendigt at vise 'Alle filer' i filvisningen.
Filen er separeret med tabulering, hvorfor vi skal vælge 1) Afgrænset som den beskrivende filtype og 2) Tabulator.
De importerede data skulle nu stå tydeligt ordnet i individuelle kolonner.

2. Herefter gentages importen af scopus.csv i et nyt Excel dataark, filen er afgrænset med komma, hvorfor dette vælges i 'Afgrænsere'.
Dataen skal nu endnu en gang stå i klart separerede kolonner.

Vi er nu klar til at begynde vores arbejde mod at afgøre fællesmængde og difference i vores datasæt.
På grund af begrænsninger i Excel, er vi dog nød til at foretage en lille ændring i vores ene Excelark - dette skal altid gøres i det Excelark vi checker værdier op imod. Her skal checkværdien altid placeres i kolonne A (ellers virker det ikke).
Vi flytter derfor kolonnen indeholdende DOI i arket med Scopus-data til kolonne A (for nemhedens skyld gør vi det samme med WoS-dataen. Dette er ikke pækrævet, men gør arbejdet nemmere.Følg nu vejledningen i forhold til LOPSLAG i Excel.

**Find fællesmængden (A ∩ B) imellem WoS- og Scopusdatasæt:**
1. Sæt en kolonne ind på position B i WoS-arket.
2. I formellinjen skriver du =LOPSLAG(
3. Herefter kan du med fordel trykke på formelknappen (Fx) for at åbne formelvinduet.
4. Skriv herefter:
![LOPSLAG](/billeder/1.jpg)
5. Tryk på 'OK'
6. Træk herefter i nederste venstre hjørne af cellen indeholdende vores formel. Udvid indtil alle værdier i arket er sammenlignet (Se billede herunder).
![LOPSLAG](/billeder/2.jpg)
De celler som får en værdi, dvs. ikke er #I/T (Ikke tilgængelig) er de artikler som findes i begge datasæt, dvs. fællesmængden.
Vi kan stille skarpt på disse ved, at vælge et filter som fravælger #I/T (se herunder).
![FILTER](/billeder/3.jpg)

Hvis vi modset vil se, hvilke artikler som er repræsenteret i WoS-datasættet, men ikke i Scopus sætter vi samme filter på, men fravælger alle undtagen #I/T.
**Bemærk**, dette er ikke differencen. Vi har ved denne øvelse kun den ene del af differencen, da der kan være artikler i Scopus som ikke er i WoS-datasættet. Vi er derfor nød til at lave den omvendte øvelse fra Scopus -> WoS for at finde disse.
Gentag derfor ovenstående øvelse, men omvendt (se billede).
![LOPSLAG](/billeder/4.jpg)
Når du har fællesmængden og de to dele af differencen er du klar til at lave visualiseringen herunder.
