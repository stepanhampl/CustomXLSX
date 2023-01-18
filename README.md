# CustomXLSX

Jedná se o mé řešení, jak s pomocí PHP scriptu vytvořit XLSX soubor a uložit jej uživateli mezi stažené soubory. 

CustomXLSX.php byl napsán v PHP 5.5.38. Zjevně tedy není kompatibilní mapříklad s PHP 7.4, takže jsem ho na svém počítači nedovedl otestovat - PHP 5.5.38 již není oficiálně podporováno, takže se mi ho nepovedlo nainstalovat. Nicméně na serveru byl soubor CustomXLSX.php úspěšně otestován.

Použití je znázorněno v souboru pouziti.php. Ani ten není otestován, ale teoreticky by měl takto fungovat. CustomXLSX->Input() přijímá dvoudimenzionální pole, přičemž každé vnitřní pole reprezentuje řádek (s výjimkou prvního - to může obsahovat jednotlivé šířky buněk). Každá jednotlivá buňka může obsahovat tagy.

Tagy, ktere je mozno pouzit (vloží se na začátek buňky v libovolném pořadí) - každý je zvlášť ohraničen "zobáčky" jako HTML tagy:
 - b - tučný text (v celé buňce)
 - bg#hexdec - hexadecimální barva pozadí
 - ce - vycentrování (vertikální i horizontální)
 - border - hranice na všech 4 stranach - když je tento tag obsažen v menším množstvím buňek, nefunguje správně

JAK TO FUNGUJE
  CustomXLSX.php založí na serveru složku, v níž vytvoří strukturu složek se soubory obsahující XML. Tato struktura složek se zabalí jako zip, nicméně s koncovkou XLSX. Prohlížeč soubor následně uloží mezi stažené soubory a strukturu složek ze serveru smaže. Tento soubor XLSX je funkční (alespoň byl při testech na serveru).
