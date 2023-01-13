<?php 
// vstup musi byt v ANSI, neboli Windows-1250 // lze pozdeji rozsirit

// prvni radek jsou desetinna cisla urcujici sirku sloupcu (cca 46 je vychozi; empty string = vychozi)
// tagy ktere je zatim mozno pouzit
	// <b> - tucny text (v cele bunce)
	// <bg#hexdec> - hexadecimalni barva pozadi
	// <ce> - vycentrovani (vertikalni i horizontalni)
	// <border> - hranice na vsech 4 stranach - kdyz tento tag ma mene bunek, nefunguje spravne


class CustomXLSX
{
	public $xlsx_name = 'default.xlsx'; // nastavi se zvenci
	private $date_now = '';
	private $microtime_now = ''; // pro rozliseni slozky pro pripad vice requests najednou
	private $general_path = ''; // viz construct
	private $start_path = 'XLSXworkspace/';
	private $array_raw = []; // krome nastaveni nazvu souboru jedinny vstup zvenci
	private $array_clean = []; // PRAVDEPODOBNE SMAZAT
	private $tabulka = []; // sem se uozi vsechna data ziskana z $array_raw + data ziskana zpracovanim $array_raw
	private $sirky_sloupcu = [];
	
	
	function __construct()
	{
		$this->date_now = gmdate('D, d M Y H:i:s \G\M\T' , time() );
		$this->microtime_now = microtime(true);
		// $this->temp_path = sys_get_temp_dir() . '/' . $microtime_now; // moznost pracovat v temp po prideleni povoleni
		$this->general_path = $this->start_path . $this->microtime_now;
	}
	
	
	public function Input(array $array_raw) // zpracovat ziskany array
	{
		function zmenit_encoding(&$pole)
		{
			$pole = iconv('Windows-1250', 'UTF-8', $pole);
		}
		array_walk_recursive($array_raw, 'zmenit_encoding'); // zmeni encoding u kazdeho pole
		
		$this->array_raw = $array_raw;
		
		$this->Zpracovat();
		
		
		$this->Rels();
		$this->Doc_props();
		$this->Xl();
		$this->Content_types();
		
		$this->Zip();
	}
	
	
	private function Zpracovat() // ziskat styly z tagu // ocistit array od tagu
	{
		$this->sirky_sloupcu = array_shift($this->array_raw);
		$array_work = $this->array_raw; //$array_raw pozdeji zbavit tagu. Zaroven tagy
		
		$abc_vzor = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];
		$abc = $abc_vzor;
		$abc_ext = 0; // kolikrat bylo $abc prodlouzeno // hodnota oznacuje pismeno, ktere se prida pred hodnoty $abc
		while(count(max($array_work)) > count($abc)) // dokud nejdelsi radek je delsi nez $abc: automaticky prodlouzit $abc
		{
			$cislo_pismene = 0;
			foreach($abc_vzor as $pismeno)
			{
				array_push($abc, $abc_vzor[$abc_ext] . $abc_vzor[$cislo_pismene]);
				$cislo_pismene++;
			}
			$abc_ext++;
		} 
		
		// vytvorit asociativni multidimenzionalni pole: tabulka=>radek=>bunka=>(styl_0, styl_1, styl_2)
		$cislo_bunky = 0;
		$cislo_radku = 0;
		$cislo_stylu_vychozi = 1;
		$this->tabulka = [];
		foreach($array_work as $radek) // iterovat po radcich
		{
			$cislo_sloupce = 0;
			
			array_push($this->tabulka, array()); // nove prazdne pole reprezentujici radek; bude obsahuvat bunky
			foreach($radek as $bunka) // iterovat po bunkach
			{
				$nove_styly = []; // pridaji se do nej nove styly // zjisti se, zda je mezi starymi styly // prida se kazdopadne do $this->tabulka[$cislo_radku][$cislo_sloupce]['styly']
				
				$oznaceni = $abc[$cislo_sloupce] . ($cislo_radku + 1);
				$cislo_sloupce++;
				
				$this->tabulka[$cislo_radku][$cislo_sloupce]['oznaceni'] = $oznaceni;
				$this->tabulka[$cislo_radku][$cislo_sloupce]['novy_styl'] = false; // aby se poznalo, jestli se ma do styles.xml zapsat novy styl
				
				$cislo_stylu = false; //pokud bunka nezacina tagem, zustane jako vychozi
				
				// odstranit tagy a zapsat styly do pole pridaneho vyse // 
				while(substr($bunka, 0, 1) === '<') // dokud $bunka zacina tagem
				{
					$cislo_stylu = true;
					$delka_bunky = strlen($bunka);
					$konec_tagu = strpos($bunka, '>');
					switch(substr($bunka, 1, 2)) // co je druhy a treti znak - prvni a druhy znak od zacatku tagu
					{
						case 'b>': // pridat tucny styl konkretni bunce
							$nove_styly['bold'] = true;
							break;
						case 'bg': // pridat barvu pozadi konkretni bunce
							$hex = 'FF' . substr($bunka, 4, 6); // ziskat barvu v hexadecimalnim formatu
							$nove_styly['barva_pozadi'] = $hex;
							break;
						case 'ce': // vycentrovat vertikalne i horizontalne
							$nove_styly['vycentrovat'] = true;
							break;
						case 'bo': // hranice // nefunguje - je u vice bunek nez ma byt
							// $hex = substr($bunka, 4, 6);
							// $nove_styly['hranice'] = $hex;
							$nove_styly['hranice'] = true;
							break;
						default: // hodit error
							die('Error: Neznamy tag na zacatku bunky ' . $oznaceni); // DOPLNIT CISLO BUNKY (napriklad C3)
					}
					$bunka = substr($bunka, $konec_tagu + 1, $delka_bunky); // odebrat prvni tag
				}
				// prohledat, jestli $nove_styly uz nejsou totozne zapsany; kdyz narazi na shodu, ukonci foreach
				if($cislo_stylu === true) // pokud jsou u aktualni bunky styl(y) == tagy
				{
					$this->tabulka[$cislo_radku][$cislo_sloupce]['novy_styl'] = true;
					$cislo_stylu = $cislo_stylu_vychozi;
					$cislo_stylu_vychozi++;
					foreach($this->tabulka as $radek_t)
					{
						foreach($radek_t as $bunka_t)
						{
							if($bunka_t['styly'] === $nove_styly) // pokud jsou $nove_styly jiz obsazeny v $this->tabulka
							{
								$cislo_stylu_vychozi--;
								$cislo_stylu = $bunka_t['cislo_stylu']; // kdyz je styl stejny, prevzit cislo stylu
								$this->tabulka[$cislo_radku][$cislo_sloupce]['novy_styl'] = false;
								break 2; // vyskocit ze dvou foreach
							}
						}
					}
				}
								
				$this->tabulka[$cislo_radku][$cislo_sloupce]['cislo_stylu'] = $cislo_stylu;
				$this->tabulka[$cislo_radku][$cislo_sloupce]['styly'] = $nove_styly;
				$this->tabulka[$cislo_radku][$cislo_sloupce]['text'] = $bunka;
				
			}
			$cislo_radku++;
		}
		file_put_contents('XLSXworkspace/log-tabulka.txt', print_r($this->tabulka, true));
		
		$this->tabulka = $this->tabulka;
		// nahradit polem $this->tabulka!!!
		$this->array_clean = $array_work; // zapsat do promenne, co bude pouzita jinymi metodami
	}
	
	
// ROOT SLOZKA 

	private function Rels() // SLOZKA (1 soubor: .rels) 
	{		
		$nazev_souboru = '.rels';
		$path = $this->general_path . '/unpacked/_rels/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
		// file_put_contents('XLSXworkspace/log.txt', $this->date_now . '---' . $path, FILE_APPEND . PHP_EOL);
	}
	
	
	private function Doc_props() // SLOZKA
	{
		$this->App_xml();
		$this->Core_xml();
	}
	
	
	private function Xl() // SLOZKA // doplnit
	{
		// slozky
		$this->Rels_workbook_xml_rels();
		$this->Theme_theme1_xml();		
		$this->Worksheets_sheet1_xml();
		
		//soubory
		$this->Shared_strings_xml();
		$this->Styles_xml();
		$this->Workbook_xml();
	}
	
	
	private function Content_types() // SOUBOR
	{
		$nazev_souboru = '[Content_Types].xml';
		$path = $this->general_path . '/unpacked/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
	}
	
	
// JEDNOTLIVE SOUBORY V PODSLOZKACH /////////
// // docProps ////
	private function App_xml()
	{
		$nazev_souboru = 'app.xml';
		$path = $this->general_path . '/unpacked/docProps/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>List1</vt:lpstr></vt:vector></TitlesOfParts><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
	}
	
	
	private function Core_xml() // pripadne zmenit datum a cas zmeny ($text_souboru)
	{
		$nazev_souboru = 'core.xml';
		$path = $this->general_path . '/unpacked/docProps/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>stepan.hampl</dc:creator><cp:lastModifiedBy>Štěpán Hampl</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2015-06-05T18:19:34Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2022-03-08T19:24:34Z</dcterms:modified></cp:coreProperties>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
	}
	
	
// // xl /////
// // // _rels
	private function Rels_workbook_xml_rels()
	{
		$nazev_souboru = 'workbook.xml.rels';
		$path = $this->general_path . '/unpacked/xl/_rels/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path);
	}
	
	
// // // theme
	private function Theme_theme1_xml()
	{
		$nazev_souboru = 'theme1.xml';
		$path = $this->general_path . '/unpacked/xl/theme/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="5B9BD5"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="4472C4"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path);
	}


// // // worksheets
	private function Worksheets_sheet1_xml()
	{
		$this->tabulka = $this->tabulka;
		$nazev_souboru = 'sheet1.xml';
		
		$col_poradi = 0;
		$cols = '<cols>';
		foreach($this->sirky_sloupcu as $sloupec)
		{
			$col_poradi++;
			if($sloupec)
			{
				$cols .= '<col min="' . $col_poradi . '" max="' . $col_poradi . '" width="' . $sloupec . '" customWidth="1"/>';
			}
		}
		$cols .= '</cols>';
		
		
		$path = $this->general_path . '/unpacked/xl/worksheets/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{F111F337-430F-4A76-9309-8543E9522614}">
<dimension ref="A1:A1"/>
<sheetViews>
<sheetView tabSelected="1" workbookViewId="0"/>
</sheetViews>
<sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.3"/>
' . $cols . '
<sheetData>';
		
		$cislo_bunky = 0;
		foreach($this->tabulka as $cislo_radku => $radek)
		{
			$text_souboru .= '
<row r="' . ($cislo_radku + 1) . '" spans="1:1" x14ac:dyDescent="0.3">';
			foreach($radek as $bunka)
			{
				if($bunka['cislo_stylu']) // pokud je vyplneno // pokud neni false
				{
					$cislo_stylu = ' s="' . $bunka['cislo_stylu'] . '"';
				}
				else
				{
					$cislo_stylu = '';
				}
				
				$obsah = $cislo_bunky; // zustane, pokud je $bunka string
				if(is_numeric($bunka['text'])) // pokud je bunka cislo
				{
					$string_info = '';
					$obsah = $bunka['text'];
				}
				else
				{
					$string_info = ' t="s"';
					$cislo_bunky++; 
				}
				
				$text_souboru .= '
	<c r="' . $bunka['oznaceni'] . '"' . $cislo_stylu . $string_info . '><v>' . $obsah . '</v></c>';
			}
			$text_souboru .= '
</row>';
		}
		
		$text_souboru .= '
</sheetData>
<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>';

	$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
	}
	
// // OSTATNI SOUBORY ve slozce xl /////

	private function Shared_strings_xml() // ze zpracovaneho pole (bez tagu) vytvori obsah XML souboru
	{
		$this->tabulka = $this->tabulka;
		$nazev_souboru = 'sharedStrings.xml';
		$path = $this->general_path . '/unpacked/xl/';
		$text_souboru1 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'; // veskery obsah souboru
		
		$pocet_radku = count($this->tabulka);
		$pocet_bunek = 0;
		foreach($this->tabulka as $radek)
		{
			foreach($radek as $bunka)
			{
				if(!is_numeric($bunka['text'])) // pokud bunka neobsahuje cislo
				{
					$text_souboru2 .= '
	<si><t>' . $bunka['text'] . '</t></si>';
					$pocet_bunek++;
				}
			}
		}
		$text_souboru2 .= '
</sst>';
		$text_souboru1 .= '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $pocet_bunek . '" uniqueCount="' . $pocet_bunek . '">';
		$text_souboru = $text_souboru1 . $text_souboru2;
		$this->Gen_xml($nazev_souboru, $text_souboru, $path); // vytvorit soubor
	}
	
// // // styles.xml
	private function Styles_xml()
	{
		$barvy = []; // $fill_id => $hex // takto by muselo byt vsude, kde to neni binarni
		$nazev_souboru = 'styles.xml';
		$path = $this->general_path . '/unpacked/xl/';
		
		$cellXfs_count = 1;
		
		$fonts_count = 1;
		$dalsi_fonty = '';
		$fills_count = 2;
		$fill_id = 2;
		$border_id = 0;
		
		$jiz_hranice = false;
		
		$dalsi_fills = ''; 
		$apply_alignment = '/>';
		$dalsi_borders = '';
		
		foreach($this->tabulka as $radek)
		{
			foreach($radek as $bunka)
			{
				$fill_id_nove = 0; // vychozi - pokud nema ['barva_pozadi'], zustava
				$font_id = 0; // po pridani dalsich stylu pridat i dalsi Id
				$fill_id_nove = 1;
				if($bunka['novy_styl']) // pokud ma bunka styl, ktery se jeste nezapsal
				{
					if($bunka['styly']['bold']) // pokud tag bold neni false
					{
						$font_id = 1; // zde to lze nastavit jednoduse. U barev pozadi zadat do array $poradi_stylu => $hexadecimalni_barva
						$dalsi_fonty_nove = '
		<font>
			<b/>
			<sz val="11"/>
			<color theme="1"/>
			<name val="Calibri"/>
			<family val="2"/>
			<charset val="238"/>
			<scheme val="minor"/>
		</font>';
						if(!strpos($dalsi_fonty, $dalsi_fonty_nove)) // pokud neni jeste zapsany tento styl (font)
						{
							$fonts_count++;
							$dalsi_fonty .= $dalsi_fonty_nove;
						}
					}
					if($bunka['styly']['barva_pozadi']) // pokud ma definovanou barvu pozadi
					{
						$hex = $bunka['styly']['barva_pozadi'];
						$dalsi_fills_nove = '
			<fill>
				<patternFill patternType="solid">
					<fgColor rgb="' . $hex . '" />
					<bgColor indexed="64" />
				</patternFill>
			</fill>';
						$fill_id_nove = $fill_id;
						$fill_id_stare = array_search($hex, $barvy); // prislusne id jiz pouzite barvy nebo false
						if($fill_id_stare) // pokud je jiz zapsany tento styl (tato barva pozadi)
						{
							$fill_id_nove = $fill_id_stare;
						}
						else
						{
							$fill_id_nove = $fill_id;
							$barvy[$fill_id_nove] = $hex; // zapsat do pole pro pripad pouziti stejne barvy v budoucnu
							$fill_id++;
							$dalsi_fills .= $dalsi_fills_nove;
							$fills_count++;
						}
						
					}
					else // pokud nema definovanou barvu pozadi
					{
						$fill_id_nove = 0;
					}

					$apply_alignment_nove = $apply_alignment;
					if($bunka['styly']['vycentrovat'])
					{
						$apply_alignment_nove = ' applyAlignment="1">
			<alignment horizontal="center" vertical="center"/>
		</xf>';
					}
					
					if($bunka['styly']['hranice'] && $jiz_hranice) // pokud bunka obsahuje tag pro hranici a zaroven jiz existuje tento styl (hranice) ve styles.xml
					{
						$border_id = 1;
					}
					elseif($bunka['styly']['hranice']) // pokud jen bunka obsahuje tag pro hranici, ale tento styl se aplikuje poprve
					{
						$border_id = 1;
						$dalsi_borders = '
		<border>
			<left style="thin">
				<color indexed="64"/>
			</left>
			<right style="thin">
				<color indexed="64"/>
			</right>
			<top style="thin">
				<color indexed="64"/>
			</top>
			<bottom style="thin">
				<color indexed="64"/>
			</bottom>
			<diagonal/>
		</border>';
					}
					$dalsi_xf .= '<xf numFmtId="0" fontId="' . $font_id . '" fillId="' . $fill_id_nove . '" borderId="' . $border_id . '" xfId="0"' . $apply_alignment_nove . '
		';
				}
			}
		}
		
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
	<fonts count="' . $fonts_count . '" x14ac:knownFonts="1">
		<font>
			<sz val="11"/>
			<color theme="1"/>
			<name val="Calibri"/>
			<family val="2"/>
			<scheme val="minor"/>
		</font>
		' . $dalsi_fonty . '
	</fonts>
		<fills count="' . $fills_count . '">
			<fill>
				<patternFill patternType="none"/>
			</fill>
			<fill>
				<patternFill patternType="gray125"/>
			</fill>
			' . $dalsi_fills . '
		</fills>
	<borders count="1">
		<border>
			<left/>
			<right/>
			<top/>
			<bottom/>
			<diagonal/>
		</border>
		' . $dalsi_borders . '
	</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
	<cellXfs count="1">
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
		' . $dalsi_xf . '
	</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>';
		
		file_put_contents('XLSXworkspace/styles.txt', $text_souboru);
		
		$this->Gen_xml($nazev_souboru, $text_souboru, $path);
	}
	
// // // workbook.xml
	private function Workbook_xml()
	{
		$nazev_souboru = 'workbook.xml';
		$path = $this->general_path . '/unpacked/xl/';
		$text_souboru = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="24931"/><workbookPr/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="https://maximareality-my.sharepoint.com/personal/stepan_hampl_maxima_cz/Documents/Documents/XLSX-research/blank-autoformatted2/" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="1" documentId="13_ncr:1_{EFEA4D14-88C5-4DB3-90ED-F4ADCC91AE68}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{8F7E8052-F85D-4D9A-BADB-74584354E242}"/><bookViews><workbookView xWindow="6300" yWindow="-13995" windowWidth="17280" windowHeight="9060" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets><sheet name="List1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="162913"/><extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:workbookPr chartTrackingRefBase="1"/></ext></extLst></workbook>';
		$this->Gen_xml($nazev_souboru, $text_souboru, $path);
	}


// ZAPSANI SOUBORU na disk (server)
	private function Gen_xml($nazev_souboru, $text_souboru, $path) // z $text_souboru vytvori XML
	{
		// $path = $this->general_path . '/unpacked/xl/';
		
		if(!file_exists($path))
		{
			mkdir($path, 0777, true); // vytvori unikatni slozku v temporary
		}
		
		$novy_soubor = fopen($path . $nazev_souboru, 'w') or die('Nelze otevrit soubor.'); // vytvorit novy soubor
		fwrite($novy_soubor, $text_souboru); // zapsat do noveho souboru (melo by se ulozit)
		$size = ftell($novy_soubor); // zjistit velikost noveho souboru
		fclose($novy_soubor);
	}
	
	
// VYTVORIT ZIP, stahnout ho a SMAZAT vse vytvorene
	private function Zip() // stahne veskery obsah vytvorene slozky jako ZIP (XLSX)
	{
		$path_files = realpath($this->general_path . '/unpacked/'); // cesta k souborum do zipu
		$path_zip = $this->general_path . '/';
		$novy_soubor = $path_zip . $this->xlsx_name;
		
		$zip = new ZipArchive();
		$zip->open($novy_soubor, ZipArchive::CREATE | ZipArchive::OVERWRITE);

		// Create recursive directory iterator
		$files = new RecursiveIteratorIterator(
			new RecursiveDirectoryIterator($path_files),
			RecursiveIteratorIterator::LEAVES_ONLY
		);

		foreach ($files as $name => $file)
		{
			// Preskocit slozky (pridany automaticky)
			if (!$file->isDir())
			{
				// Get real and relative path for current file
				$filePath = $file->getRealPath();
				$relativePath = substr($filePath, strlen($path_files) + 1);

				// Add current file to archive
				$zip->addFile($filePath, $relativePath);
			}
		}
		
		$zip->close();
		
		// header('Content-Type: application/zip');
        // header('Content-Disposition: attachment; filename= ' . $this->xlsx_name);
        // header('Last-Modified: ' . $date_now);
		
		header('Content-Disposition: attachment; filename=' . $this->xlsx_name);	
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
		header('Content-Length: ' . filesize($novy_soubor));
		header('Content-Transfer-Encoding: binary');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		readfile($novy_soubor);
		
		function odstranit_slozku($path) // odstrani slozku i s obsahem
		{
			// file_put_contents('XLSXworkspace/log.txt', $this->date_now . '---' . $path . PHP_EOL, FILE_APPEND);
			
			$ke_smazani = glob($path . '/{,.}[!.,!..]*', GLOB_MARK|GLOB_BRACE); // vcetne skrytych souboru
			foreach($ke_smazani as $soubor)
			{
				is_dir($soubor) ? odstranit_slozku($soubor) : unlink($soubor);
			}
			rmdir($path);
		}
		
		odstranit_slozku($this->general_path); // odstrani veskere vytvorene soubory
	}
}
