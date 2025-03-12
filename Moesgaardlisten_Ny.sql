USE LOIS

SELECT DISTINCT
	TRIM(sag.Sagsnummer) 'Sagsnummer'
	,convert(date,sag.Sagsdato) 'Sagsdato'
	,TRIM(info.Sagsart) 'Sagsart'
	,REPLACE(ISNULL(info.Sagsadresse,''),'  ',' ') 'Adresse'
	,TRIM(sag.Titel) 'Titel'



FROM
	Service_NOVA2.LIS_Byg_sager sag
	LEFT JOIN Service_NOVA2.LIS_Byg_info info ON sag.SagsUUID = info.SagsUUID
	LEFT JOIN Service_NOVA2.LIS_Byg_Brugerfelter brugerfelt ON sag.SagsUUID = brugerfelt.SagsUUID

WHERE
	sag.Sagsdato >= '@datoFra'
	AND sag.Sagsdato <= '@datoTil'
	AND sag.KLE_nummer NOT IN ('02.00.20','01.02.05','01.04.00')
	AND (brugerfelt.Vaerdi NOT IN ('Aktindsigt','Juridisk','Adresser') OR brugerfelt.SagsUUID IS NULL)
	AND Sagsart NOT IN ('Stikprøvekontrol','Aktindsigt','Klagesager JUR','Klagesager BYG')
	AND Sagsart NOT LIKE 'Henvist til ny ansøgningstype i BOM%'
	AND TRIM(sag.Titel) <> 'Plansager'
	AND sag.Titel NOT LIKE '%HENVIST TIL%ANSØGNINGSTYPE%'
	AND sag.Titel NOT LIKE '%Fejloprettet%'
	AND sag.Titel NOT LIKE '%testsag%'
	AND TRIM(sag.Titel) NOT LIKE 'BBR Erhvervsindsats'
	AND Sagsadresse IS NOT NULL

ORDER BY
	convert(date,sag.Sagsdato)


