# VBA-project

Ez az Excel-VBA eszköz az ügyfélbefektetés adatainak teljes életciklusát kezeli: az adatok generálásától a hibajavításon át a vezetői riport elkészítéséig. A projekt a precizitást, adatintegritást és automatizált riportálást szemlélteti.

Főbb Funkciók
Automatizált Adatgenerálás: Egy gombnyomásra létrehoz egy 50+ soros CRM export szimulációt (régiók, eszközosztályok, értékesítési adatok).

Adatintegritási Motor: Automatikusan javítja a formázási hibákat (pl. kisbetűs nevek) és azonosítja a kritikus hiányzó adatokat.

Hibakezelés: A rendszer nem engedi a riport elkészítését, amíg a felhasználó ki nem javítja a pirossal megjelölt kritikus hibákat.

Dinamikus Dashboard: Azonnal összesíti a kezelt vagyont (AUM) régiók szerint.

Rendszer Reset: Lehetővé teszi a teljes környezet alaphelyzetbe állítását és a folyamat tiszta lappal való újrakezdését.

Használati útmutató
1. Lépés (Generate Data): Létrehozza a Raw_Data munkalapot és feltölti tesztadatokkal.

2. Lépés (Clean & Validate): Lefuttatja az ellenőrzést. A PIROSSAL jelölt sorokban pótolni kell a hiányzó összegeket!

3. Lépés (Generate Dashboard): Miután az adatok hibátlanok, elkészül a végső Dashboard riport.

Rendszer törlése (Reset System): Kitörli a generált adatokat és riportokat, visszaállítva az eredeti állapotot.

4. Technikai háttér
VBA (Makrók): Moduláris felépítés, hibakezeléssel és automatizált munkalap-menedzsmenttel.

Excel Logika: Dinamikus SUMIF képletek és automatikus tartománykezelés.
