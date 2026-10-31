# ExcelToWord

Windows asztali alkalmazás, amely egy Excel-táblázat soraiból **személyre szabott
Word-dokumentumokat** készít egy Word-sablon alapján. Az eredeti felhasználás:
tartozáslistából ügyfelenkénti fizetési felszólító levelek generálása, kézi
másolgatás helyett.

**Stack:** C# · .NET 7 (Windows Forms) · EPPlus (Excel olvasás) · Open XML SDK (Word írás)

## Hogyan működik

1. Kiválasztod az Excel-fájlt, a Word-sablont és a célmappát.
2. A program a táblázat 4. sorától kezdve beolvassa a tételeket, és ügyfél (név)
   szerint csoportosítja őket, így egy ügyfél több tartozása egy levélbe kerül.
3. Minden ügyfélhez lemásolja a sablont, kitölti a név, cím és tételadatok
   helyét (bizonylatszám, jogcím, összeg, fizetési határidő), és külön
   `.docx` fájlba menti a célmappában (`<ügyfél neve>.docx`).

### Helyőrzők a Word-sablonban

| Helyőrző | Mire cseréli |
| --- | --- |
| `Name` | ügyfél neve |
| `City`, `Street` | ügyfél címe |
| `AllDebt` | összes tartozás |
| `AllDebtWithInterest` | összes tartozás kamattal |
| `Item` | a tételsor helye: ide kerül tételenként a bizonylatszám, jogcím, összeg és határidő |

## Futtatás

**Előfeltételek:** Windows, .NET 7 SDK (vagy Visual Studio 2022 / Rider).

```bash
git clone https://github.com/Kbence9/ExcelToWord.git
cd ExcelToWord/ExcelToWord
dotnet run --project ExcelToWord
```

Vagy nyisd meg az `ExcelToWord/ExcelToWord.sln` fájlt, és indítsd onnan.

## Megkötések

- Az Excel oszlopkiosztása (melyik oszlopban van a név, összeg, határidő stb.) jelenleg
  a kódban rögzített (`Services/ConvertExcelToWord.cs`), egy adott táblaformátumhoz
  készült.
- Az EPPlus nem kereskedelmi licenccel van beállítva.

## Státusz

Működő, egy konkrét feladatra készült eszköz (2024). Továbbfejlesztési irány:
konfigurálható oszlopkiosztás, és ugyanez a folyamat Pythonnal vagy n8n-nel,
telepítés nélkül.
