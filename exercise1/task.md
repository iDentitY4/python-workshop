# Openpyxl

## Ausgangssituation

Sie erhalten eine Excel-Datei namens "verkaufsdaten.xlsx" mit folgender Struktur:

| Produkt  | Preis | Menge |
| -------- | ----- | ----- |
| Laptop   | 899   | 5     |
| Maus     | 29    | 12    |
| Monitor  | 299   | 8     |
| Tastatur | 49    | 15    |

## Aufgabe

Erstellen Sie ein Python-Skript, das:

1. Die Datei "verkaufsdaten.xlsx" öffnet
2. Eine neue Spalte "Gesamtwert" hinzufügt
3. Den Gesamtwert (Preis × Menge) für jedes Produkt berechnet
4. Die Ergebnisse in einer neuen Datei "verkaufsdaten_neu.xlsx" speichert

### Erwartetes Ergebnis:

Die neue Excel-Datei sollte eine zusätzliche Spalte "Gesamtwert" enthalten mit den berechneten Werten:

- Laptop: 4495
- Maus: 348
- Monitor: 2392
- Tastatur: 735

## Zusätzliche Herausforderung:

Formatieren Sie die Gesamtwert-Spalte als Währung mit dem Euro-Symbol.
