import pandas as pd

# Excel-Datei einlesen
df = pd.read_excel("datei.xlsx")

# Spezifisches Arbeitsblatt lesen
df = pd.read_excel("datei.xlsx", sheet_name="Tabelle1")

# Mehrere Arbeitsblätter lesen
alle_tabellen = pd.read_excel(
    "datei.xlsx", sheet_name=None
)  # Dictionary mit allen Sheets
bestimmte_tabellen = pd.read_excel("datei.xlsx", sheet_name=["Tabelle1", "Tabelle2"])


# Bestimmte Spalten auswählen
df = pd.read_excel("datei.xlsx", usecols=["Name", "Alter"])

# Bestimmte Zeilen einlesen
df = pd.read_excel("datei.xlsx", nrows=10)  # Erste 10 Zeilen
df = pd.read_excel("datei.xlsx", skiprows=2)  # Erste 2 Zeilen überspringen

# Spaltentypen definieren
df = pd.read_excel("datei.xlsx", dtype={"Alter": int, "Name": str})

# NA-Werte behandeln
df = pd.read_excel("datei.xlsx", na_values=["N/A", "NA"])


# DataFrame in Excel speichern
df.to_excel("neue_datei.xlsx", index=False)

# Mit spezifischem Arbeitsblatt
df.to_excel("neue_datei.xlsx", sheet_name="Meine_Daten")

# Mehrere DataFrames in verschiedene Arbeitsblätter
with pd.ExcelWriter("neue_datei.xlsx") as writer:
    df1.to_excel(writer, sheet_name="Tabelle1")
    df2.to_excel(writer, sheet_name="Tabelle2")


# Mit Engine-Spezifikation
df.to_excel("neue_datei.xlsx", engine="openpyxl")

# Mit zusätzlichen Optionen
with pd.ExcelWriter("neue_datei.xlsx", engine="openpyxl") as writer:
    df.to_excel(
        writer,
        sheet_name="Tabelle1",
        index=False,
        float_format="%.2f",  # Dezimalstellen formatieren
        freeze_panes=(1, 0),
    )  # Erste Zeile fixieren


# Verfügbare Arbeitsblätter anzeigen
excel_datei = pd.ExcelFile("datei.xlsx")
sheets = excel_datei.sheet_names

# Vorschau der Daten
df = pd.read_excel("datei.xlsx")
print(df.head())  # Erste 5 Zeilen
print(df.info())  # Dateninformationen


# Fehlerbehanldung
try:
    df = pd.read_excel("datei.xlsx")
except FileNotFoundError:
    print("Datei nicht gefunden")
except Exception as e:
    print(f"Fehler beim Einlesen: {e}")


# Tipps

# Datum-Spalten korrekt einlesen
df = pd.read_excel("datei.xlsx", parse_dates=["Datum"])

# Spalten umbenennen beim Einlesen
df = pd.read_excel("datei.xlsx", names=["Spalte1", "Spalte2"])

# Datentypen optimieren
df = pd.read_excel("datei.xlsx").convert_dtypes()
