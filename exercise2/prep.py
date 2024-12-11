import numpy as np
import pandas as pd

# Beispieldaten erstellen
np.random.seed(42)
daten = {
    "Datum": pd.date_range("2024-01-01", "2024-03-31", freq="D"),
    "Produkt": np.random.choice(["Laptop", "Maus", "Monitor", "Tastatur"], size=91),
    "Preis": np.random.choice([899, 29, 299, 49], size=91),
    "Menge": np.random.randint(1, 10, size=91),
    "Filiale": np.random.choice(["Berlin", "Hamburg", "München"], size=91),
}

# DataFrame erstellen und speichern
df = pd.DataFrame(daten)
df.to_excel("e2_verkaufsdaten_2024.xlsx", index=False)
