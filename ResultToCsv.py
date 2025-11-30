import tkinter as tk
from tkinter import filedialog
import csv

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename(
    title="Wybierz plik .Result z EMC32",
    filetypes=[("Pliki Result", "*.Result"), ("Wszystkie", "*.*")]
)
if not file_path:
    exit()

output_path = filedialog.asksaveasfilename(
    title="Zapisz jako CSV",
    defaultextension=".csv",
    filetypes=[("CSV", "*.csv")]
)
if not output_path:
    exit()

rows = []

# Próba otwarcia w różnych kodowaniach – zawsze się uda
text = None
for enc in ('utf-16', 'cp1252', 'utf-8-sig', 'utf-8'):
    try:
        with open(file_path, "r", encoding=enc) as f:
            text = f.read()
        print(f"Odczytano jako: {enc}")
        break
    except:
        continue

if text is None:
    print("Błąd odczytu pliku!")
    exit()

lines = text.splitlines()
data_section = False

for line in lines:
    line = line.rstrip()

    # Start sekcji danych
    if "[TableValues]" in line:
        data_section = True
        continue                     # ←←←← TU BYŁO skip = 2 → teraz NIE POMIJAMY ŻADNEJ LINII NAGŁÓWKOWEJ

    if not data_section:
        continue

    # Pomijamy tylko prawdziwe linie nagłówkowe (te z literami, nie z liczbami)
    if any(x in line for x in ("Frequency", "AVG", "PK+", "QPK", "Corr", "MHz", "dBµV", "dB")):
        continue

    if not line.strip():
        continue

    # Dzielimy po tabulatorach i bierzemy tylko niepuste pola
    fields = [f.strip() for f in line.split("\t") if f.strip()]

    # Teraz bierzemy dokładnie 5 pól (nawet jeśli jest ich więcej)
    if len(fields) >= 5:
        try:
            row = [float(f.replace(",", ".")) for f in fields[:5]]
            rows.append(row)
        except ValueError:
            # Czasem w pierwszej linii jest coś w stylu "30.000000\t\t\t\t" – to pomijamy
            pass

# Zapis
with open(output_path, "w", encoding="utf-8", newline="") as f:
    writer = csv.writer(f, delimiter=";")
    writer.writerow(["Frequency_MHz", "AVG_CLRWR_dBuV", "PK+_CLRWR_dBuV", "QPK_CLRWR_dBuV", "Corr_dB"])
    for r in rows:
        writer.writerow([f"{x:.10f}".replace(".", ",") for x in r])   # więcej cyfr = pewność

print(f"\n100% GOTOWE! Zapisano {len(rows)} wierszy → powinno być dokładnie 5668")
print(f"Plik: {output_path}")
input("\nNaciśnij Enter...")