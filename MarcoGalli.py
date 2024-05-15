import pandas as pd  # Importa la libreria pandas con l'alias pd
import sqlite3  # Importa la libreria sqlite3 per lavorare con database SQLite
import re  # Importa il modulo re per gestire le espressioni regolari (regex)

# Carica i dati dall'Excel e li memorizza nel DataFrame df
df = pd.read_excel("1Sin.xls")

# Riempie i valori NaN (valori mancanti) con stringhe vuote nel DataFrame
df.fillna('', inplace=True)

# Fase 1: Converti il DataFrame in JSON e specifica l'orientamento come "records"
json_data = df.to_json(orient="records")

# Salva i dati JSON su un file chiamato "dati.json"
with open("dati.json", "w") as json_file:
    json_file.write(json_data)

# Stampa un messaggio indicando che la conversione da Excel a JSON è completata
print("Conversione da Excel a JSON completata.")

# Fase 2: Converti il JSON in un database SQLite e connettiti al database "dati.db"
conn = sqlite3.connect('dati.db')

# Itera su ogni colonna del DataFrame
for column in df.columns:
    # Rimuovi caratteri speciali dal nome della colonna e sostituisci con '_'
    table_name = re.sub(r'\W+', '_', column)

    # Crea una tabella nel database con il nome della colonna, con una colonna chiamata "Valore" di tipo TEXT
    cursor = conn.cursor()
    cursor.execute(f'''CREATE TABLE IF NOT EXISTS {table_name} (
                        Valore TEXT
                    )''')

    # Itera sui valori della colonna (escludendo la prima riga) e inserisci i valori nella tabella corrispondente
    for value in df[column][1:]:
        cursor.execute(f"INSERT INTO {table_name} (Valore) VALUES (?)", (str(value),))

# Salva i cambiamenti nel database
conn.commit()

# Chiudi la connessione al database
conn.close()

# Stampa un messaggio indicando che la conversione da JSON a SQLite è completata
print("Conversione da JSON a SQLite completata.")

# Stampa i dati del DataFrame
print("Dati del DataFrame:")
print(df)

# Stampa le query create per inserire i dati nel database SQLite
print("\nQuery per inserire i dati nel database SQLite:")
for row in df.itertuples(index=False):
    # Crea una stringa di valori separati da virgola, inclusi tra apici se sono stringhe
    values = ', '.join([f"'{str(val)}'" if isinstance(val, str) else str(val) for val in row])
    # Stampa la query SQL per inserire i valori nella tabella 'studenti'
    print(f"INSERT INTO studenti (Pr, Alunno, RELIGIONE, LINGUA_E_LETT_IT, LINGUAINGLESE, STORIA, EDUCAZIONE_CIVICA, MATEMATICA, DIRITTO_ED_ECONOMIA, FISICA, CHIMICA, Tecninformatiche, Tecne_Tecndirappr, SCDLLA_TERRA_GEO, SCIENZE_MOT_E_SPORT, COMPORTAMENTO, Media, Esito) VALUES ({values});")
