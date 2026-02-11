import pandas as pd
from openpyxl import load_workbook
import os
import warnings
import streamlit as st
import io
import traceback
import tempfile
from numbers_parser import Document

# Ignora avvisi non critici
warnings.filterwarnings("ignore")

# --- CONFIGURAZIONE FILE ---
# Questi rimangono come default o fallback
FILE_SORGENTE_DEFAULT = "Allenamento.xlsx"
FILE_MODELLO_DEFAULT = "excel.xlsx"

# --- CONFIGURAZIONE REGOLA SALTI ---
CONFIG_ANAGRAFICA = [
    # { "etichetta": "Data", "cella": "F2" }, # RIMOSSO: La data viene gestita esternamente
    { "etichetta": "Data di nascita", "cella": "B2" },
    { "etichetta": "Altezza", "cella": "C3" },
    { "etichetta": "Sesso", "cella": "G2" },
    { "etichetta": "Peso", "cella": "C4" },
    { "etichetta": "lunghezza gamba", "cella": "E5" },
    { "etichetta": "altezza dei fianchi durante flessione SJ", "cella": "E6" }
]

REGISTRO_SALTI = [
    # 1. PRIMA SERIE ABK
    {
        "tipo": "ABK", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F9", "G9", "H9"]}]
    },
    # 2. PRIMA SERIE CMJ
    {
        "tipo": "CMJ", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F10", "G10", "H10"]}]
    },
    # 3. SECONDA SERIE ABK (Se esiste)
    {
        "tipo": "ABK", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F12", "G12", "H12"]}]
    },
    # 4. SECONDA SERIE CMJ (Se esiste)
    {
        "tipo": "CMJ", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F13", "G13", "H13"]}]
    },
    # 5. SJ
    {
        "tipo": "SJ", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["S15", "T15", "U15"]}]
    },
    # 6. slCMJ Left
    {
        "tipo": "slCMJleft", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F16", "G16", "H16"]}]
    },
    # 7. slCMJ Right
    {
        "tipo": "slCMJright", "discriminante": None,
        "outputs": [{"dato": "Altezza", "celle": ["F15", "G15", "H15"]}]
    },

    # --- DJa (DROP JUMPS) ---
    {
        "tipo": "DJa", "discriminante": ("Caduta", 30),
        "outputs": [{"dato": "Altezza", "celle": ["P4", "Q4", "R4"]}, {"dato": "TC", "celle": ["T4", "U4", "V4"]}]
    },
    {
        "tipo": "DJa", "discriminante": ("Caduta", 45),
        "outputs": [{"dato": "Altezza", "celle": ["P5", "Q5", "R5"]}, {"dato": "TC", "celle": ["T5", "U5", "V5"]}]
    },
    {
        "tipo": "DJa", "discriminante": ("Caduta", 60),
        "outputs": [{"dato": "Altezza", "celle": ["P6", "Q6", "R6"]}, {"dato": "TC", "celle": ["T6", "U6", "V6"]}]
    },
    {
        "tipo": "DJa", "discriminante": ("Caduta", 75),
        "outputs": [{"dato": "Altezza", "celle": ["P7", "Q7", "R7"]}, {"dato": "TC", "celle": ["T7", "U7", "V7"]}]
    },
    {
        "tipo": "DJa", "discriminante": ("Caduta", 90),
        "outputs": [{"dato": "Altezza", "celle": ["P8", "Q8", "R8"]}, {"dato": "TC", "celle": ["T8", "U8", "V8"]}]
    },
    {
        "tipo": "DJa", "discriminante": ("Caduta", 105),
        "outputs": [{"dato": "Altezza", "celle": ["P9", "Q9", "R9"]}, {"dato": "TC", "celle": ["T9", "U9", "V9"]}]
    },

    # --- SJl (SJI - SALTI CON CARICO) ---
    # Logica dinamica: prende le serie SJl in ordine di comparsa nel file
    {
        "tipo": "SJl", "discriminante": None,
        "weight_output": "R16",
        "outputs": [{"dato": "Altezza", "celle": ["S16", "T16", "U16"]}]
    },
    {
        "tipo": "SJl", "discriminante": None,
        "weight_output": "R17",
        "outputs": [{"dato": "Altezza", "celle": ["S17", "T17", "U17"]}]
    },
    {
        "tipo": "SJl", "discriminante": None,
        "weight_output": "R18",
        "outputs": [{"dato": "Altezza", "celle": ["S18", "T18", "U18"]}]
    },
    {
        "tipo": "SJl", "discriminante": None,
        "weight_output": "R19",
        "outputs": [{"dato": "Altezza", "celle": ["S19", "T19", "U19"]}]
    }
]


def custom_round(val, decimals=0):
    """
    Arrotondamento aritmetico:
    - .5 arrotonda sempre per eccesso (es. 2.5 -> 3, 3.5 -> 4)
    - Gestisce sia numeri che stringhe convertibili
    """
    try:
        if val is None or str(val).strip() == "":
            return val
            
        n = float(str(val).replace(',', '.'))
        multiplier = 10 ** decimals
        # Aggiungiamo un epsilon piccolissimo per gestire errori di floating point
        res = int(n * multiplier + 0.5) / multiplier
        
        if decimals == 0:
            return int(res)
        return res
    except:
        return val


def carica_file_universale(uploaded_file):
    """Carica file Excel o CSV da un oggetto file-like di Streamlit o path"""
    if uploaded_file is None:
        return None

    # Se è una stringa (per retrocompatibilità o test locale), lo trattiamo come path
    if isinstance(uploaded_file, str):
        filepath = uploaded_file
        print(f"Lettura file path: {filepath}")
        try:
            return pd.read_excel(filepath, header=None)
        except:
            pass
        for sep in [',', ';', '\t']:
            try:
                df = pd.read_csv(filepath, header=None, sep=sep, encoding='latin1', on_bad_lines='skip', engine='python')
                if df.shape[1] > 1: return df
            except:
                continue
        return None

    # Altrimenti è un buffer di Streamlit
    filename = uploaded_file.name
    print(f"Lettura buffer: {filename}")

    try:
        return pd.read_excel(uploaded_file, header=None)
    except:
        pass

    # Reset del puntatore per tentativi CSV
    uploaded_file.seek(0)

    # Supporto per file Apple .numbers
    if filename.endswith('.numbers'):
        try:
            # numbers-parser richiede un file fisico su disco
            with tempfile.NamedTemporaryFile(delete=False, suffix=".numbers") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name
            
            try:
                doc = Document(tmp_path)
                sheets = doc.sheets
                if sheets:
                    tables = sheets[0].tables
                    if tables:
                        table = tables[0]
                        data = table.rows(values_only=True)
                        df = pd.DataFrame(data)
                        return df
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                    
        except Exception as e:
            print(f"Errore caricamento .numbers: {e}")

    for sep in [',', ';', '\t']:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, header=None, sep=sep, encoding='latin1', on_bad_lines='skip', engine='python')
            if df.shape[1] > 1: return df
        except:
            continue
    return None


def trova_valore_cella(df, keywords):
    """(Step 1) Cerca un valore nella griglia usando una o più parole chiave"""
    if isinstance(keywords, str): keywords = [keywords]

    df_str = df.astype(str).apply(lambda x: x.str.lower().str.strip())

    for key in keywords:
        matches = df_str.stack()
        found = matches[matches == key.lower()].index.tolist()

        if not found:
            found = matches[matches.str.contains(key.lower(), regex=False)].index.tolist()

        for (r, c) in found:
            try:
                val = str(df.iloc[r + 1, c]).strip()
                if val not in ["nan", "0", "0.0", "", "None"]:
                    return val
            except:
                continue
    return ""


def raggruppa_salti_per_serie(df_salti):
    """Divide i salti in gruppi contigui."""
    gruppi = []
    if df_salti.empty: return gruppi

    current_chunk = []
    first_row = df_salti.iloc[0]
    last_tipo = first_row['Tipo']
    last_caduta = first_row['Caduta'] if 'Caduta' in df_salti.columns else -1
    last_peso = first_row['Peso Kg'] if 'Peso Kg' in df_salti.columns else -1

    for index, row in df_salti.iterrows():
        curr_tipo = row['Tipo']
        curr_caduta = row['Caduta'] if 'Caduta' in df_salti.columns else -1
        curr_peso = row['Peso Kg'] if 'Peso Kg' in df_salti.columns else -1

        cambio_serie = (curr_tipo != last_tipo) or \
                       (abs(float(curr_caduta) - float(last_caduta)) > 0.1) or \
                       (abs(float(curr_peso) - float(last_peso)) > 0.1)

        if cambio_serie:
            gruppi.append({
                'tipo': last_tipo,
                'caduta': last_caduta,
                'peso': last_peso,
                'data': pd.DataFrame(current_chunk)
            })
            current_chunk = []
            last_tipo = curr_tipo
            last_caduta = curr_caduta
            last_peso = curr_peso

        current_chunk.append(row)

    if current_chunk:
        gruppi.append({
            'tipo': last_tipo,
            'caduta': last_caduta,
            'peso': last_peso,
            'data': pd.DataFrame(current_chunk)
        })

    return gruppi


def elabora_salti_cronologici(df, ws, data_selezionata):
    """Elabora i salti e scrive nel worksheet."""
    st.write("--- ESECUZIONE STEP 2 (ORDINE CRONOLOGICO) ---")

    riga_header = -1
    col_map = {}

    for i, riga in df.iterrows():
        riga_lista = [str(x).strip().lower() for x in riga.tolist()]
        if "tipo" in riga_lista and "altezza" in riga_lista:
            riga_header = i
            for idx, val in enumerate(riga_lista): col_map[val] = idx
            break

    if riga_header == -1:
        st.error("ERRORE: Tabella salti non trovata.")
        return

    df_data = df.iloc[riga_header + 1:].copy()

    def get_col_values(nome_col):
        if nome_col.lower() in col_map: return df_data.iloc[:, col_map[nome_col.lower()]]
        return None

    clean_df = pd.DataFrame()
    clean_df['Tipo'] = get_col_values('Tipo').astype(str).str.strip()

    raw_alt = get_col_values('Altezza')
    clean_df['Altezza'] = pd.to_numeric(raw_alt.astype(str).str.replace(',', '.'), errors='coerce') if raw_alt is not None else 0.0

    raw_tc = get_col_values('TC')
    clean_df['TC'] = pd.to_numeric(raw_tc.astype(str).str.replace(',', '.'), errors='coerce') if raw_tc is not None else 0.0

    raw_caduta = get_col_values('Caduta')
    clean_df['Caduta'] = pd.to_numeric(raw_caduta.astype(str).str.replace(',', '.'), errors='coerce').fillna(-1) if raw_caduta is not None else -1

    raw_peso = get_col_values('Peso Kg')
    if raw_peso is None: raw_peso = get_col_values('Peso')
    clean_df['Peso Kg'] = pd.to_numeric(raw_peso.astype(str).str.replace(',', '.'), errors='coerce').fillna(-1) if raw_peso is not None else -1

    # Filtraggio per data
    raw_data = get_col_values('Data')
    if raw_data is not None:
        clean_df['Data_Originale'] = raw_data.astype(str).str.strip()
        # Convertiamo la data selezionata in stringa per il confronto (formato YYYY-MM-DD)
        # Il formato nel file potrebbe variare, proviamo a normalizzare
        def confronta_date(data_file_str, data_selected):
            try:
                # Prova parsing standard
                d_file = pd.to_datetime(data_file_str).date()
                return d_file == data_selected
            except:
                # Fallback: confronto stringa parziale se fallisce il parsing
                return str(data_selected) in data_file_str

        maschera_data = clean_df['Data_Originale'].apply(lambda x: confronta_date(x, data_selezionata))
        clean_df = clean_df[maschera_data]

    clean_df = clean_df.dropna(subset=['Altezza'])

    gruppi_disponibili = raggruppa_salti_per_serie(clean_df)
    st.info(f"Trovati {len(gruppi_disponibili)} gruppi di salti per la data {data_selezionata}.")

    gruppi_usati = [False] * len(gruppi_disponibili)

    for regola in REGISTRO_SALTI:
        tipo_req = regola['tipo']
        discrim = regola['discriminante']

        gruppo_trovato = None
        idx_trovato = -1

        for i, gruppo in enumerate(gruppi_disponibili):
            if gruppi_usati[i]: continue

            tipo_ok = (gruppo['tipo'].lower() == tipo_req.lower())
            if tipo_req.lower() == "sji":
                tipo_ok = tipo_ok or (gruppo['tipo'].lower() == "sjl")

            if not tipo_ok: continue

            discrim_ok = True
            if discrim:
                nome_d, val_d = discrim
                val_gruppo = gruppo['caduta'] if nome_d == "Caduta" else gruppo['peso']
                if abs(float(val_gruppo) - float(val_d)) > 0.1:
                    discrim_ok = False

            if discrim_ok:
                gruppo_trovato = gruppo['data']
                idx_trovato = i
                break

        if gruppo_trovato is not None:
            gruppi_usati[idx_trovato] = True
            # st.write(f" -> Regola {tipo_req} (Disc: {discrim}): USATO Gruppo {idx_trovato} ({len(gruppo_trovato)} salti)")
            # --- SEZIONE AGGIORNATA ---
            if "weight_output" in regola:
                # Prende il peso dal gruppo corrente (Serie 1, Serie 2, ecc.)
                peso_effettivo = gruppi_disponibili[idx_trovato]['peso']
                # Scrive il peso nella cella R configurata arrotondato a 1 cifra
                try: 
                    peso_effettivo = custom_round(float(str(peso_effettivo).replace(',', '.')), 1)
                except: 
                    pass
                ws[regola["weight_output"]] = peso_effettivo
                if isinstance(peso_effettivo, (int, float)):
                    ws[regola["weight_output"]].number_format = '0.00'
                print(f" -> Scritto peso {peso_effettivo} in {regola['weight_output']}")
            # --------------------------
            # Prende i top 3 valori di Altezza
            df_sorted_top = gruppo_trovato.sort_values(by='Altezza', ascending=False).head(3)

            for out_conf in regola['outputs']:
                col_dato = out_conf['dato']
                celle = out_conf['celle']
                
                # Determina applicazione arrotondamento
                # Richiesta: NO arrotondamento per DJa (solo dato TC), SI arrotondamento per altri
                no_round_dja = (tipo_req == "DJa" and col_dato == "TC")

                vals_raw = df_sorted_top[col_dato].tolist()
                vals_processed = []
                for v in vals_raw:
                    try:
                        val_float = float(str(v).replace(',', '.'))
                        if no_round_dja:
                            vals_processed.append(val_float)
                        else:
                            # Usa arrotondamento custom
                            vals_processed.append(custom_round(val_float, 1))
                    except:
                        if v is not None and str(v).lower() != "nan":
                            vals_processed.append(v)
                
                # NUOVA LOGICA: Gestione 1 o 2 dati per riempire 3 caselle
                if len(vals_processed) == 1:
                    # Caso 1 dato -> Lo ripetiamo per 3 volte
                    val = vals_processed[0]
                    vals_processed = [val, val, val]
                elif len(vals_processed) == 2:
                    # Caso 2 dati -> Facciamo la media tra i due per il terzo dato
                    try:
                        v1 = float(vals_processed[0])
                        v2 = float(vals_processed[1])
                        media = (v1 + v2) / 2
                        if not no_round_dja:
                            media = custom_round(media, 1)
                        vals_processed.append(media)
                    except:
                        pass
                
                # Ordinamento CRESCENTE richiesto dall'utente
                vals_processed.sort()

                for k, cella in enumerate(celle):
                    val = vals_processed[k] if k < len(vals_processed) else ""
                    ws[cella] = val
                    if isinstance(val, (int, float)):
                        if no_round_dja:
                            ws[cella].number_format = '0.000'
                        else:
                            ws[cella].number_format = '0.00'
        else:
            # st.warning(f" -> Regola {tipo_req} (Disc: {discrim}): NESSUN GRUPPO TROVATO (Lascio bianco)")
            for out_conf in regola['outputs']:
                for cella in out_conf['celle']:
                    ws[cella] = ""


def elabora_salti_rj(df, ws, data_selezionata):
    """Elabora i salti reattivi (RJ) e scrive nel worksheet."""
    st.write("--- ESECUZIONE STEP 3 (SALTI REATTIVI RJ) ---")

    riga_header = -1
    col_map = {}

    # 1. Trova l'header della sezione RJ
    # Cerchiamo l'header che contiene 'tipo di salto' e 'tc avg'
    for i, riga in df.iterrows():
        riga_lista = [str(x).strip().lower() for x in riga.tolist()]
        if "tipo di salto" in riga_lista and "tc avg" in riga_lista:
            riga_header = i
            for idx, val in enumerate(riga_lista): col_map[val] = idx
            break

    if riga_header == -1:
        st.warning("Sezione RJ (Salti Reattivi) non trovata nel file.")
        return

    # 2. Estrai tutte le righe di riepilogo RJ
    df_data = df.iloc[riga_header + 1:].copy()
    
    # Identifichiamo le righe di riepilogo: hanno un valore in 'tipo di salto' (RJ) e una 'Data'
    def is_summary_row(row):
        tipo = str(row.iloc[col_map.get('tipo di salto', 0)]).strip().lower()
        data_val = str(row.iloc[col_map.get('data', 0)]).strip().lower()
        return "rj" in tipo and data_val not in ["nan", "none", ""]

    summary_mask = df_data.apply(is_summary_row, axis=1)
    df_summaries = df_data[summary_mask].copy()

    if df_summaries.empty:
        st.warning("Nessuna serie RJ trovata.")
        return

    # 3. Filtra per data selezionata
    def confronta_date(row, data_selected):
        data_file_str = str(row.iloc[col_map.get('data', 0)]).strip()
        try:
            d_file = pd.to_datetime(data_file_str).date()
            return d_file == data_selected
        except:
            return str(data_selected) in data_file_str

    df_summaries['date_match'] = df_summaries.apply(lambda r: confronta_date(r, data_selezionata), axis=1)
    df_filtered = df_summaries[df_summaries['date_match']]

    if df_filtered.empty:
        st.warning(f"Nessuna serie RJ trovata per la data {data_selezionata}.")
        return

    # 4. Scegli quella con TC AVG minore
    col_tc_avg = col_map.get('tc avg')
    df_filtered['TC_AVG_NUM'] = pd.to_numeric(df_filtered.iloc[:, col_tc_avg].astype(str).str.replace(',', '.'), errors='coerce')
    best_series_row = df_filtered.sort_values(by='TC_AVG_NUM', ascending=True).iloc[0]
    
    # 5. Estrai e scrivi i parametri
    # AVG altezza: F19, TC AVG: H19, AVG RSI: I19
    val_avg_altezza = best_series_row.iloc[col_map.get('avg altezza')]
    val_tc_avg = best_series_row.iloc[col_map.get('tc avg')]
    val_avg_rsi = best_series_row.iloc[col_map.get('avg rsi')]

    def secure_num(val, decimals=None):
        # Richiesta: NO arrotondamento per RJ tranne H AVG, ma usiamo custom_round se serve
        try: 
            v = float(str(val).replace(',', '.'))
            if decimals is not None:
                return custom_round(v, decimals)
            return v
        except: return val

    # MODIFICA: F19 (AVG Altezza) deve avere 2 decimali e arrotondamento custom (ma l'ultimo a 0 -> round a 1)
    # H19 (TC AVG) torna a 3 decimali default
    ws["F19"] = secure_num(val_avg_altezza, 1)
    ws["F19"].number_format = '0.00'

    ws["H19"] = secure_num(val_tc_avg) 
    ws["H19"].number_format = '0.000'
    
    ws["I19"] = secure_num(val_avg_rsi)
    ws["I19"].number_format = '0.000'
    
    st.info(f"Serie RJ selezionata (TC AVG: {val_tc_avg}). Scritto in F19 (Altezza), H19 (TC), I19 (RSI).")

    # 6. Trova TC Minore tra i singoli salti positivi
    # Le righe dei singoli salti seguono la riga di riepilogo
    idx_summary = best_series_row.name
    # Vediamo le righe successive finché non troviamo un'altra serie o la fine
    jumps_data = df.iloc[idx_summary + 1 : idx_summary + 15] # Tipicamente 10-15 salti max
    
    min_tc_pos = float('inf')
    col_tc_individual = 1 # Dai dati sembra che i singoli salti abbiano TC in colonna 1 (B)
    
    # Cerchiamo di identificare la colonna TC per i singoli salti. 
    # Spesso è la stessa colonna o quella definita nell'header come 'TC' (non AVG)
    # Dalle verifiche precedenti, colonna 1 (indice 1) conteneva 0.235, 0.226 etc.
    
    for i, row in jumps_data.iterrows():
        try:
            val_str = str(row.iloc[1]).strip().replace(',', '.')
            val_num = float(val_str)
            if val_num > 0 and val_num < min_tc_pos:
                min_tc_pos = val_num
        except:
            continue
            
    if min_tc_pos != float('inf'):
        ws["G19"] = min_tc_pos
        ws["G19"].number_format = '0.000'
        st.info(f"TC minore positivo trovato: {ws['G19'].value} (scritto in G19).")
    else:
        ws["G19"] = ""


def elabora_step1_anagrafica(df, ws, data_selezionata):
    """Elabora l'anagrafica leggendo ESCLUSIVAMENTE la riga sotto 'ID'"""
    st.write("--- ESECUZIONE STEP 1 (ANAGRAFICA RIGIDA) ---")

    # 1. Data Test (F2)
    ws["F2"] = data_selezionata.strftime("%d/%m/%Y")

    # 2. Trova la riga dell'intestazione (dove c'è scritto ID, Nome, Altezza...)
    riga_header = -1
    col_map = {}
    for i, riga in df.iterrows():
        riga_check = [str(x).strip().lower() for x in riga.tolist()]
        # Cerchiamo la riga che contiene i metadati dell'atleta
        if "id" in riga_check and "nome" in riga_check:
            riga_header = i
            for idx, col_name in enumerate(riga_check):
                col_map[col_name] = idx
            break

    if riga_header == -1:
        st.error("ERRORE: Riga 'ID' non trovata. Impossibile leggere l'altezza corretta.")
        return

    # La riga dell'atleta è quella immediatamente sotto l'header ID
    riga_atleta = df.iloc[riga_header + 1]

    # FUNZIONE LOCALE: Prende il dato SOLO dalla riga dell'atleta
    def prendi_solo_da_riga_id(nome_colonna):
        n = nome_colonna.lower().strip()
        if n in col_map:
            idx = col_map[n]
            val = str(riga_atleta.iloc[idx]).strip()
            # Il valore se 0, nan, none o vuoto deve essere "" (BIANCO)
            if val.lower() in ["nan", "none", "", "0", "0.0"]:
                return ""
            return val
        return ""

    # 3. Nome e Cognome (C1, E1)
    full_name = prendi_solo_da_riga_id("nome")
    if not full_name: full_name = prendi_solo_da_riga_id("nome persona")
    
    if full_name:
        parti = full_name.split(" ", 1)
        ws["C1"] = parti[0].strip().upper() if len(parti) > 0 else ""
        ws["E1"] = parti[1].strip().upper() if len(parti) > 1 else ""
        st.info(f" -> Splittato nome: {ws['C1'].value} (C1), {ws['E1'].value} (E1)")

    # 4. Ciclo sui campi (Altezza, Peso, ecc.) usando SOLO la riga ID
    for item in CONFIG_ANAGRAFICA:
        etichetta = item['etichetta']
        cella = item['cella']
        
        valore = prendi_solo_da_riga_id(etichetta)
        
        # Fallback per il Peso
        if not valore and etichetta.lower() == "peso":
            valore = prendi_solo_da_riga_id("peso kg")

        # Conversione Sesso
        if etichetta.lower() == "sesso":
            if valore.upper() == "M": valore = "UOMO"
            elif valore.upper() == "F": valore = "DONNA"
            
        
        # Arrotondamento anagrafica: RIMUOVERE LA VIRGOLA (Interi)
        # Richiesta: "levare i numeri con la virgola dai dati anagrafici" -> custom_round(val, 0)
        is_number = False
        try:
            val_num = float(str(valore).replace(',', '.'))
            valore = custom_round(val_num, 0) # Arrotondamento all'intero
            is_number = True
        except:
            pass

        ws[cella] = valore
        if is_number:
            ws[cella].number_format = '0' # Formato intero senza decimali
        
        st.info(f" -> {etichetta}: {valore} (scritto in {cella})")


def main():
    st.set_page_config(page_title="Elaborazione Dati Atletici", page_icon="🏃‍♂️")
    st.title("🚀 Athletic Data Excel Sync 📈")
    st.markdown("Carica il file dati e il modello Excel per generare il report finale.")

    # 1. Widget Caricamento File Sorgente
    uploaded_file_sorgente = st.file_uploader("Carica file 'Allenamento' (Excel, CSV o Numbers)", type=['xlsx', 'csv', 'numbers'])

    # 2. Widget Caricamento File Modello
    uploaded_file_modello = st.file_uploader("Carica file 'Modello' (Excel)", type=['xlsx'])

    # Opzione per usare modello locale se non caricato
    path_modello_locale = FILE_MODELLO_DEFAULT
    uso_modello_locale = False

    if not uploaded_file_modello:
        if os.path.exists(path_modello_locale):
            st.info(f"Nessun modello caricato. Verrà usato '{path_modello_locale}' se presente nella cartella.")
            uso_modello_locale = True
        else:
            st.warning(f"Attenzione: Modello '{path_modello_locale}' non trovato in locale. Caricane uno.")

    # 3. Data dei Dati
    data_test = st.date_input("📅 Test Date Selector 📅", value=pd.Timestamp.now().date(), format="DD/MM/YYYY")

    # 4. Nome Output
    output_name = st.text_input("Nome file di output", value="Risultati_Atleta.xlsx")
    if not output_name.endswith(".xlsx"):
        output_name += ".xlsx"

    # 4. Pulsante elaborazione
    if st.button("Avvia Elaborazione"):
        if not uploaded_file_sorgente:
            st.error("Per favore carica il file sorgente 'Allenamento'.")
            return

        # Determinazione del modello da usare
        modello_da_usare = None
        if uploaded_file_modello:
            modello_da_usare = uploaded_file_modello
        elif uso_modello_locale:
            modello_da_usare = path_modello_locale
        else:
            st.error("Manca il file Modello! Caricalo.")
            return

        with st.spinner("Elaborazione in corso..."):
            try:
                # A. Caricamento Dati
                df = carica_file_universale(uploaded_file_sorgente)
                if df is None:
                    st.error("Errore lettura file sorgente. Verifica il formato.")
                    return

                # B. Caricamento Modello e Processamento
                wb = load_workbook(modello_da_usare)
                if 'ATLETA' not in wb.sheetnames:
                    st.error("Il file modello non contiene il foglio 'ATLETA'.")
                    return

                ws = wb['ATLETA']

                # C. Esecuzione Step
                elabora_step1_anagrafica(df, ws, data_test)
                elabora_salti_cronologici(df, ws, data_test)
                elabora_salti_rj(df, ws, data_test)

                # D. Salvataggio in memoria (BytesIO)
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)

                st.success("Elaborazione Completata!")

                # E. Bottone Download
                st.download_button(
                    label="📥 Scarica File Elaborato",
                    data=buffer,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Errore critico: {e}")
                st.text(traceback.format_exc())

if __name__ == "__main__":
    main()