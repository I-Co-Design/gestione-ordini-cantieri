import os
import re
import sys
import json
import pandas as pd
from weasyprint import HTML

def estrai_dati_input(testo_input):
    # Riconosce le righe inviate dal Form e separa codice/descrizione dalla quantità finale
    righe = [r.strip() for r in testo_input.split('\n') if r.strip()]
    articoli_richiesti = []
    
    for riga in righe:
        parti = riga.split()
        if not parti:
            continue
        # Controlla se l'ultima parola è un numero (la quantità)
        if parti[-1].isdigit():
            quantita = int(parti[-1])
            ricerca = " ".join(parti[:-1])
        else:
            quantita = 1  # Default se dimentichi la quantità
            ricerca = " ".join(parti)
        articoli_richiesti.append({"cerca": ricerca.lower().strip(), "quantita": quantita})
    return articles_richiesti

def cerca_nel_catalogo(file_excel, nazione, articoli_richiesti):
    # Carica il foglio corretto saltando le prime 3 righe (parte da riga 4)
    df = pd.read_excel(file_excel, sheet_name=nazione, skiprows=3)
    df.columns = ['codice_cliente', 'codice_vero', 'descrizione', 'quantita_catalogo']
    
    # Pulisce i dati per la ricerca senza distinzione tra maiuscole e minuscole
    df['codice_vero_clean'] = df['codice_vero'].astype(str).str.lower().str.strip()
    df['descrizione_clean'] = df['descrizione'].astype(str).str.lower().str.strip()
    
    ordine_finale = []
    
    for art in articoli_richiesti:
        chiave = art["cerca"]
        # Cerca prima nel Codice Vero, poi nella Descrizione
        risultato = df[df['codice_vero_clean'] == chiave]
        if risultato.empty:
            risultato = df[df['descrizione_clean'].str.contains(chiave, na=False)]
            
        if not risultato.empty:
            riga = resultado.iloc[0]
            ordine_finale.append({
                "codice": riga['codice_vero'],
                "descrizione": riga['descrizione'],
                "quantita": art["quantita"]
            })
        else:
            # Se non trova la referenza nelle 1607 righe, la inserisce comunque con una nota
            ordine_finale.append({
                "codice": art["cerca"].upper(),
                "descrizione": "REFERENZA NON TROVATA NEL CATALOGO (Verificare)",
                "quantita": art["quantita"]
            })
    return ordine_finale

def genera_pdf(ordine, nazione, cantiere, montatore):
    # Crea un layout HTML pulito e moderno per il tuo PDF professionale
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: 'Helvetica', sans-serif; color: #333; margin: 30px; }}
            h1 {{ color: #0284c7; border-bottom: 2px solid #0284c7; padding-bottom: 10px; }}
            .info {{ margin-bottom: 20px; font-size: 14px; background: #f0f9ff; padding: 15px; border-left: 5px solid #0284c7; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background-color: #0284c7; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f8fafc; }}
        </style>
    </head>
    <body>
        <h1>Riepilogo Ordine Materiali</h1>
        <div class="info">
            <p><strong>Nazione:</strong> {nazione}</p>
            <p><strong>Punto Vendita / Cantiere:</strong> {cantiere}</p>
            <p><strong>Montatore:</strong> {montatore}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th>Codice Articolo</th>
                    <th>Descrizione Componente</th>
                    <th>Quantità</th>
                </tr>
            </thead>
            <tbody>
    """
    for item in ordine:
        html_content += f"""
                <tr>
                    <td><strong>{item['codice']}</strong></td>
                    <td>{item['descrizione']}</td>
                    <td>{item['quantita']} pezzi</td>
                </tr>
        """
    html_content += """
            </tbody>
        </table>
    </body>
    </html>
    """
    HTML(string=html_content).write_pdf("Ordine_Materiali.pdf")

if __name__ == "__main__":
    # Recupera i dati passati da GitHub Actions via variabili d'ambiente
    testo_form = os.getenv("DATI_ARTICOLI", "")
    nazione = os.getenv("NAZIONE", "ITALIA").upper()
    cantiere = os.getenv("CANTIERE", "Generico")
    montatore = os.getenv("MONTATORE", "Non specificato")
    
    articoli = estrai_dati_input(testo_form)
    risultato_ordine = cerca_nel_catalogo("ELABORATO COPIA INCOLLA.xlsx", nazione, articoli)
    genera_pdf(risultato_ordine, nazione, cantiere, montatore)
