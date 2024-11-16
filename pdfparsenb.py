# -*- coding: utf-8 -*-
"""
Created on Sat Nov 16 09:22:57 2024

@author: Nanni Bassetti - nannibassetti.com
"""

import fitz  # PyMuPDF
from datetime import datetime
import os
import pandas as pd
import re

def parse_pdf_date(date_str):
    """Parsa la data dal formato PDF in datetime"""
    try:
        match = re.match(r"D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})", date_str)
        if match:
            year, month, day, hour, minute, second = map(int, match.groups())
            return datetime(year, month, day, hour, minute, second)
    except Exception as e:
        print(f"Errore nel parsing della data PDF: {e}")
    return None

def extract_metadata_from_pdf(pdf_path):
    """Estrae i metadati di interesse da un PDF"""
    try:
        doc = fitz.open(pdf_path)
        metadata = doc.metadata
        doc.close()
        
        creation_date = parse_pdf_date(metadata.get("creationDate", ""))
        software = metadata.get("producer", "N/D")  # Software di editing (Producer)
        creator = metadata.get("creator", "N/D")   # Software di creazione (Creator)
        
        return creation_date, software, creator
    except Exception as e:
        print(f"Errore nell'elaborazione del file {pdf_path}: {e}")
        return None, "Errore", "Errore"

def process_pdf_directory(directory_path):
    """Legge tutti i PDF nella directory e ordina per data di creazione"""
    pdf_data = []
    for file_name in os.listdir(directory_path):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(directory_path, file_name)
            creation_date, software, creator = extract_metadata_from_pdf(pdf_path)
            pdf_data.append((file_name, creation_date, software, creator))
    
    # Ordina i dati per data di creazione (gestendo eventuali valori nulli)
    pdf_data = sorted(pdf_data, key=lambda x: x[1] or datetime.min)
    return pdf_data

def create_table(pdf_data, output_csv=None):
    """Crea e mostra una tabella con i risultati e salva in CSV se richiesto"""
    df = pd.DataFrame(pdf_data, columns=["File Name", "Creation Date", "Software", "Creator"])
    print(df.to_string(index=False))
    if output_csv:
        df.to_csv(output_csv, index=False)
        print(f"Tabella salvata in: {output_csv}")

# Main
directory_path = "percorso_alla_tua_directory"  # Modifica con il percorso della directory
output_csv = "tabella_metadati_pdf.csv"  # Modifica con il percorso di output desiderato

pdf_data = process_pdf_directory(directory_path)
create_table(pdf_data, output_csv)
