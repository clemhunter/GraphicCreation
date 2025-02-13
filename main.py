import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas.tseries.offsets import BDay
import io
import zipfile

st.title("Grafiken generieren")

# Drag & Drop File Uploader (nur .xlsx-Dateien)
uploaded_file = st.file_uploader("Ziehe deine Excel-Datei hierher", type=["xlsx"])

if uploaded_file:
    if st.button("Grafiken generieren"):
        # Excel-Datei einlesen
        df = pd.read_excel(
            uploaded_file,
            sheet_name='Sheet2',
            usecols=['Datum', 'Remaining Business Days', 'Event', 'Achieved', 'Completion', 'Accounts Invited']
        )

        # Einzigartige Events
        unique_events = df['Event'].unique()

        st.write("Grafiken werden generiert ...")

        # Dictionary zum Speichern der Bild-Puffer für jedes Event
        images = {}

        for event in unique_events:
            # Daten filtern, kopieren und nach Datum sortieren
            event_data = df[df['Event'] == event].copy()
            event_data['Datum'] = pd.to_datetime(event_data['Datum'], errors='coerce')
            event_data = event_data.sort_values('Datum')

            # Enddatum unter Berücksichtigung von Business Days berechnen
            event_data['Enddatum'] = event_data['Datum'] + event_data['Remaining Business Days'].apply(
                lambda x: BDay(x))

            # Berechnung des Ockerwerts (als Integer, z. B. 60)
            hline_value = int(np.ceil(event_data['Achieved'].iloc[0] / event_data['Completion'].iloc[0]))

            # Nur Zeilen mit gültigen Werten berücksichtigen
            event_data = event_data.dropna(subset=['Achieved', 'Accounts Invited'])

            # Letztes Datum (bei vorhandenen Daten) und Enddatum ermitteln
            last_date = event_data['Datum'].max()
            end_date = event_data['Enddatum'].max()

            # x-Achse: Gleichmäßige Abstände mit fortlaufendem Index
            x_data = np.arange(len(event_data))

            # Erstelle das Plot-Objekt
            fig, ax = plt.subplots(figsize=(12, 6))

            # Flächen füllen
            ax.fill_between(x_data, event_data['Accounts Invited'], color='grey', alpha=0.3)
            ax.fill_between(x_data, event_data['Achieved'], color='#DAA520', alpha=0.85)

            # Linienplots
            ax.plot(x_data, event_data['Achieved'], color='#DAA520', linestyle='-')
            ax.plot(x_data, event_data['Accounts Invited'], color='grey', linestyle='-')

            # Dynamische y-Skalierung
            accounts_invited_max = event_data['Accounts Invited'].max()
            newest_signups = event_data['Achieved'].iloc[-1]
            if accounts_invited_max < hline_value and newest_signups < hline_value:
                y_max = hline_value * 1.2  # 20% höher als hline_value
                y_ticks = [25, hline_value]
            else:
                y_val = max(accounts_invited_max, newest_signups)
                y_max = y_val * 1.2
                y_ticks = [25, hline_value, accounts_invited_max, newest_signups]
            y_ticks = sorted(set(y_ticks))
            ax.set_ylim(0, y_max)
            ax.set_yticks(y_ticks)

            # x-Achse
            x_tick_positions = list(x_data) + [len(event_data)]
            x_tick_labels = [d.strftime('%d.%m.') for d in event_data['Datum']]
            x_tick_labels.append(end_date.strftime('%d.%m.'))
            ax.set_xlim(0, len(event_data))
            ax.set_xticks(x_tick_positions)
            ax.set_xticklabels(x_tick_labels, rotation=45, color='white')
            ax.margins(x=0)

            # Anpassungen: Achsenrahmen, Farben, Schrift etc.
            for spine in ax.spines.values():
                spine.set_color('white')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.tick_params(axis='y', colors='white')
            plt.setp(ax.get_xticklabels(), fontsize=14, fontname='Arial', fontweight='bold')
            for label in ax.get_yticklabels():
                label.set_fontsize(20)
                label.set_fontname('Arial')
                label.set_fontweight('bold')
                try:
                    if int(label.get_text()) > 25:
                        label.set_color('#DAA520')
                except ValueError:
                    pass

            # Direkte Beschriftung: "Sign-Ups" und "Accounts-invited"
            x_signups = np.median(x_data)
            achieved_min = event_data['Achieved'].min()
            signups_y = max(achieved_min - 5, 0)
            ax.text(x_signups, signups_y, 'Sign-Ups', color='white', fontsize=13.5, fontname='Arial', ha='center',
                    va='top')

            x_accounts = np.median(x_data)
            accounts_y = min(accounts_invited_max + 3, y_max)
            ax.text(x_accounts, accounts_y, 'Accounts-invited', color='white', fontsize=13.5, fontname='Arial',
                    ha='center', va='bottom')

            plt.tight_layout()

            # Speichere den Plot in einen BytesIO-Puffer
            buf = io.BytesIO()
            fig.savefig(buf, format="png", transparent=True)
            buf.seek(0)
            plt.close(fig)

            # Speichere den Buffer im Dictionary
            images[event] = buf

            # Zusätzlich: Individueller Download-Button für jedes Event (mit unique key)
            st.download_button(
                label=f"Download Grafik für {event}",
                data=buf,
                file_name=f"{event}_plot.png",
                mime="image/png",
                key=f"{event}_download"
            )

        # Erstelle ein ZIP-Archiv, das alle Grafiken enthält
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for event, image_buf in images.items():
                zip_file.writestr(f"{event}_plot.png", image_buf.getvalue())
        zip_buffer.seek(0)

        # Biete einen Download-Button für das ZIP-Archiv an
        st.download_button(
            label="Download alle Grafiken als ZIP",
            data=zip_buffer,
            file_name="plots.zip",
            mime="application/zip",
            key="all_plots_zip"
        )
