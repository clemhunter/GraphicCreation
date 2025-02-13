import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas.tseries.offsets import BDay
import io
import zipfile

st.title("Grafiken generieren")

# Drag & Drop File Uploader (only .xlsx files)
uploaded_file = st.file_uploader("Ziehe deine Excel-Datei hierher", type=["xlsx"])

if uploaded_file:
    if st.button("Grafiken generieren"):
        # Read the Excel file – ensure that the "Enddate" column is present
        df = pd.read_excel(
            uploaded_file,
            sheet_name='Sheet2',
            usecols=['Datum', 'Enddate', 'Remaining Business Days', 'Event', 'Achieved', 'Completion', 'Accounts Invited']
        )

        # Get unique events
        unique_events = df['Event'].unique()

        st.write("Grafiken werden generiert ...")

        # List of colors (five colors):
        # 1. Ocker (already used), 2. eb8c00, 3. db536a, 4. a chosen color, 5. another chosen color
        colors = ['#DAA520', '#eb8c00', '#db536a', '#009688', '#8a2be2']

        # Dictionary to store image buffers for each event
        images = {}

        for i, event in enumerate(unique_events):
            # Select a color for this event (cycling if more events than colors)
            color = colors[i % len(colors)]

            # Filter, copy and sort the data by date
            event_data = df[df['Event'] == event].copy()
            event_data['Datum'] = pd.to_datetime(event_data['Datum'], errors='coerce')
            event_data = event_data.sort_values('Datum')

            # Use the "Enddate" column from the Excel file (convert it to datetime)
            event_data['Enddatum'] = pd.to_datetime(event_data['Enddate'], errors='coerce')

            # Calculate the hline_value (e.g., 60) as an integer
            hline_value = int(np.ceil(event_data['Achieved'].iloc[0] / event_data['Completion'].iloc[0]))

            # Drop rows with missing values in 'Achieved' or 'Accounts Invited'
            event_data = event_data.dropna(subset=['Achieved', 'Accounts Invited'])

            # Determine the last date and the end date from the data
            last_date = event_data['Datum'].max()
            end_date = event_data['Enddatum'].max()

            # Create x-axis values with equal spacing
            x_data = np.arange(len(event_data))

            # Create the plot object
            fig, ax = plt.subplots(figsize=(12, 6))

            # Fill areas
            ax.fill_between(x_data, event_data['Accounts Invited'], color='grey', alpha=0.3)
            ax.fill_between(x_data, event_data['Achieved'], color=color, alpha=0.85)

            # Plot lines
            ax.plot(x_data, event_data['Achieved'], color=color, linestyle='-')
            ax.plot(x_data, event_data['Accounts Invited'], color='grey', linestyle='-')

            # Dynamic y-axis scaling
            accounts_invited_max = event_data['Accounts Invited'].max()
            newest_signups = event_data['Achieved'].iloc[-1]
            if accounts_invited_max < hline_value and newest_signups < hline_value:
                y_max = hline_value * 1.2  # 20% higher than hline_value
                y_ticks = [25, hline_value]
            else:
                y_val = max(accounts_invited_max, newest_signups)
                y_max = y_val * 1.2
                y_ticks = [25, hline_value, accounts_invited_max, newest_signups]
            y_ticks = sorted(set(y_ticks))
            ax.set_ylim(0, y_max)
            ax.set_yticks(y_ticks)

            # x-axis tick positions and labels
            x_tick_positions = list(x_data) + [len(event_data)]
            x_tick_labels = [d.strftime('%d.%m.') for d in event_data['Datum']]
            x_tick_labels.append(end_date.strftime('%d.%m.'))
            ax.set_xlim(0, len(event_data))
            ax.set_xticks(x_tick_positions)
            ax.set_xticklabels(x_tick_labels, rotation=45, color='white')
            ax.margins(x=0)

            # Adjust axes: spines, tick parameters, fonts, and colors
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
                        label.set_color(color)
                except ValueError:
                    pass

            # Direct annotations: "Sign-Ups" and "Accounts-invited"
            x_signups = np.median(x_data)
            achieved_min = event_data['Achieved'].min()
            signups_y = max(achieved_min - 5, 0)
            ax.text(x_signups, signups_y, 'Sign-Ups', color='white', fontsize=13.5, fontname='Arial',
                    ha='center', va='top')

            x_accounts = np.median(x_data)
            accounts_y = min(accounts_invited_max + 3, y_max)
            ax.text(x_accounts, accounts_y, 'Accounts-invited', color='white', fontsize=13.5, fontname='Arial',
                    ha='center', va='bottom')

            plt.tight_layout()

            # Save the plot to a BytesIO buffer
            buf = io.BytesIO()
            fig.savefig(buf, format="png", transparent=True)
            buf.seek(0)
            plt.close(fig)

            # Store the buffer in the dictionary
            images[event] = buf

            # Provide an individual download button for each event's plot
            st.download_button(
                label=f"Download Grafik für {event}",
                data=buf,
                file_name=f"{event}_plot.png",
                mime="image/png",
                key=f"{event}_download"
            )

        # Create a ZIP archive containing all the plots
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for event, image_buf in images.items():
                zip_file.writestr(f"{event}_plot.png", image_buf.getvalue())
        zip_buffer.seek(0)

        # Provide a download button for the ZIP archive
        st.download_button(
            label="Download alle Grafiken als ZIP",
            data=zip_buffer,
            file_name="plots.zip",
            mime="application/zip",
            key="all_plots_zip"
        )
