import streamlit as st
from datetime import datetime, timedelta
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# --- Style CSS ---
st.markdown("""
<style>
/* Style du tableau affiché */
.stDataFrame {
    width: 100%;
    margin: 20px 0;
    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
}

.stDataFrame th {
    background-color: #1eaabd !important;
    color: white !important;
    font-weight: bold;
    text-align: center !important;
}

.stDataFrame td {
    padding: 10px !important;
    border: 1px solid #e0e0e0 !important;
}

.stDataFrame tr:nth-child(even) {
    background-color: #f5f5f5;
}

/* Boutons */
.stButton button {
    background-color: #1eaabd !important;
    color: white !important;
    border: none !important;
    margin: 5px 0;
}
</style>
""", unsafe_allow_html=True)

# --- Fonctions ---
def convert_to_minutes(time_str):
    try:
        h, m = map(int, time_str.split(':'))
        return h * 60 + m
    except:
        return 0

def format_time(minutes):
    return f"{minutes//60}h{minutes%60:02d}"

def generate_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Planning')
        workbook = writer.book
        worksheet = writer.sheets['Planning']
        
        # Style pour les en-têtes
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1eaabd", end_color="1eaabd", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), 
                           right=Side(style='thin'), 
                           top=Side(style='thin'), 
                           bottom=Side(style='thin'))
        
        # Appliquer le style aux en-têtes
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center')
        
        # Ajuster la largeur des colonnes
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2) * 1.2
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()

def generate_html(df):
    html = f"""
    <html>
    <head>
    <title>Planning Hebdomadaire</title>
    <style>
        body {{ font-family: Arial; margin: 20px; }}
        table {{ border-collapse: collapse; width: 100%; }}
        th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
        th {{ background-color: #1eaabd; color: white; }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
    </style>
    </head>
    <body>
    <h2 style="color: #1eaabd;">Planning Hebdomadaire</h2>
    {df.to_html(index=False)}
    </body>
    </html>
    """
    return html

# --- Initialisation ---
jours = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

if 'data' not in st.session_state:
    default_date = datetime.now() - timedelta(days=datetime.now().weekday())
    st.session_state.data = {
        'start_date': default_date,
        'horaires': {j: {'arrivee': '', 'depart': ''} for j in jours},
        'repos': {j: False for j in jours}
    }

# --- Interface ---
st.title("📅 Planning Hebdomadaire")

# Sélection de la semaine
new_start_date = st.date_input(
    "Date de début (lundi de la semaine)",
    value=st.session_state.data['start_date']
)

if new_start_date != st.session_state.data['start_date']:
    st.session_state.data['start_date'] = new_start_date
    st.rerun()

week_dates = [st.session_state.data['start_date'] + timedelta(days=i) for i in range(7)]
st.subheader(f"Semaine du {week_dates[0].strftime('%d/%m')} au {week_dates[-1].strftime('%d/%m')}")

# Saisie des horaires
cols = st.columns(7)
for i, (jour, date) in enumerate(zip(jours, week_dates)):
    with cols[i]:
        st.write(f"**{jour}**")
        st.write(date.strftime('%d/%m'))
        
        st.session_state.data['repos'][jour] = st.checkbox(
            "Repos",
            value=st.session_state.data['repos'][jour],
            key=f"repos_{jour}"
        )
        
        if not st.session_state.data['repos'][jour]:
            st.session_state.data['horaires'][jour]['arrivee'] = st.text_input(
                "Arrivée",
                value=st.session_state.data['horaires'][jour]['arrivee'],
                key=f"arr_{jour}",
                placeholder="09:00"
            )
            st.session_state.data['horaires'][jour]['depart'] = st.text_input(
                "Départ",
                value=st.session_state.data['horaires'][jour]['depart'],
                key=f"dep_{jour}",
                placeholder="17:00"
            )

# Calcul et affichage
if st.button("Générer le planning"):
    data = []
    total_minutes = 0
    
    for jour, date in zip(jours, week_dates):
        date_str = date.strftime('%d/%m')
        if st.session_state.data['repos'][jour]:
            data.append([f"{jour} {date_str}", "Repos", "", "0h00"])
        else:
            arrivee = st.session_state.data['horaires'][jour]['arrivee']
            depart = st.session_state.data['horaires'][jour]['depart']
            
            if arrivee and depart:
                try:
                    minutes = convert_to_minutes(arrivee)
                    end = convert_to_minutes(depart)
                    duration = end - minutes if end > minutes else (1440 - minutes) + end
                    total_minutes += duration
                    formatted = format_time(duration)
                    data.append([f"{jour} {date_str}", arrivee, depart, formatted])
                except:
                    data.append([f"{jour} {date_str}", "Erreur", "Format invalide", "0h00"])
            else:
                data.append([f"{jour} {date_str}", "Non", "renseigné", "0h00"])
    
    # Création du DataFrame
    df = pd.DataFrame(data, columns=["Jour", "Arrivée", "Départ", "Temps travaillé"])
    
    # Affichage dans la page
    st.write("### Planning Complet")
    st.dataframe(df)
    
    # Calcul du total
    total_heures = total_minutes // 60
    total_minutes_rest = total_minutes % 60
    st.metric("Total hebdomadaire", f"{total_heures}h{total_minutes_rest:02d}")
    
    # Section Export
    st.markdown("---")
    st.subheader("Exporter le planning")
    
    col1, col2 = st.columns(2)
    
    with col1:
        try:
            excel_data = generate_excel(df)
            st.download_button(
                "📊 Télécharger Excel",
                excel_data,
                file_name=f"planning_{week_dates[0].strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erreur lors de la génération Excel: {str(e)}")
    
    with col2:
        html_data = generate_html(df)
        st.download_button(
            "🌐 Télécharger HTML",
            html_data,
            file_name=f"planning_{week_dates[0].strftime('%Y%m%d')}.html",
            mime="text/html"
        )

    # Export CSV additionnel
    st.download_button(
        "📝 Télécharger CSV",
        df.to_csv(index=False, sep=";"),
        file_name=f"planning_{week_dates[0].strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )
