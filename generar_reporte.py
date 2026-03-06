import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import os

# Plantilla GGPLOT2 para Plotly (Fondo blanco y estilo ggplot)
pio.templates.default = "ggplot2"

def main():
    print("Leyendo el archivo de Excel...")
    try:
        df = pd.read_excel("DF72F2A6-export.xlsx", sheet_name="ListadoCitasComprometidas")
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return

    # Limpieza inicial de nombres de columnas
    df.columns = [
        "CVE_DEL", "SOLICITO", "NOMSOLI", "FECHASOLICITUD", "HOSPITAL", "NOMHOSP", "SERVICIO", 
        "nomServ", "FECHACITA", "CONSULTORIO", "TURNO", "HORACITA", "NSS", "AGREGADO", 
        "NOMBRE", "PATERNO", "MATERNO", "DiasHabilRefer_Cita", "TEL", "CEL", "MAIL"
    ]
    df.dropna(inplace=True)
    df = df.iloc[1:] # Eliminar la fila extra de encabezado si aplica

    print("Procesando datos...")
    
    # 1. Concatenar NSS y AGREGADO en una sola columna
    df['NSS_Limpio'] = df['NSS'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['AGREGADO_Limpio'] = df['AGREGADO'].astype(str).str.strip()
    
    df['NSS_AGREGADO'] = df['NSS_Limpio'] + "_" + df['AGREGADO_Limpio']

    # Unidades especificadas en el contexto
    unidades_umf = [
        'UMF 78 NETZAHUALCOYOTL', 'UMF 75 NETZAHUALCOYOTL', 'UMF 62 CUAUTITLAN', 
        'UMF 64 TEQUESQUINAHUAC', 'UMF 198 COACALCO', 'UMF 79 V.CEYLAN', 'HGOMF 60 TLANEPANTLA', 
        'UMF 92 CD.AZTECA', 'UMF 95 PANTACO', 'UMF 195 CHALCO', 'UMAA 199 TLANEPANTLA',
        'UMF 91 VILLA FLORES', 'UMF 191 ECATEPEC', 'HGZMF 76 XALOSTOC', 'UMF 186 IZTACALA', 
        'UMF 56 JILOTEPEC', 'UMF 185 L CARTAGENA', 'UMF 52 CUAUTITLAN I.', 'UMF 55 ZUMPANGO', 
        'HGZ 57 LA QUEBRADA', 'UMF 181 CHALCO II', 'UMF 68 TULPETLAC', 'HGZ 98 COACALCO',
        'HGR 72 GUSTAVO BAZ', 'UMF 59 LECHERIA', 'UMF 184 INFONAVIT SUR', 'UMF 93 CERRO GORDO', 
        'UMF 77 SAN AGUSTIN', 'HGR 196 FIDEL VELAZQUEZ', 'UMAA 198 SAN RAFAEL', 'UMF 67 STA.CLARA', 
        'UMF 69 TEXCOCO', 'UMF 188 TEPALCATES', 'HGZ 68 TULPETLAC', 'UMF 54 APAZCO', 'UMF 70 AYOTLA',
        'UMF 84 CHIMALHUACAN', 'UMF 86 IXTAPALUCA', 'UMF 180 CHALCO I', 'UMF 87 OZUMBA', 
        'UMF 193 CHALCO', 'UMF 74 SAN RAFAEL', 'UMF 73 AMECAMECA', 'UMF 81 JUCHITEPEC', 
        'UMAA 180 CHALCO I ', 'UMF 182 EL SOL', 'UMF 96 TEPOZA', 'UMF 83 CHICOLOAPAN',
        'HGZ 53 LOS REYES PAZ', 'HGZ 197 TEXCOCO', 'UMF 189 CHIMALHUACAN', 'UMF 82 ATENCO', 
        'UMF 183 REY NETZA', 'HGR 200 TECAMAC', 'UMF 89 OTUMBA'
    ]

    df_umf = df[df['NOMSOLI'].isin(unidades_umf)].copy()

    # ---------------------------------------------------------
    # 2. Conteo de solicitudes y por especialidad de CADA UNIDAD
    # ---------------------------------------------------------
    solicitudes_por_unidad = df_umf.groupby('NOMSOLI').size().reset_index(name='Total_Solicitudes').sort_values(by='Total_Solicitudes', ascending=False)
    especialidad_por_unidad = df_umf.groupby(['NOMSOLI', 'nomServ']).size().reset_index(name='Total_Por_Especialidad')

    # ---------------------------------------------------------
    # 3. VER CUANTAS CITAS DUPLICADAS O TRIPLICADAS POR UNIDAD
    #    (Basado en NSS_AGREGADO y también nomServ - Especialidad)
    # ---------------------------------------------------------
    citas_por_paciente = df_umf.groupby(['NOMSOLI', 'NSS_AGREGADO', 'nomServ']).size().reset_index(name='Num_Citas_Paciente')
    pacientes_duplicados = citas_por_paciente[citas_por_paciente['Num_Citas_Paciente'] > 1].copy()
    pacientes_duplicados['Citas_Extras'] = pacientes_duplicados['Num_Citas_Paciente'] - 1
    
    resumen_duplicadas_umf = pacientes_duplicados.groupby('NOMSOLI').agg(
        Pacientes_Con_Multiples_Citas=('NSS_AGREGADO', 'count'),
        Total_Citas_Duplicadas=('Citas_Extras', 'sum')
    ).reset_index().sort_values(by='Total_Citas_Duplicadas', ascending=False)

    df_duplicados_detalle = df_umf.merge(pacientes_duplicados[['NOMSOLI', 'NSS_AGREGADO', 'nomServ', 'Num_Citas_Paciente']], on=['NOMSOLI', 'NSS_AGREGADO', 'nomServ'], how='inner')
    df_duplicados_detalle.sort_values(by=['NOMSOLI', 'NSS_AGREGADO', 'FECHASOLICITUD'], inplace=True)

    print("Generando gráficos...")

    fig_totales = px.bar(
        solicitudes_por_unidad, x='NOMSOLI', y='Total_Solicitudes',
        title="1. Solicitudes Totales por Unidad (UMF)",
        labels={'NOMSOLI': 'Unidad (UMF)', 'Total_Solicitudes': 'Cantidad de Solicitudes Emitidas'},
        color='Total_Solicitudes', color_continuous_scale="Viridis", text_auto=True
    ).update_layout(plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-45, margin=dict(b=150))

    top_20_umf = solicitudes_por_unidad['NOMSOLI'].head(20).tolist()
    esp_top_20 = especialidad_por_unidad[especialidad_por_unidad['NOMSOLI'].isin(top_20_umf)]
    
    fig_especialidad = px.bar(
        esp_top_20, x='NOMSOLI', y='Total_Por_Especialidad', color='nomServ',
        title="2. Solicitudes por Especialidad en el Top 20 de Unidades",
        labels={'NOMSOLI': 'Unidad', 'Total_Por_Especialidad': 'Número de Solicitudes', 'nomServ': 'Especialidad'},
        barmode='stack'
    ).update_layout(plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-45, margin=dict(b=150))

    fig_duplicados = px.bar(
        resumen_duplicadas_umf, x='NOMSOLI', y='Total_Citas_Duplicadas',
        title="3. Cantidad Total de Citas Duplicadas/Triplicadas por Unidad",
        labels={'NOMSOLI': 'Unidad (UMF)', 'Total_Citas_Duplicadas': 'Citas Duplicadas'},
        color='Total_Citas_Duplicadas', color_continuous_scale="Reds", text_auto=True
    ).update_layout(plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-45, margin=dict(b=150))

    html_fig1 = pio.to_html(fig_totales, full_html=False, default_width='100%', default_height='600px')
    html_fig2 = pio.to_html(fig_especialidad, full_html=False, default_width='100%', default_height='800px')
    html_fig3 = pio.to_html(fig_duplicados, full_html=False, default_width='100%', default_height='600px')

    resumen_final = solicitudes_por_unidad.merge(resumen_duplicadas_umf, on='NOMSOLI', how='left').fillna(0)
    resumen_final['Total_Citas_Duplicadas'] = resumen_final['Total_Citas_Duplicadas'].astype(int)
    resumen_final.columns = ["Unidad (UMF)", "Total Solicitudes Emitidas", "Pacientes c/ Múltiples Citas", "Total Citas Duplicadas"]
    html_tabla_resumen = resumen_final.to_html(classes="table table-striped table-hover display", index=False, table_id="tablaResumenUMF")

    columnas_mostrar = ["NOMSOLI", "NSS_AGREGADO", "NOMBRE", "nomServ", "FECHACITA", "HORACITA", "Num_Citas_Paciente"]
    df_mostrar_detalles = df_duplicados_detalle[columnas_mostrar]
    df_mostrar_detalles.columns = ["Unidad (UMF)", "NSS + AGREGADO", "Nombre del Paciente", "Especialidad", "Fecha Cita", "Hora Cita", "Total de Citas Creadas"]
    html_tabla_detalles = df_mostrar_detalles.to_html(classes="table table-striped table-hover display", index=False, table_id="tablaDetallesDuplicados")

    print("Generando archivo HTML estático...")
    
    html_template = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard Citas UMF y Duplicados</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css">
        <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/buttons/2.4.1/css/buttons.dataTables.min.css">
        <style>
            body {{ background-color: #f4f7f6; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
            .card {{ border-radius: 15px; border: none; box-shadow: 0 10px 20px rgba(0,0,0,0.05); margin-bottom: 30px; background-color: white; }}
            h1, h2, h3, h4 {{ font-weight: 700; color: #2c3e50; }}
            .kpi-card {{ border-radius: 15px; padding: 20px; text-align: center; background: linear-gradient(135deg, #ffffff 0%, #f9f9f9 100%); box-shadow: 0 4px 15px rgba(0,0,0,0.05); border-bottom: 5px solid #3498db; }}
            .kpi-card h2 {{ font-size: 2.5rem; margin-bottom: 0; color: #2980b9; }}
            .kpi-card p {{ color: #7f8c8d; font-size: 1.1rem; margin-top: 10px; text-transform: uppercase; letter-spacing: 1px; }}
            .kpi-rojo {{ border-bottom-color: #e74c3c; }}
            .kpi-rojo h2 {{ color: #c0392b; }}
            .kpi-verde {{ border-bottom-color: #2ecc71; }}
            .kpi-verde h2 {{ color: #27ae60; }}
            nav.navbar {{ background-color: #ffffff !important; box-shadow: 0 2px 10px rgba(0,0,0,0.05); padding: 15px 0; }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-expand-lg navbar-light bg-light mb-5">
            <div class="container">
                <a class="navbar-brand fw-bold" style="color: #2c3e50; font-size: 1.5rem;" href="#">Análisis Estratégico de Citas por Unidad (UMF)</a>
            </div>
        </nav>

        <div class="container">
            <div class="row mb-4">
                <div class="col-md-4"><div class="kpi-card"><h2>{len(df_umf)}</h2><p>Total de Citas Evaluadas</p></div></div>
                <div class="col-md-4"><div class="kpi-card kpi-rojo"><h2>{resumen_final['Total Citas Duplicadas'].sum()}</h2><p>Citas Duplicadas Encontradas</p></div></div>
                <div class="col-md-4"><div class="kpi-card kpi-verde"><h2>{df_umf['NOMSOLI'].nunique()}</h2><p>Total de Unidades (UMF)</p></div></div>
            </div>

            <div class="row"><div class="col-12"><div class="card p-4">{html_fig1}</div></div></div>
            <div class="row"><div class="col-12"><div class="card p-4">{html_fig2}</div></div></div>
            <div class="row"><div class="col-12"><div class="card p-4">{html_fig3}</div></div></div>

            <div class="row">
                <div class="col-12">
                    <div class="card p-4">
                        <h3 class="mb-4 text-center">Resumen Ejecutivo por Unidad</h3>
                        <div class="table-responsive">
                            {html_tabla_resumen}
                        </div>
                    </div>
                </div>
            </div>

            <div class="row">
                <div class="col-12">
                    <div class="card p-4">
                        <h3 class="mb-4 text-center">Detalle de Pacientes con Citas Duplicadas / Triplicadas</h3>
                        <div class="table-responsive">
                            {html_tabla_detalles}
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
        <script src="https://cdnt.datatables.net/1.13.6/js/jquery.dataTables.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
        <script>
            $(document).ready(function() {{
                $('#tablaResumenUMF').DataTable({{
                    dom: 'Bfrtip',
                    buttons: ['copyHtml5', 'excelHtml5', 'csvHtml5'],
                    language: {{ url: "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" }},
                    pageLength: 10,
                    order: [[ 3, "desc" ]]
                }});

                $('#tablaDetallesDuplicados').DataTable({{
                    dom: 'Bfrtip',
                    buttons: ['copyHtml5', 'excelHtml5', 'csvHtml5'],
                    language: {{ url: "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" }},
                    pageLength: 15,
                    order: [[ 0, "asc" ], [1, "asc"]]
                }});
            }});
        </script>
    </body>
    </html>
    """

    with open("Reporte_Dashboard_Citas.html", "w", encoding="utf-8") as f:
        f.write(html_template)
    
    print("\n¡Éxito! El reporte se ha guardado como 'Reporte_Dashboard_Citas.html'")

if __name__ == "__main__":
    main()
