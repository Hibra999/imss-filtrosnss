import pandas as pd
import json
import os

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
    df.dropna(subset=['NOMSOLI', 'NOMHOSP'], inplace=True)
    df = df.iloc[1:].copy() # Eliminar la fila extra de encabezado si aplica

    print("Procesando datos...")
    
    # Concatenar NSS y AGREGADO en una sola columna
    df['NSS_Limpio'] = df['NSS'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df['AGREGADO_Limpio'] = df['AGREGADO'].astype(str).str.strip()
    df['NSS_AGREGADO'] = df['NSS_Limpio'] + "_" + df['AGREGADO_Limpio']
    
    # Asegurar que fechas y horas sean strings para evitar errores de serialización JSON
    for col in ['FECHASOLICITUD', 'FECHACITA', 'HORACITA']:
        if col in df.columns:
            df[col] = df[col].astype(str)

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

    # Obtener la lista única de Hospitales
    hospitales_unicos = df_umf['NOMHOSP'].dropna().unique().tolist()
    hospitales_unicos = sorted([str(h) for h in hospitales_unicos])

    print("Calculando agregaciones por hospital...")
    data_store = {}

    for hosp in hospitales_unicos:
        df_hosp = df_umf[df_umf['NOMHOSP'] == hosp].copy()
        
        # 1. Solicitudes por UMF
        sol_umf = df_hosp.groupby('NOMSOLI').size().reset_index(name='Total_Solicitudes').sort_values(by='Total_Solicitudes', ascending=False)
        solicitudes_dict = sol_umf.to_dict(orient='records')
        
        # 2. Especialidad
        esp_umf = df_hosp.groupby(['NOMSOLI', 'nomServ']).size().reset_index(name='Total_Por_Especialidad')
        top_20 = sol_umf['NOMSOLI'].head(20).tolist()
        esp_top_20 = esp_umf[esp_umf['NOMSOLI'].isin(top_20)]
        esp_dict = esp_top_20.to_dict(orient='records')
        
        # 3. Duplicados (Agrupando NSS_AGREGADO y nomServ, como en el original)
        citas_paciente = df_hosp.groupby(['NOMSOLI', 'NSS_AGREGADO', 'nomServ']).size().reset_index(name='Num_Citas_Paciente')
        pacientes_dup = citas_paciente[citas_paciente['Num_Citas_Paciente'] > 1].copy()
        pacientes_dup['Citas_Extras'] = pacientes_dup['Num_Citas_Paciente'] - 1
        
        if not pacientes_dup.empty:
            resumen_dup = pacientes_dup.groupby('NOMSOLI').agg(
                Pacientes_Con_Multiples_Citas=('NSS_AGREGADO', 'count'),
                Total_Citas_Duplicadas=('Citas_Extras', 'sum')
            ).reset_index().sort_values(by='Total_Citas_Duplicadas', ascending=False)
            
            df_dup_detalle = df_hosp.merge(pacientes_dup[['NOMSOLI', 'NSS_AGREGADO', 'nomServ', 'Num_Citas_Paciente']], on=['NOMSOLI', 'NSS_AGREGADO', 'nomServ'], how='inner')
            df_dup_detalle.sort_values(by=['NOMSOLI', 'NSS_AGREGADO', 'FECHASOLICITUD'], inplace=True)
            
            cols_mostrar = ["NOMSOLI", "NSS_AGREGADO", "NOMBRE", "nomServ", "FECHACITA", "HORACITA", "Num_Citas_Paciente"]
            for col in cols_mostrar:
                if col not in df_dup_detalle.columns:
                    df_dup_detalle[col] = ""

            # Manejar nulos en el json
            df_mostrar_detalles = df_dup_detalle[cols_mostrar].fillna("")
        else:
            resumen_dup = pd.DataFrame(columns=['NOMSOLI', 'Pacientes_Con_Multiples_Citas', 'Total_Citas_Duplicadas'])
            df_mostrar_detalles = pd.DataFrame(columns=["NOMSOLI", "NSS_AGREGADO", "NOMBRE", "nomServ", "FECHACITA", "HORACITA", "Num_Citas_Paciente"])

        resumen_duplicadas_dict = resumen_dup.to_dict(orient='records')
        detalles_dict = df_mostrar_detalles.to_dict(orient='records')

        # Tabla Resumen Total
        resumen_final = sol_umf.merge(resumen_dup, on='NOMSOLI', how='left').fillna(0)
        resumen_final['Pacientes_Con_Multiples_Citas'] = resumen_final['Pacientes_Con_Multiples_Citas'].astype(int)
        resumen_final['Total_Citas_Duplicadas'] = resumen_final['Total_Citas_Duplicadas'].astype(int)
        resumen_final_dict = resumen_final.to_dict(orient='records')
        
        total_citas_duplicadas = int(resumen_final['Total_Citas_Duplicadas'].sum()) if not resumen_final.empty else 0

        data_store[hosp] = {
            "kpis": {
                "total_citas": len(df_hosp),
                "citas_duplicadas": total_citas_duplicadas,
                "total_umf": df_hosp['NOMSOLI'].nunique()
            },
            "chart1_solicitudes": solicitudes_dict,
            "chart2_especialidad": esp_dict,
            "chart3_duplicados": resumen_duplicadas_dict,
            "tabla_resumen": resumen_final_dict,
            "tabla_detalles": detalles_dict
        }

    # Seleccionar 'UMAA 198 SAN RAFAEL' por defecto si existe, o el primero
    default_hosp = 'UMAA 198 SAN RAFAEL'
    if default_hosp not in hospitales_unicos and len(hospitales_unicos) > 0:
        default_hosp = hospitales_unicos[0]

    json_data = json.dumps(data_store, ensure_ascii=False)
    hospitales_json = json.dumps(hospitales_unicos, ensure_ascii=False)

    print("Generando archivo HTML estático interactivo...")
    
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
        <!-- Plotly JS -->
        <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
        
        <style>
            body {{ background-color: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
            .card {{ border-radius: 12px; border: 1px solid #e0e0e0; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 30px; background-color: white; }}
            h1, h2, h3, h4 {{ font-weight: 600; color: #134e39; }}
            .kpi-card {{ border-radius: 12px; padding: 20px; text-align: center; background: white; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-bottom: 5px solid #134e39; transition: all 0.3s ease; }}
            .kpi-card:hover {{ transform: translateY(-5px); box-shadow: 0 8px 15px rgba(0,0,0,0.1); }}
            .kpi-card h2 {{ font-size: 2.5rem; margin-bottom: 0; color: #134e39; font-weight: 700; transition: color 0.3s; }}
            .kpi-card p {{ color: #555; font-size: 1rem; margin-top: 10px; text-transform: uppercase; letter-spacing: 1px; font-weight: 500; }}
            .kpi-gris {{ border-bottom-color: #8a9597; }}
            .kpi-gris h2 {{ color: #4a4a4a; }}
            .kpi-claro {{ border-bottom-color: #006455; }}
            .kpi-claro h2 {{ color: #006455; }}
            nav.navbar {{ background-color: #ffffff !important; border-bottom: 3px solid #134e39; box-shadow: 0 2px 10px rgba(0,0,0,0.05); padding: 15px 0; }}
            .navbar-brand {{ color: #134e39 !important; font-weight: 800; letter-spacing: 0.5px; }}
            
            /* Ajustes para centrar el contenido y diseño de las tablas estilo IMSS */
            .table thead th {{
                background-color: #134e39 !important;
                color: white !important;
                text-align: center !important;
                vertical-align: middle !important;
                padding: 15px !important;
                font-weight: 600;
                border-bottom: 2px solid #006455 !important;
            }}
            .table tbody td {{
                text-align: center !important;
                vertical-align: middle !important;
                padding: 12px !important;
                color: #333;
            }}
            .table-striped>tbody>tr:nth-of-type(odd)>* {{
                background-color: #f2f7f5 !important;
            }}
            .table-hover>tbody>tr:hover>* {{
                background-color: #e6f0eb !important;
            }}
            .dataTables_wrapper .dataTables_filter input {{
                border-radius: 5px;
                border: 1px solid #bdc3c7;
                padding: 5px 10px;
                margin-left: 10px;
            }}
            
            /* Select Custom Style */
            .hospital-select-container {{
                background: white;
                padding: 20px;
                border-radius: 12px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.05);
                margin-bottom: 30px;
                border-left: 5px solid #134e39;
            }}
            .form-select-lg {{
                border-color: #134e39;
                color: #134e39;
                font-weight: 600;
            }}
            .form-select-lg:focus {{
                border-color: #006455;
                box-shadow: 0 0 0 0.25rem rgba(19, 78, 57, 0.25);
            }}

            .zero-state-message {{
                text-align: center;
                padding: 50px 20px;
                color: #6c757d;
                font-size: 1.2rem;
                background-color: #f8f9fa;
                border-radius: 8px;
                border: 1px dashed #dee2e6;
            }}
        </style>
    </head>
    <body>
        <nav class="navbar navbar-expand-lg navbar-light bg-light mb-4">
            <div class="container">
                <a class="navbar-brand fw-bold" style="color: #2c3e50; font-size: 1.5rem;" href="#">Análisis Estratégico de Citas por Unidad (UMF)</a>
            </div>
        </nav>

        <div class="container">
            <!-- Hospital Selector -->
            <div class="row">
                <div class="col-12">
                    <div class="hospital-select-container d-flex align-items-center justify-content-between flex-wrap">
                        <h4 class="mb-2 mb-md-0">Seleccione un Hospital:</h4>
                        <select id="hospitalSelect" class="form-select form-select-lg w-auto" style="min-width: 300px;">
                            <!-- Opciones inyectadas por JS -->
                        </select>
                    </div>
                </div>
            </div>

            <!-- KPIs -->
            <div class="row mb-4">
                <div class="col-md-4"><div class="kpi-card"><h2 id="kpiTotalCitas">0</h2><p>Total de Citas Evaluadas</p></div></div>
                <div class="col-md-4"><div class="kpi-card kpi-gris"><h2 id="kpiCitasDuplicadas">0</h2><p>Citas Duplicadas Encontradas</p></div></div>
                <div class="col-md-4"><div class="kpi-card kpi-claro"><h2 id="kpiTotalUmf">0</h2><p>Total de Unidades (UMF)</p></div></div>
            </div>

            <!-- Charts -->
            <div class="row">
                <div class="col-12"><div class="card p-4"><div id="divFig1" style="width: 100%; height: 600px;"></div></div></div>
            </div>
            <div class="row">
                <div class="col-12"><div class="card p-4"><div id="divFig2" style="width: 100%; height: 800px;"></div></div></div>
            </div>
            <div class="row">
                <div class="col-12">
                    <div class="card p-4">
                        <div id="divFig3" style="width: 100%; height: 600px;"></div>
                        <div id="zeroStateFig3" class="zero-state-message d-none">
                            <i class="bi bi-check-circle text-success" style="font-size: 2rem;"></i><br>
                            Excelente: No se encontraron citas duplicadas en el hospital seleccionado.
                        </div>
                    </div>
                </div>
            </div>

            <!-- Resumen Table -->
            <div class="row">
                <div class="col-12">
                    <div class="card p-4">
                        <h3 class="mb-4 text-center">Resumen Ejecutivo por Unidad</h3>
                        <div class="table-responsive">
                            <table id="tablaResumenUMF" class="table table-striped table-hover display table-bordered text-center w-100">
                                <thead>
                                    <tr>
                                        <th>Unidad (UMF)</th>
                                        <th>Total Solicitudes Emitidas</th>
                                        <th>Pacientes c/ Múltiples Citas</th>
                                        <th>Total Citas Duplicadas</th>
                                    </tr>
                                </thead>
                                <tbody></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Detalle Table -->
            <div class="row">
                <div class="col-12">
                    <div class="card p-4">
                        <h3 class="mb-4 text-center">Detalle de Pacientes con Citas Duplicadas / Triplicadas</h3>
                        <div id="zeroStateTable" class="zero-state-message d-none mb-3">
                            No hay datos de duplicados para mostrar en esta tabla.
                        </div>
                        <div class="table-responsive" id="tableDetalleContainer">
                            <table id="tablaDetallesDuplicados" class="table table-striped table-hover display table-bordered text-center w-100">
                                <thead>
                                    <tr>
                                        <th>Unidad (UMF)</th>
                                        <th>NSS + AGREGADO</th>
                                        <th>Nombre del Paciente</th>
                                        <th>Especialidad</th>
                                        <th>Fecha Cita</th>
                                        <th>Hora Cita</th>
                                        <th>Total de Citas Encontradas</th>
                                    </tr>
                                </thead>
                                <tbody></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/dataTables.buttons.min.js"></script>
        <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
        <script src="https://cdn.datatables.net/buttons/2.4.1/js/buttons.html5.min.js"></script>
        
        <script>
            // Inyectar datos desde Python
            const DATA_STORE = {json_data};
            const HOSPITALES = {hospitales_json};
            const DEFAULT_HOSPITAL = "{default_hosp}";

            let dtResumen = null;
            let dtDetalles = null;

            $(document).ready(function() {{
                // Llenar el select de hospitales
                const $select = $('#hospitalSelect');
                HOSPITALES.forEach(h => {{
                    const isSelected = (h === DEFAULT_HOSPITAL) ? 'selected' : '';
                    $select.append(`<option value="${{h}}" ${{isSelected}}>${{h}}</option>`);
                }});

                // Inicializar DataTables
                dtResumen = $('#tablaResumenUMF').DataTable({{
                    dom: 'Bfrtip',
                    buttons: ['copyHtml5', 'excelHtml5', 'csvHtml5'],
                    language: {{ url: "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" }},
                    pageLength: 10,
                    order: [[ 3, "desc" ]]
                }});

                dtDetalles = $('#tablaDetallesDuplicados').DataTable({{
                    dom: 'Bfrtip',
                    buttons: ['copyHtml5', 'excelHtml5', 'csvHtml5'],
                    language: {{ url: "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" }},
                    pageLength: 15,
                    order: [[ 0, "asc" ], [1, "asc"]]
                }});

                // Actualizar dashboard por primera vez
                if(DATA_STORE[DEFAULT_HOSPITAL]) {{
                    updateDashboard(DEFAULT_HOSPITAL);
                }}

                // Evento al cambiar el hospital
                $select.on('change', function() {{
                    const selected = $(this).val();
                    if(DATA_STORE[selected]) {{
                        updateDashboard(selected);
                    }}
                }});
            }});

            function updateDashboard(hospital) {{
                const data = DATA_STORE[hospital];
                
                // 1. Actualizar KPIs con animación
                animateValue("kpiTotalCitas", parseInt($('#kpiTotalCitas').text()), data.kpis.total_citas, 500);
                animateValue("kpiCitasDuplicadas", parseInt($('#kpiCitasDuplicadas').text()), data.kpis.citas_duplicadas, 500);
                animateValue("kpiTotalUmf", parseInt($('#kpiTotalUmf').text()), data.kpis.total_umf, 500);

                // 2. Gráfica 1: Solicitudes
                const trace1 = {{
                    x: data.chart1_solicitudes.map(d => d.NOMSOLI),
                    y: data.chart1_solicitudes.map(d => d.Total_Solicitudes),
                    type: 'bar',
                    marker: {{color: '#134e39'}},
                    text: data.chart1_solicitudes.map(d => String(d.Total_Solicitudes)),
                    textposition: 'auto'
                }};
                Plotly.react('divFig1', [trace1], {{
                    title: '1. Solicitudes Totales por Unidad (UMF)',
                    xaxis: {{tickangle: -45, title: 'Unidad (UMF)'}},
                    yaxis: {{title: 'Cantidad de Solicitudes Emitidas'}},
                    margin: {{b: 150}},
                    plot_bgcolor: "white", paper_bgcolor: "white"
                }});

                // 3. Gráfica 2: Especialidad
                const epsData = data.chart2_especialidad;
                const especialidades = [...new Set(epsData.map(d => d.nomServ))];
                const traces2 = especialidades.map(esp => {{
                    const filtered = epsData.filter(d => d.nomServ === esp);
                    return {{
                        x: filtered.map(d => d.NOMSOLI),
                        y: filtered.map(d => d.Total_Por_Especialidad),
                        name: esp,
                        type: 'bar'
                    }};
                }});
                Plotly.react('divFig2', traces2, {{
                    title: '2. Solicitudes por Especialidad en el Top 20 de Unidades',
                    barmode: 'stack',
                    xaxis: {{tickangle: -45, title: 'Unidad'}},
                    yaxis: {{title: 'Número de Solicitudes'}},
                    margin: {{b: 150}},
                    plot_bgcolor: "white", paper_bgcolor: "white"
                }});

                // 4. Gráfica 3: Duplicados
                if(data.chart3_duplicados.length > 0) {{
                    $('#divFig3').removeClass('d-none');
                    $('#zeroStateFig3').addClass('d-none');
                    const trace3 = {{
                        x: data.chart3_duplicados.map(d => d.NOMSOLI),
                        y: data.chart3_duplicados.map(d => d.Total_Citas_Duplicadas),
                        type: 'bar',
                        marker: {{color: '#006455'}},
                        text: data.chart3_duplicados.map(d => String(d.Total_Citas_Duplicadas)),
                        textposition: 'auto'
                    }};
                    Plotly.react('divFig3', [trace3], {{
                        title: '3. Cantidad Total de Citas Duplicadas/Triplicadas por Unidad',
                        xaxis: {{tickangle: -45, title: 'Unidad (UMF)'}},
                        yaxis: {{title: 'Citas Duplicadas'}},
                        margin: {{b: 150}},
                        plot_bgcolor: "white", paper_bgcolor: "white"
                    }});
                }} else {{
                    $('#divFig3').addClass('d-none');
                    $('#zeroStateFig3').removeClass('d-none');
                }}

                // 5. Tabla Resumen
                dtResumen.clear();
                data.tabla_resumen.forEach(r => {{
                    dtResumen.row.add([
                        r.NOMSOLI,
                        r.Total_Solicitudes,
                        r.Pacientes_Con_Multiples_Citas,
                        r.Total_Citas_Duplicadas
                    ]);
                }});
                dtResumen.draw();

                // 6. Tabla Detalles
                if(data.tabla_detalles.length > 0) {{
                    $('#tableDetalleContainer').removeClass('d-none');
                    $('#zeroStateTable').addClass('d-none');
                    dtDetalles.clear();
                    data.tabla_detalles.forEach(r => {{
                        dtDetalles.row.add([
                            r.NOMSOLI,
                            r.NSS_AGREGADO,
                            r.NOMBRE,
                            r.nomServ,
                            r.FECHACITA,
                            r.HORACITA,
                            r.Num_Citas_Paciente
                        ]);
                    }});
                    dtDetalles.draw();
                }} else {{
                    $('#tableDetalleContainer').addClass('d-none');
                    $('#zeroStateTable').removeClass('d-none');
                }}
            }}

            // Función para animar números de KPIs
            function animateValue(id, start, end, duration) {{
                if (start === end) return;
                let range = end - start;
                let current = start;
                let increment = end > start ? 1 : -1;
                let stepTime = Math.abs(Math.floor(duration / range));
                if(stepTime < 10) stepTime = 10;
                let obj = document.getElementById(id);
                let timer = setInterval(function() {{
                    current += increment;
                    obj.innerHTML = current;
                    if (current == end) {{
                        clearInterval(timer);
                    }}
                }}, stepTime);
            }}
        </script>
        <!-- FontAwesome/Bootstrap icons if needed -->
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css">
    </body>
    </html>
    """

    os.makedirs("docs", exist_ok=True)
    with open("docs/index.html", "w", encoding="utf-8") as f:
        f.write(html_template)
    
    print(f"\n¡Éxito! El reporte se ha guardado como 'docs/index.html' con {len(hospitales_unicos)} hospitales procesados.")

if __name__ == "__main__":
    main()
