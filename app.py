from flask import Flask, request, jsonify, render_template, send_file
import anthropic
import os
import json
import sqlite3
import base64
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB max upload
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ─── BASE DE DATOS ────────────────────────────────────────────────────────────
def get_db():
    db = sqlite3.connect('propuestas.db')
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    db.executescript('''
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            razon_social TEXT NOT NULL,
            nit TEXT,
            tipo TEXT,
            direccion TEXT,
            ciudad TEXT,
            telefono TEXT,
            email TEXT,
            rep_legal TEXT,
            cc_rep_legal TEXT,
            camara_comercio TEXT,
            fecha_vencimiento_camara TEXT,
            contador_nombre TEXT,
            contador_cc TEXT,
            contador_tp TEXT,
            contador_camara TEXT,
            revisor_nombre TEXT,
            revisor_cc TEXT,
            revisor_tp TEXT,
            revisor_camara TEXT,
            capital_trabajo TEXT,
            patrimonio TEXT,
            liquidez TEXT,
            endeudamiento TEXT,
            rentabilidad TEXT,
            tiene_discapacidad INTEGER DEFAULT 0,
            tiene_mujeres INTEGER DEFAULT 0,
            es_mipyme INTEGER DEFAULT 0,
            exonerada_parafiscales INTEGER DEFAULT 0,
            fecha_constitucion TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS experiencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER,
            entidad TEXT,
            objeto TEXT,
            valor TEXT,
            anio TEXT,
            plazo TEXT,
            acta TEXT,
            unspsc TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id)
        );
        CREATE TABLE IF NOT EXISTS documentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER,
            tipo TEXT,
            nombre_archivo TEXT,
            ruta TEXT,
            fecha_subida TEXT,
            fecha_vencimiento TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id)
        );
    ''')
    db.commit()
    db.close()

init_db()

# ─── CLAUDE API ───────────────────────────────────────────────────────────────
def call_claude(prompt, pdf_base64=None, media_type="application/pdf"):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    messages_content = []
    if pdf_base64:
        messages_content.append({
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": pdf_base64
            }
        })
    messages_content.append({"type": "text", "text": prompt})
    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        messages=[{"role": "user", "content": messages_content}]
    )
    return response.content[0].text

# ─── RUTAS PRINCIPALES ────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html')

# ─── EMPRESAS ─────────────────────────────────────────────────────────────────
@app.route('/api/empresas', methods=['GET'])
def get_empresas():
    db = get_db()
    empresas = db.execute('SELECT * FROM empresas ORDER BY razon_social').fetchall()
    result = []
    today = datetime.now().strftime('%Y-%m-%d')
    for e in empresas:
        emp = dict(e)
        # Verificar vigencia documentos
        docs = db.execute('SELECT * FROM documentos WHERE empresa_id=?', (e['id'],)).fetchall()
        emp['documentos'] = [dict(d) for d in docs]
        emp['alertas'] = []
        for d in docs:
            if d['fecha_vencimiento']:
                dias = (datetime.strptime(d['fecha_vencimiento'], '%Y-%m-%d') - datetime.now()).days
                if dias < 0:
                    emp['alertas'].append({'tipo': d['tipo'], 'msg': f"VENCIDO hace {abs(dias)} días", 'nivel': 'rojo'})
                elif dias <= 30:
                    emp['alertas'].append({'tipo': d['tipo'], 'msg': f"Vence en {dias} días", 'nivel': 'amarillo'})
        # Experiencia
        exp = db.execute('SELECT * FROM experiencia WHERE empresa_id=?', (e['id'],)).fetchall()
        emp['experiencia'] = [dict(x) for x in exp]
        result.append(emp)
    db.close()
    return jsonify(result)

@app.route('/api/empresas', methods=['POST'])
def create_empresa():
    data = request.json
    db = get_db()
    cur = db.execute('''INSERT INTO empresas 
        (razon_social,nit,tipo,direccion,ciudad,telefono,email,rep_legal,cc_rep_legal,
         camara_comercio,fecha_vencimiento_camara,contador_nombre,contador_cc,contador_tp,contador_camara,
         revisor_nombre,revisor_cc,revisor_tp,revisor_camara,capital_trabajo,patrimonio,
         liquidez,endeudamiento,rentabilidad,tiene_discapacidad,tiene_mujeres,es_mipyme,
         exonerada_parafiscales,fecha_constitucion)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (data.get('razon_social',''), data.get('nit',''), data.get('tipo',''),
         data.get('direccion',''), data.get('ciudad',''), data.get('telefono',''),
         data.get('email',''), data.get('rep_legal',''), data.get('cc_rep_legal',''),
         data.get('camara_comercio',''), data.get('fecha_vencimiento_camara',''),
         data.get('contador_nombre',''), data.get('contador_cc',''), data.get('contador_tp',''),
         data.get('contador_camara',''), data.get('revisor_nombre',''), data.get('revisor_cc',''),
         data.get('revisor_tp',''), data.get('revisor_camara',''), data.get('capital_trabajo',''),
         data.get('patrimonio',''), data.get('liquidez',''), data.get('endeudamiento',''),
         data.get('rentabilidad',''), data.get('tiene_discapacidad',0), data.get('tiene_mujeres',0),
         data.get('es_mipyme',0), data.get('exonerada_parafiscales',0), data.get('fecha_constitucion','')))
    empresa_id = cur.lastrowid
    db.commit()
    db.close()
    return jsonify({'id': empresa_id, 'ok': True})

@app.route('/api/empresas/<int:empresa_id>', methods=['PUT'])
def update_empresa(empresa_id):
    data = request.json
    db = get_db()
    db.execute('''UPDATE empresas SET razon_social=?,nit=?,tipo=?,direccion=?,ciudad=?,telefono=?,
        email=?,rep_legal=?,cc_rep_legal=?,camara_comercio=?,fecha_vencimiento_camara=?,
        contador_nombre=?,contador_cc=?,contador_tp=?,contador_camara=?,revisor_nombre=?,
        revisor_cc=?,revisor_tp=?,revisor_camara=?,capital_trabajo=?,patrimonio=?,liquidez=?,
        endeudamiento=?,rentabilidad=?,tiene_discapacidad=?,tiene_mujeres=?,es_mipyme=?,
        exonerada_parafiscales=?,fecha_constitucion=? WHERE id=?''',
        (data.get('razon_social',''), data.get('nit',''), data.get('tipo',''),
         data.get('direccion',''), data.get('ciudad',''), data.get('telefono',''),
         data.get('email',''), data.get('rep_legal',''), data.get('cc_rep_legal',''),
         data.get('camara_comercio',''), data.get('fecha_vencimiento_camara',''),
         data.get('contador_nombre',''), data.get('contador_cc',''), data.get('contador_tp',''),
         data.get('contador_camara',''), data.get('revisor_nombre',''), data.get('revisor_cc',''),
         data.get('revisor_tp',''), data.get('revisor_camara',''), data.get('capital_trabajo',''),
         data.get('patrimonio',''), data.get('liquidez',''), data.get('endeudamiento',''),
         data.get('rentabilidad',''), data.get('tiene_discapacidad',0), data.get('tiene_mujeres',0),
         data.get('es_mipyme',0), data.get('exonerada_parafiscales',0), data.get('fecha_constitucion',''),
         empresa_id))
    db.commit()
    db.close()
    return jsonify({'ok': True})

@app.route('/api/empresas/<int:empresa_id>', methods=['DELETE'])
def delete_empresa(empresa_id):
    db = get_db()
    db.execute('DELETE FROM experiencia WHERE empresa_id=?', (empresa_id,))
    db.execute('DELETE FROM documentos WHERE empresa_id=?', (empresa_id,))
    db.execute('DELETE FROM empresas WHERE id=?', (empresa_id,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

# ─── EXPERIENCIA ──────────────────────────────────────────────────────────────
@app.route('/api/empresas/<int:empresa_id>/experiencia', methods=['POST'])
def add_experiencia(empresa_id):
    data = request.json
    db = get_db()
    cur = db.execute('INSERT INTO experiencia (empresa_id,entidad,objeto,valor,anio,plazo,acta,unspsc) VALUES (?,?,?,?,?,?,?,?)',
        (empresa_id, data.get('entidad',''), data.get('objeto',''), data.get('valor',''),
         data.get('anio',''), data.get('plazo',''), data.get('acta',''), data.get('unspsc','')))
    db.commit()
    exp_id = cur.lastrowid
    db.close()
    return jsonify({'id': exp_id, 'ok': True})

@app.route('/api/experiencia/<int:exp_id>', methods=['DELETE'])
def delete_experiencia(exp_id):
    db = get_db()
    db.execute('DELETE FROM experiencia WHERE id=?', (exp_id,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

# ─── SUBIR DOCUMENTOS Y EXTRAER CON IA ───────────────────────────────────────
@app.route('/api/empresas/<int:empresa_id>/subir-documento', methods=['POST'])
def subir_documento(empresa_id):
    tipo = request.form.get('tipo')
    archivo = request.files.get('archivo')
    if not archivo:
        return jsonify({'error': 'No se recibió archivo'}), 400

    # Guardar archivo
    filename = f"{empresa_id}_{tipo}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{archivo.filename}"
    ruta = os.path.join(UPLOAD_FOLDER, filename)
    archivo.save(ruta)

    # Leer PDF en base64 para Claude
    with open(ruta, 'rb') as f:
        pdf_b64 = base64.standard_b64encode(f.read()).decode('utf-8')

    # Prompts según tipo de documento
    prompts = {
        'camara': '''Analiza este documento de Cámara de Comercio colombiano y extrae en formato JSON exacto:
{
  "razon_social": "",
  "nit": "",
  "tipo": "",
  "direccion": "",
  "ciudad": "",
  "rep_legal": "",
  "cc_rep_legal": "",
  "objeto_social": "",
  "fecha_matricula": "",
  "fecha_renovacion": "",
  "fecha_vencimiento": "YYYY-MM-DD"
}
Solo responde el JSON, sin texto adicional.''',

        'rup': '''Analiza este Registro Único de Proponentes (RUP) colombiano y extrae en formato JSON exacto:
{
  "razon_social": "",
  "nit": "",
  "capital_trabajo": "",
  "patrimonio": "",
  "liquidez": "",
  "endeudamiento": "",
  "rentabilidad": "",
  "fecha_inscripcion": "",
  "fecha_vencimiento": "YYYY-MM-DD",
  "clasificacion_unspsc": []
}
Solo responde el JSON, sin texto adicional.''',

        'cedula': '''Analiza esta cédula de ciudadanía colombiana y extrae en formato JSON:
{
  "nombre_completo": "",
  "numero_cedula": "",
  "fecha_expedicion": ""
}
Solo responde el JSON, sin texto adicional.''',

        'tarjeta_profesional': '''Analiza esta tarjeta profesional colombiana y extrae en formato JSON:
{
  "nombre_completo": "",
  "numero_tp": "",
  "profesion": "",
  "fecha_expedicion": "",
  "fecha_vencimiento": "YYYY-MM-DD"
}
Solo responde el JSON, sin texto adicional.''',

        'cert_discapacidad': '''Analiza este certificado del Ministerio de Trabajo de Colombia sobre personas con discapacidad y extrae en JSON:
{
  "empresa": "",
  "total_trabajadores": "",
  "trabajadores_discapacidad": "",
  "fecha_expedicion": "",
  "fecha_vencimiento": "YYYY-MM-DD"
}
Solo responde el JSON, sin texto adicional.''',

        'acta_accionaria': '''Analiza este acta de composición accionaria colombiana y extrae en JSON:
{
  "empresa": "",
  "porcentaje_mujeres": "",
  "fecha_acta": "",
  "socias": []
}
Solo responde el JSON, sin texto adicional.'''
    }

    prompt = prompts.get(tipo, 'Extrae los datos principales de este documento en formato JSON.')

    try:
        respuesta = call_claude(prompt, pdf_b64)
        # Limpiar respuesta y parsear JSON
        respuesta_limpia = respuesta.strip()
        if '```' in respuesta_limpia:
            respuesta_limpia = respuesta_limpia.split('```')[1]
            if respuesta_limpia.startswith('json'):
                respuesta_limpia = respuesta_limpia[4:]
        datos = json.loads(respuesta_limpia)
    except Exception as e:
        datos = {}

    # Calcular fecha vencimiento para alertas
    fecha_venc = datos.get('fecha_vencimiento', '')

    # Guardar documento en DB
    db = get_db()
    db.execute('INSERT INTO documentos (empresa_id,tipo,nombre_archivo,ruta,fecha_subida,fecha_vencimiento) VALUES (?,?,?,?,?,?)',
        (empresa_id, tipo, archivo.filename, ruta, datetime.now().strftime('%Y-%m-%d'), fecha_venc))

    # Actualizar datos de empresa según tipo
    if tipo == 'camara':
        if datos.get('fecha_vencimiento'):
            db.execute('UPDATE empresas SET fecha_vencimiento_camara=? WHERE id=?', (fecha_venc, empresa_id))
    elif tipo == 'rup':
        db.execute('''UPDATE empresas SET capital_trabajo=?,patrimonio=?,liquidez=?,endeudamiento=?,rentabilidad=? WHERE id=?''',
            (datos.get('capital_trabajo',''), datos.get('patrimonio',''), datos.get('liquidez',''),
             datos.get('endeudamiento',''), datos.get('rentabilidad',''), empresa_id))

    db.commit()
    db.close()
    return jsonify({'ok': True, 'datos_extraidos': datos})

# ─── ANALIZAR PLIEGO ─────────────────────────────────────────────────────────
@app.route('/api/analizar-pliego', methods=['POST'])
def analizar_pliego():
    archivo = request.files.get('pliego')
    if not archivo:
        return jsonify({'error': 'No se recibió el pliego'}), 400

    with open(os.path.join(UPLOAD_FOLDER, 'pliego_temp.pdf'), 'wb') as f:
        archivo.save(f)

    with open(os.path.join(UPLOAD_FOLDER, 'pliego_temp.pdf'), 'rb') as f:
        pdf_b64 = base64.standard_b64encode(f.read()).decode('utf-8')

    prompt = '''Analiza este pliego de condiciones colombiano (pliego tipo de infraestructura de transporte, Resolución 465/2024) y extrae en formato JSON exacto:
{
  "entidad": "",
  "ciudad_entidad": "",
  "direccion_entidad": "",
  "numero_proceso": "",
  "objeto": "",
  "presupuesto_oficial": "",
  "plazo_ejecucion": "",
  "fecha_cierre": "",
  "lote": "",
  "requisitos_habilitantes": {
    "juridicos": [],
    "financieros": {
      "liquidez_minima": "",
      "endeudamiento_maximo": "",
      "rentabilidad_minima": "",
      "capital_trabajo_minimo": ""
    },
    "tecnicos": {
      "experiencia_minima_valor": "",
      "experiencia_unspsc": [],
      "capacidad_residual_minima": ""
    }
  },
  "formatos_habilitantes": ["F1","F2","F6"],
  "formatos_puntaje": {
    "F7A": {"aplica": true, "puntaje": ""},
    "F7B": {"aplica": false, "puntaje": ""},
    "F7C": {"aplica": false, "puntaje": ""},
    "F8": {"aplica": false, "puntaje": ""},
    "F9A": {"aplica": true, "puntaje": "", "bienes_relevantes": []},
    "F9B": {"aplica": false, "puntaje": ""},
    "F12": {"aplica": false, "puntaje": ""},
    "F13": {"aplica": false, "puntaje": ""},
    "F14": {"aplica": false, "puntaje": ""}
  },
  "total_puntaje": "100",
  "formula_consorcio_financiero": ""
}
Solo responde el JSON sin texto adicional. Si un campo no aparece en el pliego déjalo vacío o false.'''

    try:
        respuesta = call_claude(prompt, pdf_b64)
        respuesta_limpia = respuesta.strip()
        if '```' in respuesta_limpia:
            respuesta_limpia = respuesta_limpia.split('```')[1]
            if respuesta_limpia.startswith('json'):
                respuesta_limpia = respuesta_limpia[4:]
        datos_pliego = json.loads(respuesta_limpia)
    except Exception as e:
        return jsonify({'error': f'Error analizando el pliego: {str(e)}'}), 500

    return jsonify({'ok': True, 'pliego': datos_pliego})

# ─── VERIFICAR CUMPLIMIENTO ───────────────────────────────────────────────────
@app.route('/api/verificar-cumplimiento', methods=['POST'])
def verificar_cumplimiento():
    data = request.json
    pliego = data.get('pliego', {})
    empresas_ids = data.get('empresas', [])  # [{id, pct}]
    experiencia_ids = data.get('experiencia_ids', [])

    db = get_db()
    resultado = {'empresas': [], 'cumple_financiero': True, 'cumple_experiencia': True, 'alertas_docs': [], 'formatos_a_generar': []}

    req_fin = pliego.get('requisitos_habilitantes', {}).get('financieros', {})

    for emp_data in empresas_ids:
        emp = db.execute('SELECT * FROM empresas WHERE id=?', (emp_data['id'],)).fetchone()
        if not emp:
            continue
        emp = dict(emp)
        pct = float(emp_data.get('pct', 100)) / 100

        # Verificar indicadores financieros
        indicadores = []
        def chk(label, valor_emp, minimo, mayor_es_mejor=True):
            try:
                v = float(str(valor_emp).replace(',','.').replace('$','').replace('.','').strip() if valor_emp else '0')
                m = float(str(minimo).replace(',','.').strip() if minimo else '0')
                if m == 0:
                    return {'label': label, 'valor': str(valor_emp), 'requerido': str(minimo), 'cumple': True, 'na': True}
                cumple = v >= m if mayor_es_mejor else v <= m
                return {'label': label, 'valor': str(valor_emp), 'requerido': str(minimo), 'cumple': cumple}
            except:
                return {'label': label, 'valor': str(valor_emp), 'requerido': str(minimo), 'cumple': None}

        indicadores.append(chk('Liquidez', emp.get('liquidez'), req_fin.get('liquidez_minima')))
        indicadores.append(chk('Endeudamiento', emp.get('endeudamiento'), req_fin.get('endeudamiento_maximo'), False))
        indicadores.append(chk('Rentabilidad', emp.get('rentabilidad'), req_fin.get('rentabilidad_minima')))

        # Alertas de documentos
        docs = db.execute('SELECT * FROM documentos WHERE empresa_id=?', (emp_data['id'],)).fetchall()
        for d in docs:
            if d['fecha_vencimiento']:
                dias = (datetime.strptime(d['fecha_vencimiento'], '%Y-%m-%d') - datetime.now()).days
                if dias < 0:
                    resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': d['tipo'], 'msg': f'VENCIDO hace {abs(dias)} días', 'nivel': 'rojo'})
                elif dias <= 30:
                    resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': d['tipo'], 'msg': f'Vence en {dias} días', 'nivel': 'amarillo'})

        no_cumple = [i for i in indicadores if i.get('cumple') == False]
        if no_cumple:
            resultado['cumple_financiero'] = False

        resultado['empresas'].append({
            'id': emp_data['id'],
            'razon_social': emp['razon_social'],
            'pct': emp_data.get('pct', 100),
            'indicadores': indicadores,
            'tiene_discapacidad': emp['tiene_discapacidad'],
            'tiene_mujeres': emp['tiene_mujeres'],
            'es_mipyme': emp['es_mipyme']
        })

    # Determinar formatos a generar
    fmt_puntaje = pliego.get('formatos_puntaje', {})
    formatos = ['F1']
    if len(empresas_ids) > 1:
        formatos.append('F2')
    formatos.append('F6')
    if fmt_puntaje.get('F7A', {}).get('aplica'): formatos.append('F7A')
    if fmt_puntaje.get('F7B', {}).get('aplica'): formatos.append('F7B')
    if fmt_puntaje.get('F7C', {}).get('aplica'): formatos.append('F7C')
    if fmt_puntaje.get('F8', {}).get('aplica'): formatos.append('F8')
    if fmt_puntaje.get('F9A', {}).get('aplica'): formatos.append('F9A')
    elif fmt_puntaje.get('F9B', {}).get('aplica'): formatos.append('F9B')
    if fmt_puntaje.get('F12', {}).get('aplica'): formatos.append('F12')
    if fmt_puntaje.get('F13', {}).get('aplica'): formatos.append('F13')
    if fmt_puntaje.get('F14', {}).get('aplica'): formatos.append('F14')
    formatos.extend(['ANX4', 'F10', 'F11'])

    resultado['formatos_a_generar'] = formatos
    db.close()
    return jsonify(resultado)

# ─── GENERAR FORMATOS WORD ────────────────────────────────────────────────────
@app.route('/api/generar-formatos', methods=['POST'])
def generar_formatos():
    data = request.json
    pliego = data.get('pliego', {})
    empresas_data = data.get('empresas', [])
    tipo_proponente = data.get('tipo_proponente', 'individual')
    nombre_plural = data.get('nombre_plural', '')
    obj_plural = data.get('obj_plural', '')
    dur_plural = data.get('dur_plural', '')
    rep_plural = data.get('rep_plural', '')
    cc_rep_plural = data.get('cc_rep_plural', '')
    rep_suplente = data.get('rep_suplente', '')
    cc_suplente = data.get('cc_suplente', '')
    experiencia_seleccionada = data.get('experiencia_ids', [])
    formatos_a_generar = data.get('formatos', [])

    db = get_db()
    empresas = []
    for ed in empresas_data:
        emp = db.execute('SELECT * FROM empresas WHERE id=?', (ed['id'],)).fetchone()
        if emp:
            e = dict(emp)
            e['pct'] = ed.get('pct', 100)
            e['experiencia_sel'] = []
            for exp_id in experiencia_seleccionada:
                exp = db.execute('SELECT * FROM experiencia WHERE id=? AND empresa_id=?', (exp_id, ed['id'])).fetchone()
                if exp:
                    e['experiencia_sel'].append(dict(exp))
            empresas.append(e)
    db.close()

    if not empresas:
        return jsonify({'error': 'No se encontraron empresas'}), 400

    e0 = empresas[0]
    proponente_str = e0['razon_social'] if tipo_proponente == 'individual' else f"{'CONSORCIO' if tipo_proponente == 'consorcio' else 'UNIÓN TEMPORAL'} {nombre_plural}"
    rep_str = e0['rep_legal'] if tipo_proponente == 'individual' else rep_plural
    cc_str = e0['cc_rep_legal'] if tipo_proponente == 'individual' else cc_rep_plural

    # Generar cada formato con Claude y crear Word
    archivos_generados = {}

    def gen_texto(fmt_id):
        base = f"""
Entidad: {pliego.get('entidad','')}
Ciudad: {pliego.get('ciudad_entidad','')}
Dirección entidad: {pliego.get('direccion_entidad','')}
Proceso No.: {pliego.get('numero_proceso','')}
Objeto: {pliego.get('objeto','')}
Presupuesto oficial: {pliego.get('presupuesto_oficial','')}
Plazo: {pliego.get('plazo_ejecucion','')}
Fecha cierre: {pliego.get('fecha_cierre','')}
Proponente: {proponente_str}
Representante legal: {rep_str}
CC representante: {cc_str}
"""
        prompts_fmt = {
            'F1': f"""Redacta el FORMATO 1 - CARTA DE PRESENTACIÓN DE LA PROPUESTA completo según la Resolución 465/2024 Colombia Compra Eficiente (pliego tipo infraestructura de transporte versión 4).
{base}
Tipo de proponente: {'Empresa individual' if tipo_proponente=='individual' else tipo_proponente.title()}
{'Integrantes del ' + tipo_proponente + ': ' + ', '.join([f"{e['razon_social']} ({e['pct']}%)" for e in empresas]) if tipo_proponente != 'individual' else f"NIT: {e0['nit']} | Dirección: {e0['direccion']} | Email: {e0['email']}"}
Redacta la carta con TODOS los numerales del 1 al final, bajo la gravedad del juramento, con todos los compromisos del formato oficial. Incluye ciudad y fecha, datos del destinatario y espacio para firma al final.""",

            'F2': f"""Redacta el FORMATO 2{'A — DOCUMENTO DE CONFORMACIÓN DE CONSORCIO' if tipo_proponente=='consorcio' else 'B — DOCUMENTO DE CONFORMACIÓN DE UNIÓN TEMPORAL'} completo según la Resolución 465/2024.
{base}
Nombre del {tipo_proponente}: {nombre_plural}
Objeto del {tipo_proponente}: {obj_plural}
Duración: {dur_plural}
Representante: {rep_plural} | CC: {cc_rep_plural}
Representante suplente: {rep_suplente} | CC: {cc_suplente}
Integrantes:
{chr(10).join([f"- {e['razon_social']} | NIT: {e['nit']} | Rep. legal: {e['rep_legal']} | CC: {e['cc_rep_legal']} | Participación: {e['pct']}%" for e in empresas])}
Redacta el documento completo con tabla de integrantes, todos los numerales y espacios de firma para cada representante legal. Responsabilidad {'solidaria' if tipo_proponente=='consorcio' else 'proporcional al porcentaje de participación'}.""",

            'F6': f"""Redacta el FORMATO 6 — PAGOS DE SEGURIDAD SOCIAL Y APORTES LEGALES completo según Res. 465/2024 y Art. 50 Ley 789/2002.
{base}
{chr(10).join([f'''Empresa: {e['razon_social']} | NIT: {e['nit']}
Rep. Legal: {e['rep_legal']} | CC: {e['cc_rep_legal']}
Contador: {e['contador_nombre']} | CC: {e['contador_cc']} | T.P.: {e['contador_tp']} | Cámara: {e['contador_camara']}
Revisor fiscal: {e['revisor_nombre'] or 'No aplica'} | CC: {e['revisor_cc']} | T.P.: {e['revisor_tp']}
Exonerada parafiscales: {'Sí' if e['exonerada_parafiscales'] else 'No'}''' for e in empresas])}
Redacta la declaración juramentada completa certificando pago de aportes de salud, pensiones, riesgos, ICBF, SENA, cajas de compensación y FIC durante los últimos 6 meses. Incluye texto del representante legal Y del revisor fiscal (si aplica). Espacio para firmas al final.""",

            'F7A': f"""Redacta el FORMATO 7A — PROGRAMA DE GERENCIA DE PROYECTOS completo según Res. 465/2024.
{base}
Dirección: {e0['direccion']} | Email: {e0['email']}
Redacta el compromiso juramentado completo de implementar programa de gerencia con profesional en ingeniería o arquitectura. Incluye todos los campos de firma.""",

            'F7B': f"""Redacta el FORMATO 7B — DISPONIBILIDAD Y CONDICIONES FUNCIONALES DE LA MAQUINARIA DE OBRA completo según Res. 465/2024.
{base}
Redacta el compromiso de disponibilidad de maquinaria requerida para el proyecto. Incluye tabla de equipos y firmas.""",

            'F7C': f"""Redacta el FORMATO 7C — PLAN DE CALIDAD completo según Res. 465/2024.
{base}
Redacta el compromiso de implementar plan de calidad según normas técnicas. Incluye firmas.""",

            'F8': f"""Redacta el FORMATO 8 — VINCULACIÓN DE PERSONAS EN CONDICIÓN DE DISCAPACIDAD completo según Res. 465/2024.
{base}
Redacta la certificación con tabla de número total de trabajadores y personas con discapacidad, referencia al certificado del Ministerio de Trabajo. Incluye espacio para firma.""",

            'F9A': f"""Redacta el FORMATO 9A — PUNTAJE DE INDUSTRIA NACIONAL (servicios nacionales) completo según Res. 465/2024.
{base}
Bienes relevantes que exige el pliego: {pliego.get('formatos_puntaje',{}).get('F9A',{}).get('bienes_relevantes',[])}
Redacta el ofrecimiento de apoyo a la industria nacional con tabla de bienes nacionales relevantes y compromiso de incorporarlos. Incluye firmas.""",

            'F9B': f"""Redacta el FORMATO 9B — INCORPORACIÓN DE COMPONENTE NACIONAL EN SERVICIOS EXTRANJEROS completo según Res. 465/2024.
{base}
Redacta el formato para proponentes extranjeros que incorporan componente nacional. Incluye tabla y firmas.""",

            'F12': f"""Redacta el FORMATO 12 — ACREDITACIÓN DE EMPRENDIMIENTOS Y EMPRESAS DE MUJERES completo según Res. 465/2024.
{base}
Redacta la acreditación de empresa de mujeres según normativa vigente. Incluye declaración y firmas.""",

            'F13': f"""Redacta el FORMATO 13 — ACREDITACIÓN MIPYME completo según Res. 465/2024 y Ley 590/2000.
{base}
Redacta la acreditación como micro, pequeña o mediana empresa para obtener puntaje adicional. Incluye declaración y firmas.""",

            'F14': f"""Redacta el FORMATO 14 — FACTOR DE CALIDAD CRITERIOS ADICIONALES AMBIENTALES Y SOCIALES completo según Res. 465/2024 y Decreto 142/2023.
{base}
Redacta el compromiso de implementar programa ambiental y social según numeral 4.2.4 del documento base. Incluye firmas.""",

            'ANX4': f"""Redacta el ANEXO 4 — PACTO DE TRANSPARENCIA completo según Res. 465/2024.
{base}
Redacta todos los compromisos i) al xv) del pacto oficial: cumplir la ley, buena fe, no falsificación, libre competencia, no colusión, no sobornos, preguntas por escrito, lealtad, compostura en audiencias y demás. Espacio para firma.""",

            'F10': f"""Redacta el FORMATO 10 — FACTORES DE DESEMPATE completo según Res. 465/2024.
{base}
Indica factores de desempate aplicables: Mipyme={'Sí' if any(e['es_mipyme'] for e in empresas) else 'No'}, Empresa mujeres={'Sí' if any(e['tiene_mujeres'] for e in empresas) else 'No'}, Discapacidad={'Sí' if any(e['tiene_discapacidad'] for e in empresas) else 'No'}. Incluye declaración y firmas.""",

            'F11': f"""Redacta el FORMATO 11 — AUTORIZACIÓN PARA EL TRATAMIENTO DE DATOS PERSONALES completo según Res. 465/2024 y Ley 1581/2012.
{base}
Titulares: {', '.join([e['rep_legal'] or e['razon_social'] for e in empresas])}
Redacta la autorización completa para tratamiento de datos. Incluye firmas de cada titular.""",
        }

        return prompts_fmt.get(fmt_id, '')

    def crear_word(fmt_id, texto):
        doc = Document()
        # Estilo
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)

        # Encabezado
        header = doc.sections[0].header
        hp = header.paragraphs[0]
        hp.text = f"Colombia Compra Eficiente — Res. 465/2024 V4 | {pliego.get('numero_proceso','')} | {fmt_id}"
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        hp.runs[0].font.size = Pt(9)
        hp.runs[0].font.color.rgb = RGBColor(0x18, 0x5F, 0xA5)

        # Número de proceso
        p = doc.add_paragraph()
        p.add_run(f"[{pliego.get('numero_proceso','')}]").bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        # Título formato
        nombres = {
            'F1':'FORMATO 1\nCARTA DE PRESENTACIÓN DE LA PROPUESTA',
            'F2':f"FORMATO {'2A' if tipo_proponente=='consorcio' else '2B'}\nCONFORMACIÓN DE PROPONENTE PLURAL",
            'F6':'FORMATO 6\nPAGOS DE SEGURIDAD SOCIAL Y APORTES LEGALES',
            'F7A':'FORMATO 7A\nPROGRAMA DE GERENCIA DE PROYECTOS',
            'F7B':'FORMATO 7B\nDISPONIBILIDAD Y CONDICIONES DE MAQUINARIA',
            'F7C':'FORMATO 7C\nPLAN DE CALIDAD',
            'F8':'FORMATO 8\nVINCULACIÓN PERSONAS EN CONDICIÓN DE DISCAPACIDAD',
            'F9A':'FORMATO 9A\nPUNTAJE DE INDUSTRIA NACIONAL',
            'F9B':'FORMATO 9B\nINCORPORACIÓN COMPONENTE NACIONAL',
            'F12':'FORMATO 12\nACREDITACIÓN EMPRESAS DE MUJERES',
            'F13':'FORMATO 13\nACREDITACIÓN MIPYME',
            'F14':'FORMATO 14\nFACTOR DE CALIDAD AMBIENTAL Y SOCIAL',
            'ANX4':'ANEXO 4\nPACTO DE TRANSPARENCIA',
            'F10':'FORMATO 10\nFACTORES DE DESEMPATE',
            'F11':'FORMATO 11\nAUTORIZACIÓN TRATAMIENTO DE DATOS PERSONALES',
        }
        titulo = doc.add_heading(nombres.get(fmt_id, fmt_id), level=1)
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()

        # Contenido generado por IA
        for linea in texto.split('\n'):
            p = doc.add_paragraph(linea)
            p.paragraph_format.space_after = Pt(2)

        # Pie de página
        footer = doc.sections[0].footer
        fp = footer.paragraphs[0]
        fp.text = f"Generado por sistema propuestas-vias | {datetime.now().strftime('%d/%m/%Y')}"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.runs[0].font.size = Pt(8)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    # Generar todos los formatos
    archivos = {}
    for fmt_id in formatos_a_generar:
        prompt = gen_texto(fmt_id)
        if not prompt:
            continue
        try:
            texto = call_claude(prompt)
            word_buf = crear_word(fmt_id, texto)
            archivos[fmt_id] = word_buf.getvalue()
        except Exception as e:
            pass

    # Empacar en ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        proceso = pliego.get('numero_proceso', 'proceso').replace('/', '-')
        for fmt_id, contenido in archivos.items():
            nombres_archivo = {
                'F1': 'Formato1_CartaPresentacion',
                'F2': 'Formato2_Consorcio_UT',
                'F6': 'Formato6_SeguridadSocial',
                'F7A': 'Formato7A_GerenciaProyectos',
                'F7B': 'Formato7B_Maquinaria',
                'F7C': 'Formato7C_PlanCalidad',
                'F8': 'Formato8_Discapacidad',
                'F9A': 'Formato9A_IndustriaNacional',
                'F9B': 'Formato9B_ComponenteNacional',
                'F12': 'Formato12_EmpresasMujeres',
                'F13': 'Formato13_Mipyme',
                'F14': 'Formato14_AmbientalSocial',
                'ANX4': 'Anexo4_PactoTransparencia',
                'F10': 'Formato10_Desempate',
                'F11': 'Formato11_DatosPersonales',
            }
            nombre = f"{proceso}_{nombres_archivo.get(fmt_id, fmt_id)}.docx"
            zf.writestr(nombre, contenido)

    zip_buf.seek(0)
    return send_file(
        zip_buf,
        mimetype='application/zip',
        as_attachment=True,
        download_name=f"Propuesta_{pliego.get('numero_proceso','proceso').replace('/','_')}.zip"
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
