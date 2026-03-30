from flask import Flask, request, jsonify, render_template, send_file
import anthropic, os, json, sqlite3, base64, io, zipfile, requests
from datetime import datetime, timedelta
import cloudinary, cloudinary.uploader, cloudinary.api
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

cloudinary.config(
    cloud_name=os.environ.get("CLOUDINARY_CLOUD_NAME"),
    api_key=os.environ.get("CLOUDINARY_API_KEY"),
    api_secret=os.environ.get("CLOUDINARY_API_SECRET"),
    secure=True
)

def get_db():
    db = sqlite3.connect('propuestas.db')
    db.row_factory = sqlite3.Row
    return db

def init_db():
    db = get_db()
    db.executescript('''
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            razon_social TEXT, nit TEXT, tipo TEXT, direccion TEXT, ciudad TEXT,
            telefono TEXT, email TEXT,
            rep_legal TEXT, cc_rep_legal TEXT, rep_es_ingeniero INTEGER DEFAULT 0, rep_matricula TEXT,
            camara_comercio TEXT, fecha_vencimiento_camara TEXT, fecha_vencimiento_rup TEXT,
            contador_nombre TEXT, contador_cc TEXT, contador_tp TEXT,
            revisor_nombre TEXT, revisor_cc TEXT, revisor_tp TEXT,
            cont_ind_nombre TEXT, cont_ind_cc TEXT, cont_ind_tp TEXT,
            capital_trabajo TEXT, patrimonio TEXT, liquidez TEXT,
            endeudamiento TEXT, rentabilidad TEXT, rentabilidad_activo TEXT,
            tiene_discapacidad INTEGER DEFAULT 0, tiene_mujeres INTEGER DEFAULT 0,
            es_mipyme INTEGER DEFAULT 0, exonerada_parafiscales INTEGER DEFAULT 0,
            fecha_constitucion TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS experiencia (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER, entidad TEXT, objeto TEXT, valor TEXT,
            fecha_inicio TEXT, fecha_fin TEXT, plazo TEXT,
            consecutivo_rup TEXT, acta TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id)
        );
        CREATE TABLE IF NOT EXISTS documentos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            empresa_id INTEGER, experiencia_id INTEGER,
            tipo TEXT, nombre_archivo TEXT, url_cloudinary TEXT,
            public_id_cloudinary TEXT, fecha_subida TEXT, fecha_vencimiento TEXT,
            FOREIGN KEY (empresa_id) REFERENCES empresas(id),
            FOREIGN KEY (experiencia_id) REFERENCES experiencia(id)
        );
    ''')
    db.commit()
    db.close()

init_db()

def call_claude(prompt, pdf_base64=None):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    content = []
    if pdf_base64:
        content.append({"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_base64}})
    content.append({"type": "text", "text": prompt})
    r = client.messages.create(model="claude-sonnet-4-20250514", max_tokens=4000, messages=[{"role": "user", "content": content}])
    return r.content[0].text

def dias_hasta(fecha_str):
    try:
        return (datetime.strptime(fecha_str, '%Y-%m-%d') - datetime.now()).days
    except:
        return None

def calcular_alertas(docs):
    alertas = []
    for d in docs:
        if d['fecha_vencimiento']:
            dias = dias_hasta(d['fecha_vencimiento'])
            if dias is not None:
                if dias < 0:
                    alertas.append({'tipo': d['tipo'], 'msg': f"VENCIDO hace {abs(dias)} días", 'nivel': 'rojo'})
                elif dias <= 30:
                    alertas.append({'tipo': d['tipo'], 'msg': f"Vence en {dias} días", 'nivel': 'amarillo'})
    return alertas

@app.route('/')
def index():
    return render_template('index.html')

# ─── EMPRESAS ─────────────────────────────────────────────────────────────────
@app.route('/api/empresas', methods=['GET'])
def get_empresas():
    db = get_db()
    empresas = db.execute('SELECT * FROM empresas ORDER BY razon_social').fetchall()
    result = []
    for e in empresas:
        emp = dict(e)
        docs = [dict(d) for d in db.execute('SELECT * FROM documentos WHERE empresa_id=? AND experiencia_id IS NULL', (e['id'],)).fetchall()]
        emp['documentos'] = docs
        emp['alertas'] = calcular_alertas(docs)
        # Alertas RUP y cámara
        for campo, label in [('fecha_vencimiento_camara','Cámara de comercio'), ('fecha_vencimiento_rup','RUP')]:
            if emp.get(campo):
                dias = dias_hasta(emp[campo])
                if dias is not None:
                    if dias < 0:
                        emp['alertas'].append({'tipo': label, 'msg': f"VENCIDO hace {abs(dias)} días", 'nivel': 'rojo'})
                    elif dias <= 30:
                        emp['alertas'].append({'tipo': label, 'msg': f"Vence en {dias} días", 'nivel': 'amarillo'})
        exp = db.execute('SELECT * FROM experiencia WHERE empresa_id=?', (e['id'],)).fetchall()
        exp_list = []
        for x in exp:
            xd = dict(x)
            xd['documentos'] = [dict(d) for d in db.execute('SELECT * FROM documentos WHERE experiencia_id=?', (x['id'],)).fetchall()]
            exp_list.append(xd)
        emp['experiencia'] = exp_list
        result.append(emp)
    db.close()
    return jsonify(result)

@app.route('/api/empresas', methods=['POST'])
def create_empresa():
    d = request.json
    db = get_db()
    cur = db.execute('''INSERT INTO empresas (razon_social,nit,tipo,direccion,ciudad,telefono,email,
        rep_legal,cc_rep_legal,rep_es_ingeniero,rep_matricula,camara_comercio,fecha_vencimiento_camara,
        fecha_vencimiento_rup,contador_nombre,contador_cc,contador_tp,revisor_nombre,revisor_cc,revisor_tp,
        cont_ind_nombre,cont_ind_cc,cont_ind_tp,capital_trabajo,patrimonio,liquidez,endeudamiento,
        rentabilidad,rentabilidad_activo,tiene_discapacidad,tiene_mujeres,es_mipyme,exonerada_parafiscales,fecha_constitucion)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        (d.get('razon_social',''), d.get('nit',''), d.get('tipo',''), d.get('direccion',''),
         d.get('ciudad',''), d.get('telefono',''), d.get('email',''), d.get('rep_legal',''),
         d.get('cc_rep_legal',''), d.get('rep_es_ingeniero',0), d.get('rep_matricula',''),
         d.get('camara_comercio',''), d.get('fecha_vencimiento_camara',''), d.get('fecha_vencimiento_rup',''),
         d.get('contador_nombre',''), d.get('contador_cc',''), d.get('contador_tp',''),
         d.get('revisor_nombre',''), d.get('revisor_cc',''), d.get('revisor_tp',''),
         d.get('cont_ind_nombre',''), d.get('cont_ind_cc',''), d.get('cont_ind_tp',''),
         d.get('capital_trabajo',''), d.get('patrimonio',''), d.get('liquidez',''),
         d.get('endeudamiento',''), d.get('rentabilidad',''), d.get('rentabilidad_activo',''),
         d.get('tiene_discapacidad',0), d.get('tiene_mujeres',0), d.get('es_mipyme',0),
         d.get('exonerada_parafiscales',0), d.get('fecha_constitucion','')))
    eid = cur.lastrowid
    db.commit()
    db.close()
    return jsonify({'id': eid, 'ok': True})

@app.route('/api/empresas/<int:eid>', methods=['PUT'])
def update_empresa(eid):
    d = request.json
    db = get_db()
    db.execute('''UPDATE empresas SET razon_social=?,nit=?,tipo=?,direccion=?,ciudad=?,telefono=?,email=?,
        rep_legal=?,cc_rep_legal=?,rep_es_ingeniero=?,rep_matricula=?,camara_comercio=?,fecha_vencimiento_camara=?,
        fecha_vencimiento_rup=?,contador_nombre=?,contador_cc=?,contador_tp=?,revisor_nombre=?,revisor_cc=?,revisor_tp=?,
        cont_ind_nombre=?,cont_ind_cc=?,cont_ind_tp=?,capital_trabajo=?,patrimonio=?,liquidez=?,endeudamiento=?,
        rentabilidad=?,rentabilidad_activo=?,tiene_discapacidad=?,tiene_mujeres=?,es_mipyme=?,exonerada_parafiscales=?,
        fecha_constitucion=? WHERE id=?''',
        (d.get('razon_social',''), d.get('nit',''), d.get('tipo',''), d.get('direccion',''),
         d.get('ciudad',''), d.get('telefono',''), d.get('email',''), d.get('rep_legal',''),
         d.get('cc_rep_legal',''), d.get('rep_es_ingeniero',0), d.get('rep_matricula',''),
         d.get('camara_comercio',''), d.get('fecha_vencimiento_camara',''), d.get('fecha_vencimiento_rup',''),
         d.get('contador_nombre',''), d.get('contador_cc',''), d.get('contador_tp',''),
         d.get('revisor_nombre',''), d.get('revisor_cc',''), d.get('revisor_tp',''),
         d.get('cont_ind_nombre',''), d.get('cont_ind_cc',''), d.get('cont_ind_tp',''),
         d.get('capital_trabajo',''), d.get('patrimonio',''), d.get('liquidez',''),
         d.get('endeudamiento',''), d.get('rentabilidad',''), d.get('rentabilidad_activo',''),
         d.get('tiene_discapacidad',0), d.get('tiene_mujeres',0), d.get('es_mipyme',0),
         d.get('exonerada_parafiscales',0), d.get('fecha_constitucion',''), eid))
    db.commit()
    db.close()
    return jsonify({'ok': True})

@app.route('/api/empresas/<int:eid>', methods=['DELETE'])
def delete_empresa(eid):
    db = get_db()
    # Borrar docs de cloudinary
    docs = db.execute('SELECT public_id_cloudinary FROM documentos WHERE empresa_id=?', (eid,)).fetchall()
    for d in docs:
        if d['public_id_cloudinary']:
            try: cloudinary.uploader.destroy(d['public_id_cloudinary'], resource_type='raw')
            except: pass
    db.execute('DELETE FROM documentos WHERE empresa_id=?', (eid,))
    db.execute('DELETE FROM experiencia WHERE empresa_id=?', (eid,))
    db.execute('DELETE FROM empresas WHERE id=?', (eid,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

# ─── EXPERIENCIA ──────────────────────────────────────────────────────────────
@app.route('/api/empresas/<int:eid>/experiencia', methods=['POST'])
def add_experiencia(eid):
    d = request.json
    db = get_db()
    cur = db.execute('INSERT INTO experiencia (empresa_id,entidad,objeto,valor,fecha_inicio,fecha_fin,plazo,consecutivo_rup,acta) VALUES (?,?,?,?,?,?,?,?,?)',
        (eid, d.get('entidad',''), d.get('objeto',''), d.get('valor',''), d.get('fecha_inicio',''),
         d.get('fecha_fin',''), d.get('plazo',''), d.get('consecutivo_rup',''), d.get('acta','')))
    db.commit()
    xid = cur.lastrowid
    db.close()
    return jsonify({'id': xid, 'ok': True})

@app.route('/api/experiencia/<int:xid>', methods=['DELETE'])
def delete_experiencia(xid):
    db = get_db()
    docs = db.execute('SELECT public_id_cloudinary FROM documentos WHERE experiencia_id=?', (xid,)).fetchall()
    for d in docs:
        if d['public_id_cloudinary']:
            try: cloudinary.uploader.destroy(d['public_id_cloudinary'], resource_type='raw')
            except: pass
    db.execute('DELETE FROM documentos WHERE experiencia_id=?', (xid,))
    db.execute('DELETE FROM experiencia WHERE id=?', (xid,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

# ─── SUBIR DOCUMENTOS ─────────────────────────────────────────────────────────
@app.route('/api/subir-documento', methods=['POST'])
def subir_documento():
    empresa_id = request.form.get('empresa_id')
    experiencia_id = request.form.get('experiencia_id')
    tipo = request.form.get('tipo', 'soporte')
    archivo = request.files.get('archivo')
    extraer = request.form.get('extraer', 'false') == 'true'

    if not archivo:
        return jsonify({'error': 'No se recibió archivo'}), 400

    # Subir a Cloudinary
    try:
        folder = f"propuestas/empresa_{empresa_id}"
        resultado = cloudinary.uploader.upload(archivo, folder=folder, resource_type='raw',
            public_id=f"{tipo}_{datetime.now().strftime('%Y%m%d%H%M%S')}", use_filename=True)
        url = resultado['secure_url']
        public_id = resultado['public_id']
    except Exception as e:
        return jsonify({'error': f'Error subiendo a Cloudinary: {str(e)}'}), 500

    # Fechas de vencimiento según tipo (30 días para cámara y RUP)
    fecha_venc = ''
    datos_extraidos = {}

    if extraer and empresa_id:
        # Leer PDF para extracción
        archivo.seek(0)
        pdf_b64 = base64.standard_b64encode(archivo.read()).decode('utf-8')

        prompts = {
            'camara': '''Analiza este certificado del Registro Único de Proponentes (RUP) o Cámara de Comercio colombiano y extrae en JSON:
{
  "razon_social": "",
  "nit": "",
  "tipo_empresa": "",
  "direccion": "",
  "ciudad": "",
  "telefono": "",
  "email": "",
  "rep_legal": "",
  "cc_rep_legal": "",
  "revisor_nombre": "",
  "revisor_cc": "",
  "revisor_tp": "",
  "fecha_expedicion": "YYYY-MM-DD",
  "es_mipyme": false,
  "liquidez": "",
  "endeudamiento": "",
  "rentabilidad": "",
  "rentabilidad_activo": "",
  "capital_trabajo": "",
  "patrimonio": ""
}
Solo responde el JSON exacto sin texto adicional. La fecha_expedicion es la fecha en que fue expedido el certificado.''',

            'rup': '''Analiza este Registro Único de Proponentes (RUP) colombiano de la Cámara de Comercio y extrae en JSON:
{
  "razon_social": "",
  "nit": "",
  "tipo_empresa": "",
  "direccion": "",
  "ciudad": "",
  "telefono": "",
  "email": "",
  "rep_legal": "",
  "cc_rep_legal": "",
  "revisor_nombre": "",
  "revisor_cc": "",
  "revisor_tp": "",
  "fecha_expedicion": "YYYY-MM-DD",
  "es_mipyme": false,
  "liquidez": "",
  "endeudamiento": "",
  "rentabilidad": "",
  "rentabilidad_activo": "",
  "capital_trabajo": "",
  "patrimonio": ""
}
Busca: INDICE DE LIQUIDEZ, INDICE DE ENDEUDAMIENTO, RENTABILIDAD DEL PATRIMONIO, RENTABILIDAD DEL ACTIVO, ACTIVO CORRIENTE - PASIVO CORRIENTE = capital de trabajo, PATRIMONIO.
La fecha_expedicion es la fecha en la parte superior del documento.
Solo responde el JSON exacto sin texto adicional.''',

            'cedula': '''Analiza esta cédula de ciudadanía colombiana y extrae en JSON:
{"nombre_completo": "", "numero_cedula": "", "fecha_expedicion": "YYYY-MM-DD"}
Solo responde el JSON.''',

            'tarjeta_profesional': '''Analiza esta tarjeta profesional colombiana y extrae en JSON:
{"nombre_completo": "", "numero_tp": "", "profesion": "", "fecha_expedicion": "YYYY-MM-DD", "fecha_vencimiento": "YYYY-MM-DD"}
Solo responde el JSON.''',

            'cert_discapacidad': '''Analiza este certificado del Ministerio de Trabajo colombiano y extrae en JSON:
{"empresa": "", "total_trabajadores": "", "trabajadores_discapacidad": "", "fecha_expedicion": "YYYY-MM-DD", "fecha_vencimiento": "YYYY-MM-DD"}
Solo responde el JSON.''',

            'acta_accionaria': '''Analiza este acta de composición accionaria y extrae en JSON:
{"empresa": "", "porcentaje_mujeres": "", "fecha_acta": "YYYY-MM-DD", "socias": []}
Solo responde el JSON.'''
        }

        prompt = prompts.get(tipo, '')
        if prompt:
            try:
                resp = call_claude(prompt, pdf_b64)
                resp = resp.strip()
                if '```' in resp:
                    resp = resp.split('```')[1]
                    if resp.startswith('json'): resp = resp[4:]
                datos_extraidos = json.loads(resp.strip())

                # Calcular vencimiento 30 días para cámara y RUP
                if tipo in ['camara', 'rup'] and datos_extraidos.get('fecha_expedicion'):
                    try:
                        fexp = datetime.strptime(datos_extraidos['fecha_expedicion'], '%Y-%m-%d')
                        fecha_venc = (fexp + timedelta(days=30)).strftime('%Y-%m-%d')
                    except: pass

                # Actualizar empresa con datos extraídos
                if empresa_id and datos_extraidos:
                    db = get_db()
                    if tipo in ['camara', 'rup']:
                        updates = []
                        vals = []
                        campos = {
                            'razon_social': 'razon_social', 'nit': 'nit', 'tipo_empresa': 'tipo',
                            'direccion': 'direccion', 'ciudad': 'ciudad', 'telefono': 'telefono',
                            'email': 'email', 'rep_legal': 'rep_legal', 'cc_rep_legal': 'cc_rep_legal',
                            'revisor_nombre': 'revisor_nombre', 'revisor_cc': 'revisor_cc',
                            'revisor_tp': 'revisor_tp', 'liquidez': 'liquidez',
                            'endeudamiento': 'endeudamiento', 'rentabilidad': 'rentabilidad',
                            'rentabilidad_activo': 'rentabilidad_activo', 'capital_trabajo': 'capital_trabajo',
                            'patrimonio': 'patrimonio'
                        }
                        for k, col in campos.items():
                            if datos_extraidos.get(k):
                                updates.append(f'{col}=?')
                                vals.append(datos_extraidos[k])
                        if tipo == 'camara' and fecha_venc:
                            updates.append('fecha_vencimiento_camara=?')
                            vals.append(fecha_venc)
                        if tipo == 'rup' and fecha_venc:
                            updates.append('fecha_vencimiento_rup=?')
                            vals.append(fecha_venc)
                        if datos_extraidos.get('es_mipyme'):
                            updates.append('es_mipyme=?')
                            vals.append(1)
                        if updates:
                            vals.append(empresa_id)
                            db.execute(f"UPDATE empresas SET {','.join(updates)} WHERE id=?", vals)
                            db.commit()
                    db.close()
            except Exception as e:
                datos_extraidos = {'error': str(e)}

    # Guardar en DB
    db = get_db()
    db.execute('''INSERT INTO documentos (empresa_id, experiencia_id, tipo, nombre_archivo, url_cloudinary, public_id_cloudinary, fecha_subida, fecha_vencimiento)
        VALUES (?,?,?,?,?,?,?,?)''',
        (empresa_id, experiencia_id, tipo, archivo.filename, url, public_id,
         datetime.now().strftime('%Y-%m-%d'), fecha_venc))
    db.commit()
    db.close()

    return jsonify({'ok': True, 'url': url, 'datos_extraidos': datos_extraidos, 'fecha_vencimiento': fecha_venc})

@app.route('/api/documentos/<int:doc_id>', methods=['DELETE'])
def delete_documento(doc_id):
    db = get_db()
    doc = db.execute('SELECT * FROM documentos WHERE id=?', (doc_id,)).fetchone()
    if doc and doc['public_id_cloudinary']:
        try: cloudinary.uploader.destroy(doc['public_id_cloudinary'], resource_type='raw')
        except: pass
    db.execute('DELETE FROM documentos WHERE id=?', (doc_id,))
    db.commit()
    db.close()
    return jsonify({'ok': True})

# ─── ANALIZAR PLIEGO ─────────────────────────────────────────────────────────
@app.route('/api/analizar-pliego', methods=['POST'])
def analizar_pliego():
    archivo = request.files.get('pliego')
    if not archivo:
        return jsonify({'error': 'No se recibió el pliego'}), 400
    pdf_b64 = base64.standard_b64encode(archivo.read()).decode('utf-8')

    prompt = '''Analiza este pliego de condiciones colombiano (pliego tipo infraestructura de transporte Resolución 465/2024) y extrae en JSON exacto:
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
    "financieros": {
      "liquidez_minima": "",
      "endeudamiento_maximo": "",
      "rentabilidad_minima": "",
      "capital_trabajo_minimo": ""
    },
    "tecnicos": {
      "experiencia_minima_valor": "",
      "capacidad_residual_minima": ""
    }
  },
  "formatos_puntaje": {
    "F7A": {"aplica": false, "puntaje": ""},
    "F7B": {"aplica": false, "puntaje": ""},
    "F7C": {"aplica": false, "puntaje": ""},
    "F8": {"aplica": false, "puntaje": ""},
    "F9A": {"aplica": false, "puntaje": "", "tiene_bienes_relevantes": false, "bienes_relevantes": [], "solo_nacionales": true},
    "F9B": {"aplica": false, "puntaje": ""},
    "F12": {"aplica": false, "puntaje": ""},
    "F13": {"aplica": false, "puntaje": ""},
    "F14": {"aplica": false, "puntaje": ""}
  },
  "formula_consorcio": "",
  "total_puntaje": "100"
}
Solo responde el JSON sin texto adicional.'''

    try:
        resp = call_claude(prompt, pdf_b64)
        resp = resp.strip()
        if '```' in resp:
            resp = resp.split('```')[1]
            if resp.startswith('json'): resp = resp[4:]
        datos = json.loads(resp.strip())
        return jsonify({'ok': True, 'pliego': datos})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ─── VERIFICAR CUMPLIMIENTO ───────────────────────────────────────────────────
@app.route('/api/verificar-cumplimiento', methods=['POST'])
def verificar_cumplimiento():
    data = request.json
    pliego = data.get('pliego', {})
    empresas_ids = data.get('empresas', [])
    db = get_db()
    req_fin = pliego.get('requisitos_habilitantes', {}).get('financieros', {})
    resultado = {'empresas': [], 'alertas_docs': [], 'formatos_a_generar': [], 'info_f9': {}}

    for ed in empresas_ids:
        emp = db.execute('SELECT * FROM empresas WHERE id=?', (ed['id'],)).fetchone()
        if not emp: continue
        emp = dict(emp)
        pct = float(ed.get('pct', 100))

        def chk(label, val, minimo, mayor=True):
            try:
                v = float(str(val or '0').replace(',','.').replace('$','').replace('.','').strip())
                m = float(str(minimo or '0').replace(',','.').strip())
                if m == 0: return {'label': label, 'valor': str(val), 'requerido': str(minimo), 'cumple': True, 'na': True}
                return {'label': label, 'valor': str(val), 'requerido': str(minimo), 'cumple': v>=m if mayor else v<=m}
            except:
                return {'label': label, 'valor': str(val), 'requerido': str(minimo), 'cumple': None}

        indicadores = [
            chk('Liquidez', emp.get('liquidez'), req_fin.get('liquidez_minima')),
            chk('Endeudamiento', emp.get('endeudamiento'), req_fin.get('endeudamiento_maximo'), False),
            chk('Rentabilidad patrimonio', emp.get('rentabilidad'), req_fin.get('rentabilidad_minima')),
        ]

        # Alertas documentos
        docs = db.execute('SELECT * FROM documentos WHERE empresa_id=? AND experiencia_id IS NULL', (ed['id'],)).fetchall()
        for d in docs:
            if d['fecha_vencimiento']:
                dias = dias_hasta(d['fecha_vencimiento'])
                if dias is not None:
                    if dias < 0:
                        resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': d['tipo'], 'msg': f'VENCIDO hace {abs(dias)} días', 'nivel': 'rojo'})
                    elif dias <= 30:
                        resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': d['tipo'], 'msg': f'Vence en {dias} días', 'nivel': 'amarillo'})
        for campo, label in [('fecha_vencimiento_camara','Cámara'), ('fecha_vencimiento_rup','RUP')]:
            if emp.get(campo):
                dias = dias_hasta(emp[campo])
                if dias is not None:
                    if dias < 0:
                        resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': label, 'msg': f'VENCIDO hace {abs(dias)} días', 'nivel': 'rojo'})
                    elif dias <= 30:
                        resultado['alertas_docs'].append({'empresa': emp['razon_social'], 'doc': label, 'msg': f'Vence en {dias} días', 'nivel': 'amarillo'})

        resultado['empresas'].append({
            'id': ed['id'], 'razon_social': emp['razon_social'], 'pct': pct,
            'indicadores': indicadores,
            'tiene_discapacidad': emp['tiene_discapacidad'],
            'tiene_mujeres': emp['tiene_mujeres'],
            'es_mipyme': emp['es_mipyme'],
            'rep_es_ingeniero': emp['rep_es_ingeniero']
        })

    # Formatos a generar
    fmt = pliego.get('formatos_puntaje', {})
    formatos = ['F1']
    if len(empresas_ids) > 1: formatos.append('F2')
    formatos.append('F6')
    if fmt.get('F7A',{}).get('aplica'): formatos.append('F7A')
    if fmt.get('F7B',{}).get('aplica'): formatos.append('F7B')
    if fmt.get('F7C',{}).get('aplica'): formatos.append('F7C')
    if fmt.get('F8',{}).get('aplica'): formatos.append('F8')

    # Info F9 para preview editable
    f9a = fmt.get('F9A', {})
    f9b = fmt.get('F9B', {})
    if f9a.get('aplica'):
        formatos.append('F9A')
        resultado['info_f9'] = {
            'formato': 'F9A',
            'tiene_bienes_relevantes': f9a.get('tiene_bienes_relevantes', False),
            'bienes_relevantes': f9a.get('bienes_relevantes', []),
            'solo_nacionales': f9a.get('solo_nacionales', True),
            'puntaje': f9a.get('puntaje', ''),
            'descripcion': 'Promoción de servicios nacionales o con trato nacional'
        }
    elif f9b.get('aplica'):
        formatos.append('F9B')
        resultado['info_f9'] = {
            'formato': 'F9B',
            'tiene_bienes_relevantes': False,
            'bienes_relevantes': [],
            'solo_nacionales': False,
            'puntaje': f9b.get('puntaje', ''),
            'descripcion': 'Incorporación de componente nacional en servicios extranjeros'
        }

    if fmt.get('F12',{}).get('aplica'): formatos.append('F12')
    if fmt.get('F13',{}).get('aplica'): formatos.append('F13')
    if fmt.get('F14',{}).get('aplica'): formatos.append('F14')
    formatos.extend(['ANX4', 'F10', 'F11'])
    resultado['formatos_a_generar'] = formatos
    db.close()
    return jsonify(resultado)

# ─── GENERAR FORMATOS ─────────────────────────────────────────────────────────
@app.route('/api/generar-formatos', methods=['POST'])
def generar_formatos():
    data = request.json
    pliego = data.get('pliego', {})
    empresas_data = data.get('empresas', [])
    tipo_prop = data.get('tipo_proponente', 'individual')
    nombre_plural = data.get('nombre_plural', '')
    obj_plural = data.get('obj_plural', '')
    dur_plural = data.get('dur_plural', '')
    rep_plural = data.get('rep_plural', '')
    cc_rep_plural = data.get('cc_rep_plural', '')
    rep_suplente = data.get('rep_suplente', '')
    cc_suplente = data.get('cc_suplente', '')
    ing_avalador = data.get('ing_avalador', {})
    exp_ids = data.get('experiencia_ids', [])
    formatos = data.get('formatos', [])
    f9_config = data.get('f9_config', {})

    db = get_db()
    empresas = []
    for ed in empresas_data:
        emp = db.execute('SELECT * FROM empresas WHERE id=?', (ed['id'],)).fetchone()
        if emp:
            e = dict(emp)
            e['pct'] = ed.get('pct', 100)
            e['exp_sel'] = []
            for xid in exp_ids:
                x = db.execute('SELECT * FROM experiencia WHERE id=? AND empresa_id=?', (xid, ed['id'])).fetchone()
                if x:
                    xd = dict(x)
                    xd['docs'] = [dict(d) for d in db.execute('SELECT * FROM documentos WHERE experiencia_id=?', (xid,)).fetchall()]
                    e['exp_sel'].append(xd)
            e['docs'] = [dict(d) for d in db.execute('SELECT * FROM documentos WHERE empresa_id=? AND experiencia_id IS NULL', (ed['id'],)).fetchall()]
            empresas.append(e)
    db.close()

    if not empresas:
        return jsonify({'error': 'No se encontraron empresas'}), 400

    e0 = empresas[0]
    prop_str = e0['razon_social'] if tipo_prop == 'individual' else f"{'CONSORCIO' if tipo_prop=='consorcio' else 'UNIÓN TEMPORAL'} {nombre_plural}"
    rep_str = e0['rep_legal'] if tipo_prop == 'individual' else rep_plural
    cc_str = e0['cc_rep_legal'] if tipo_prop == 'individual' else cc_rep_plural

    base = f"""Entidad: {pliego.get('entidad','')}
Ciudad: {pliego.get('ciudad_entidad','')}
Dirección: {pliego.get('direccion_entidad','')}
Proceso No.: {pliego.get('numero_proceso','')}
Objeto: {pliego.get('objeto','')}
Presupuesto: {pliego.get('presupuesto_oficial','')}
Plazo: {pliego.get('plazo_ejecucion','')}
Fecha cierre: {pliego.get('fecha_cierre','')}
Proponente: {prop_str}
Representante legal: {rep_str} | CC: {cc_str}"""

    def prompt_fmt(fid):
        emp_str = '\n'.join([f"- {e['razon_social']} | NIT: {e['nit']} | Rep: {e['rep_legal']} | CC: {e['cc_rep_legal']} | {e['pct']}%" for e in empresas])
        exp_str = '\n'.join([f"{i+1}. {x['entidad']} | {x['objeto'][:80]} | {x['valor']} | Inicio: {x['fecha_inicio']} | Fin: {x['fecha_fin']} | Cons. RUP: {x['consecutivo_rup']}" for e in empresas for i,x in enumerate(e['exp_sel'])])

        prompts = {
            'F1': f"""Redacta el FORMATO 1 - CARTA DE PRESENTACIÓN DE LA PROPUESTA completo según Resolución 465/2024 Colombia Compra Eficiente versión 4.
{base}
Tipo: {'Empresa individual' if tipo_prop=='individual' else tipo_prop.title()}
{emp_str if tipo_prop!='individual' else f"NIT: {e0['nit']} | Dir: {e0['direccion']} | Email: {e0['email']}"}
Redacta con TODOS los numerales del 1 al final bajo gravedad del juramento. Ciudad y fecha al inicio. Espacio para firma al final.""",

            'F2': f"""Redacta el FORMATO {'2A — CONSORCIO' if tipo_prop=='consorcio' else '2B — UNIÓN TEMPORAL'} completo según Resolución 465/2024.
{base}
Nombre: {nombre_plural} | Objeto: {obj_plural} | Duración: {dur_plural}
Rep: {rep_plural} CC: {cc_rep_plural} | Suplente: {rep_suplente} CC: {cc_suplente}
Integrantes:
{emp_str}
Responsabilidad {'solidaria' if tipo_prop=='consorcio' else 'proporcional'}. Incluye tabla y firmas de cada rep legal.""",

            'F6': f"""Redacta el FORMATO 6 — PAGOS DE SEGURIDAD SOCIAL completo según Res. 465/2024 Art. 50 Ley 789/2002.
{base}
{chr(10).join([f"Empresa: {e['razon_social']} | NIT: {e['nit']} | Rep: {e['rep_legal']} | CC: {e['cc_rep_legal']}\\nContador: {e['contador_nombre']} | CC: {e['contador_cc']} | TP: {e['contador_tp']}\\nRevisor: {e['revisor_nombre'] or 'No aplica'} | CC: {e['revisor_cc']} | TP: {e['revisor_tp']}\\nCont.Ind: {e['cont_ind_nombre'] or 'No aplica'} | CC: {e['cont_ind_cc']} | TP: {e['cont_ind_tp']}\\nExonerada parafiscales: {'Sí' if e['exonerada_parafiscales'] else 'No'}" for e in empresas])}
Declaración juramentada de pago de aportes salud, pensiones, riesgos, ICBF, SENA, cajas compensación y FIC últimos 6 meses. Incluye texto rep legal Y revisor/contador. Firmas al final.""",

            'F7A': f"""Redacta FORMATO 7A — PROGRAMA DE GERENCIA DE PROYECTOS completo según Res. 465/2024.
{base}
Redacta compromiso juramentado de implementar programa de gerencia con profesional en ingeniería o arquitectura. Firmas.""",

            'F7B': f"""Redacta FORMATO 7B — DISPONIBILIDAD Y CONDICIONES DE MAQUINARIA completo según Res. 465/2024.
{base}
Redacta compromiso de disponibilidad de maquinaria requerida. Tabla de equipos y firmas.""",

            'F7C': f"""Redacta FORMATO 7C — PLAN DE CALIDAD completo según Res. 465/2024.
{base}
Redacta compromiso de implementar plan de calidad según normas técnicas. Firmas.""",

            'F8': f"""Redacta FORMATO 8 — VINCULACIÓN PERSONAS CON DISCAPACIDAD completo según Res. 465/2024.
{base}
Certifica número de trabajadores y personas con discapacidad. Referencia certificado Ministerio de Trabajo. Tabla y firmas.""",

            'F9A': f"""Redacta FORMATO 9A — PUNTAJE INDUSTRIA NACIONAL completo según Res. 465/2024.
{base}
Configuración: {json.dumps(f9_config)}
{'Con bienes relevantes: ' + ', '.join(f9_config.get('bienes_relevantes',[])) if f9_config.get('tiene_bienes_relevantes') else 'Sin bienes relevantes específicos'}
Proponente nacional. Ofrecimiento apoyo industria nacional. Tabla bienes y firmas.""",

            'F9B': f"""Redacta FORMATO 9B — COMPONENTE NACIONAL EN SERVICIOS EXTRANJEROS completo según Res. 465/2024.
{base}
Proponente extranjero incorporando componente nacional. Tabla y firmas.""",

            'F12': f"""Redacta FORMATO 12 — ACREDITACIÓN EMPRESAS DE MUJERES completo según Res. 465/2024.
{base}
Acreditación empresa de mujeres según normativa. Declaración y firmas.""",

            'F13': f"""Redacta FORMATO 13 — ACREDITACIÓN MIPYME completo según Res. 465/2024 Ley 590/2000.
{base}
Acreditación como micro/pequeña/mediana empresa para puntaje adicional. Declaración y firmas.""",

            'F14': f"""Redacta FORMATO 14 — CRITERIOS AMBIENTALES Y SOCIALES completo según Res. 465/2024 Decreto 142/2023.
{base}
Compromiso programa ambiental y social numeral 4.2.4 documento base. Firmas.""",

            'ANX4': f"""Redacta ANEXO 4 — PACTO DE TRANSPARENCIA completo según Res. 465/2024.
{base}
Todos los compromisos i) al xv): cumplir ley, buena fe, no falsificación, libre competencia, no colusión, no sobornos, preguntas escritas, lealtad, compostura audiencias. Firma al final.""",

            'F10': f"""Redacta FORMATO 10 — FACTORES DE DESEMPATE completo según Res. 465/2024.
{base}
Mipyme: {'Sí' if any(e['es_mipyme'] for e in empresas) else 'No'} | Mujeres: {'Sí' if any(e['tiene_mujeres'] for e in empresas) else 'No'} | Discapacidad: {'Sí' if any(e['tiene_discapacidad'] for e in empresas) else 'No'}
Declaración factores desempate aplicables. Firmas.""",

            'F11': f"""Redacta FORMATO 11 — AUTORIZACIÓN DATOS PERSONALES completo según Res. 465/2024 Ley 1581/2012.
{base}
Titulares: {', '.join([e['rep_legal'] or e['razon_social'] for e in empresas])}
Autorización tratamiento datos personales. Firmas.""",
        }
        return prompts.get(fid, '')

    def crear_word(fid, texto):
        doc = Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(11)
        header = doc.sections[0].header
        hp = header.paragraphs[0]
        hp.text = f"Colombia Compra Eficiente · Res. 465/2024 V4 · {pliego.get('numero_proceso','')} · {fid}"
        hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if hp.runs: hp.runs[0].font.size = Pt(9)
        nombres = {
            'F1':'FORMATO 1 — CARTA DE PRESENTACIÓN DE LA PROPUESTA',
            'F2':f"FORMATO {'2A' if tipo_prop=='consorcio' else '2B'} — CONFORMACIÓN DE PROPONENTE PLURAL",
            'F6':'FORMATO 6 — PAGOS DE SEGURIDAD SOCIAL Y APORTES LEGALES',
            'F7A':'FORMATO 7A — PROGRAMA DE GERENCIA DE PROYECTOS',
            'F7B':'FORMATO 7B — DISPONIBILIDAD Y CONDICIONES DE MAQUINARIA',
            'F7C':'FORMATO 7C — PLAN DE CALIDAD',
            'F8':'FORMATO 8 — VINCULACIÓN PERSONAS CON DISCAPACIDAD',
            'F9A':'FORMATO 9A — PUNTAJE DE INDUSTRIA NACIONAL',
            'F9B':'FORMATO 9B — INCORPORACIÓN COMPONENTE NACIONAL',
            'F12':'FORMATO 12 — ACREDITACIÓN EMPRESAS DE MUJERES',
            'F13':'FORMATO 13 — ACREDITACIÓN MIPYME',
            'F14':'FORMATO 14 — CRITERIOS AMBIENTALES Y SOCIALES',
            'ANX4':'ANEXO 4 — PACTO DE TRANSPARENCIA',
            'F10':'FORMATO 10 — FACTORES DE DESEMPATE',
            'F11':'FORMATO 11 — AUTORIZACIÓN DATOS PERSONALES',
        }
        h = doc.add_heading(nombres.get(fid, fid), level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph()
        for linea in texto.split('\n'):
            doc.add_paragraph(linea)
        footer = doc.sections[0].footer
        fp = footer.paragraphs[0]
        fp.text = f"Propuestas vías · {datetime.now().strftime('%d/%m/%Y')}"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if fp.runs: fp.runs[0].font.size = Pt(8)
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()

    # Generar ZIP con carpetas
    zip_buf = io.BytesIO()
    proceso = pliego.get('numero_proceso', 'proceso').replace('/', '-')
    nombres_archivo = {
        'F1':'Formato1_CartaPresentacion','F2':'Formato2_Consorcio_UT',
        'F6':'Formato6_SeguridadSocial','F7A':'Formato7A_GerenciaProyectos',
        'F7B':'Formato7B_Maquinaria','F7C':'Formato7C_PlanCalidad',
        'F8':'Formato8_Discapacidad','F9A':'Formato9A_IndustriaNacional',
        'F9B':'Formato9B_ComponenteNacional','F12':'Formato12_EmpresasMujeres',
        'F13':'Formato13_Mipyme','F14':'Formato14_AmbientalSocial',
        'ANX4':'Anexo4_PactoTransparencia','F10':'Formato10_Desempate',
        'F11':'Formato11_DatosPersonales',
    }
    habilitantes = ['F1','F2','F6']
    puntaje_fmts = ['F7A','F7B','F7C','F8','F9A','F9B','F12','F13','F14']
    otros_fmts = ['ANX4','F10','F11']

    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fid in formatos:
            prompt = prompt_fmt(fid)
            if not prompt: continue
            try:
                texto = call_claude(prompt)
                word = crear_word(fid, texto)
                carpeta = 'Habilitantes' if fid in habilitantes else ('Puntaje' if fid in puntaje_fmts else 'Otros')
                nombre = f"{proceso}/{carpeta}/{nombres_archivo.get(fid,fid)}.docx"
                zf.writestr(nombre, word)
            except: pass

        # Agregar documentos subidos de cada empresa
        for e in empresas:
            for doc in e.get('docs', []):
                if doc.get('url_cloudinary'):
                    try:
                        r = requests.get(doc['url_cloudinary'], timeout=15)
                        if r.ok:
                            carpeta = 'Habilitantes' if doc['tipo'] in ['camara','rup','cedula','redam','estados_financieros','capacidad_residual'] else 'Soportes'
                            nombre_doc = f"{proceso}/{carpeta}/{e['razon_social'].replace(' ','_')}_{doc['tipo']}_{doc['nombre_archivo']}"
                            zf.writestr(nombre_doc, r.content)
                    except: pass

            # Documentos de experiencia seleccionada
            for x in e.get('exp_sel', []):
                for doc in x.get('docs', []):
                    if doc.get('url_cloudinary'):
                        try:
                            r = requests.get(doc['url_cloudinary'], timeout=15)
                            if r.ok:
                                nombre_doc = f"{proceso}/Habilitantes/Experiencia/{e['razon_social'].replace(' ','_')}_{x['entidad'].replace(' ','_')}_{doc['nombre_archivo']}"
                                zf.writestr(nombre_doc, r.content)
                        except: pass

        # Documentos ingeniero avalador (temporales, no guardados en DB)
        for doc_ing in data.get('docs_ingeniero', []):
            if doc_ing.get('url'):
                try:
                    r = requests.get(doc_ing['url'], timeout=15)
                    if r.ok:
                        zf.writestr(f"{proceso}/Habilitantes/IngenieroAvalador_{doc_ing['tipo']}", r.content)
                except: pass

    zip_buf.seek(0)
    return send_file(zip_buf, mimetype='application/zip', as_attachment=True,
        download_name=f"Propuesta_{proceso}.zip")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
