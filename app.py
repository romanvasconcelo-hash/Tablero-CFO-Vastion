# ============================================================================
# TABLERO CFO VASTION · Cargador de Roberto  v0.2
# Pantalla de arrastrar-y-soltar. Roberto sube los archivos del mes,
# da un clic, y ve el veredicto: VALIDADO / VALIDADO con advertencias / RETENIDO.
# Reusa los parsers y compuertas ya validados contra dato real.
#
# v0.2: integridad respeta severidad (continuidad ya no bloquea, solo advierte).
# ============================================================================
import streamlit as st
import psycopg2
from psycopg2.extras import execute_values
import xml.etree.ElementTree as ET
import openpyxl, hashlib, io
from datetime import datetime, date

st.set_page_config(page_title="Tablero CFO Vastion · Carga", page_icon="📊", layout="centered")

# ---------------------------------------------------------------------------
# CONEXIÓN (desde st.secrets — configurada en Streamlit Cloud, no aquí)
# ---------------------------------------------------------------------------
def get_conn():
    s = st.secrets["db"]
    return psycopg2.connect(host=s["host"], port=s["port"], dbname=s["dbname"],
                            user=s["user"], password=s["password"])

# ---------------------------------------------------------------------------
# PARSERS (validados; leen desde bytes del archivo subido)
# ---------------------------------------------------------------------------
def _ns(root): return root.tag[root.tag.find('{')+1:root.tag.find('}')]
def _sha(data): return hashlib.sha256(data).hexdigest()

def clasificar(nombre, data):
    """Detecta qué es cada archivo por su contenido, no por su nombre."""
    if nombre.lower().endswith('.xml'):
        root = ET.fromstring(data)
        tag = root.tag.lower()
        if 'balanza' in tag: return 'BALANZA'
        if 'catalogo' in tag or 'catálogo' in tag: return 'CATALOGO'
        return 'XML_DESCONOCIDO'
    if nombre.lower().endswith('.xlsx'):
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True); s = wb.sheetnames[0].upper(); wb.close()
        if 'INGRESO' in s: return 'CFDI_EMITIDO'
        if 'EGRESO' in s:  return 'CFDI_RECIBIDO'
        return 'XLSX_DESCONOCIDO'
    return 'DESCONOCIDO'

def parse_catalogo(data):
    root = ET.fromstring(data); ns = {'x': _ns(root)}
    return [dict(num_cuenta=c.attrib['NumCta'], descripcion=c.attrib.get('Desc'),
        cod_agrupador=c.attrib.get('CodAgrup'), subcuenta_de=c.attrib.get('SubCtaDe'),
        nivel=int(c.attrib['Nivel']) if c.attrib.get('Nivel') else None,
        naturaleza=c.attrib.get('Natur')) for c in root.findall('x:Ctas', ns)]

def parse_balanza(data):
    root = ET.fromstring(data); ns = {'x': _ns(root)}
    meta = dict(rfc=root.attrib.get('RFC'), anio=root.attrib.get('Anio'),
                mes=root.attrib.get('Mes'), envio=root.attrib.get('TipoEnvio'))
    rows = [(c.attrib['NumCta'], float(c.attrib['SaldoIni']), float(c.attrib['Debe']),
             float(c.attrib['Haber']), float(c.attrib['SaldoFin'])) for c in root.findall('x:Ctas', ns)]
    return meta, rows

def _num(v):
    try: return float(v) if v not in (None, '') else None
    except: return None
def _fecha(v):
    if isinstance(v, (datetime, date)): return v
    if not v: return None
    s = str(v)[:19]
    for f in ('%d-%m-%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%dT%H:%M:%S','%Y-%m-%d','%d/%m/%Y'):
        try: return datetime.strptime(s, f)
        except: pass
    return None
def _col(h, n):
    for i,c in enumerate(h):
        if c and n.lower()==str(c).strip().lower(): return i
    for i,c in enumerate(h):
        if c and n.lower() in str(c).strip().lower(): return i

def parse_cfdi(data):
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]; direc = 'EMITIDO' if 'INGRESO' in wb.sheetnames[0].upper() else 'RECIBIDO'
    rows = list(ws.iter_rows(values_only=True)); h = rows[3]; wb.close()
    C = {k:_col(h,v) for k,v in dict(rfc='RFC',nom='Razón social',ftim='Fecha Timbrado',
        fexp='Fecha Expedición',uuid='UUID',est='Estatus Sat',tipo='Tipo',uso='Uso CFDI',
        prod='Producto',sub='Subtotal sin descuentos',desc='Descuento',iva16='IVA 16%',iva8='IVA 8%',
        ivar='IVA Retención',isrr='ISR Retención',total='Total',cc='Cuenta Contable',
        cco='Centro de Costos',uuidr='UUID Relacionado').items()}
    out = []; cont = {}
    for r in rows[4:]:
        u = r[C['uuid']]
        if not u: continue
        cont[u] = cont.get(u,0)+1
        cc = r[C['cc']]; acc = str(cc).split(' - ')[0].strip() if cc else None
        fx = _fecha(r[C['fexp']]); per = date(fx.year, fx.month, 1) if fx else None
        out.append((u,cont[u],direc,per,fx,_fecha(r[C['ftim']]),r[C['tipo']],r[C['uso']],r[C['est']],
            r[C['rfc']], str(r[C['nom']])[:120] if r[C['nom']] else None,
            str(r[C['prod']])[:200] if r[C['prod']] else None,
            _num(r[C['sub']]),_num(r[C['desc']]),_num(r[C['iva16']]),_num(r[C['iva8']]),
            _num(r[C['ivar']]),_num(r[C['isrr']]),_num(r[C['total']]),acc,r[C['cco']],r[C['uuidr']]))
    return out

# ---------------------------------------------------------------------------
# CARGA + VALIDACIÓN (misma lógica que el cargador validado)
# ---------------------------------------------------------------------------
def registrar(cur, cli, per, tipo, envio, nombre, data):
    cur.execute("""INSERT INTO origen_archivo (cliente_id,periodo,tipo,envio,storage_path,hash_sha256,bytes,cargado_por)
        VALUES (%s,%s,%s,%s,%s,%s,%s,'streamlit') ON CONFLICT (cliente_id,periodo,tipo,hash_sha256) DO NOTHING RETURNING id""",
        (cli, per, tipo, envio, nombre, _sha(data), len(data)))
    row = cur.fetchone()
    if row: return row[0], True
    cur.execute("SELECT id FROM origen_archivo WHERE cliente_id=%s AND periodo=%s AND tipo=%s AND hash_sha256=%s",
                (cli, per, tipo, _sha(data))); return cur.fetchone()[0], False

def procesar(cli, archivos):
    """archivos: dict tipo -> (nombre, data). Devuelve (periodo, integridad, madurez)."""
    conn = get_conn(); conn.autocommit = False; cur = conn.cursor()
    try:
        if 'CATALOGO' in archivos:
            nom, data = archivos['CATALOGO']
            cat_rows = parse_catalogo(data)
        nomb, datab = archivos['BALANZA']
        meta, brows = parse_balanza(datab); per = f"{meta['anio']}-{meta['mes']}-01"
        envio = 'COMPLEMENTARIA' if meta['envio']=='C' else 'NORMAL'
        cur.execute("UPDATE cliente SET rfc=%s WHERE id=%s AND (rfc IS NULL OR rfc='')", (meta['rfc'], cli))
        if 'CATALOGO' in archivos:
            cat_id, cat_new = registrar(cur, cli, per, 'CATALOGO', None, nom, data)
            if cat_new:
                execute_values(cur, """INSERT INTO raw_catalogo (archivo_id,cliente_id,num_cuenta,descripcion,cod_agrupador,subcuenta_de,nivel,naturaleza)
                    VALUES %s ON CONFLICT DO NOTHING""",
                    [(cat_id,cli,c['num_cuenta'],c['descripcion'],c['cod_agrupador'],c['subcuenta_de'],c['nivel'],c['naturaleza']) for c in cat_rows])
        cur.execute("""SELECT id FROM origen_archivo WHERE cliente_id=%s AND tipo='CATALOGO' AND periodo<=%s
                       ORDER BY periodo DESC,version DESC LIMIT 1""", (cli, per))
        rcat = cur.fetchone()
        if not rcat:
            raise RuntimeError("No hay catálogo para este cliente. Sube el catálogo XML la primera vez.")
        cat_id = rcat[0]
        bal_id, bal_new = registrar(cur, cli, per, 'BALANZA', envio, nomb, datab)
        if bal_new:
            execute_values(cur, "INSERT INTO raw_balanza (archivo_id,num_cuenta,saldo_inicial,debe,haber,saldo_final) VALUES %s ON CONFLICT DO NOTHING",
                [(bal_id,)+r for r in brows])
        cur.execute("DELETE FROM insumos_balanza WHERE archivo_id=%s", (bal_id,))
        cur.execute("""INSERT INTO insumos_balanza (archivo_id,cliente_id,periodo,num_cuenta,cod_agrupador,naturaleza,es_hoja,es_orden,bloque,es_laboral,saldo_final)
            SELECT b.archivo_id,%(cli)s,%(per)s,b.num_cuenta,c.cod_agrupador,c.naturaleza,
              (b.num_cuenta NOT IN (SELECT subcuenta_de FROM raw_catalogo WHERE archivo_id=%(cat)s AND subcuenta_de IS NOT NULL)),
              (b.num_cuenta LIKE '8%%'),fb.bloque,fb.es_laboral,b.saldo_final
            FROM raw_balanza b JOIN raw_catalogo c ON c.archivo_id=%(cat)s AND c.num_cuenta=b.num_cuenta
            CROSS JOIN LATERAL fn_bloque(%(cli)s,c.cod_agrupador) fb WHERE b.archivo_id=%(bal)s""",
            dict(cli=cli, per=per, cat=cat_id, bal=bal_id))
        cur.execute("""INSERT INTO periodo_estado (cliente_id,periodo,estado,archivo_vigente) VALUES (%s,%s,'RECIBIDO',%s)
                       ON CONFLICT (cliente_id,periodo) DO UPDATE SET archivo_vigente=EXCLUDED.archivo_vigente""", (cli, per, bal_id))
        for tipo in ('CFDI_EMITIDO', 'CFDI_RECIBIDO'):
            if tipo not in archivos: continue
            nomc, datac = archivos[tipo]; crows = parse_cfdi(datac)
            if not crows: continue
            pr = crows[0][3]
            arch, new = registrar(cur, cli, pr, 'CFDI', None, nomc, datac)
            if new:
                execute_values(cur, """INSERT INTO raw_cfdi (uuid,renglon,direccion,periodo,fecha_emision,fecha_timbrado,tipo_cfdi,uso_cfdi,estatus_sat,contraparte_rfc,contraparte_nom,concepto,subtotal,descuento,iva_16,iva_8,iva_retenido,isr_retenido,total,cuenta_contable,centro_costos,uuid_relacionado,cliente_id,archivo_id)
                    VALUES %s ON CONFLICT DO NOTHING""", [t+(cli,arch) for t in crows])
        cur.execute("SELECT prueba,paso,severidad,detalle FROM fn_validar_periodo(%s)", (bal_id,)); integ = cur.fetchall()
        cur.execute("SELECT prueba,paso,severidad,detalle FROM fn_validar_madurez(%s)", (bal_id,)); madz = cur.fetchall()
        bloqueo = any((not ok) and sev=='BLOQUEANTE' for _,ok,sev,_ in integ) or any((not ok) and sev=='BLOQUEANTE' for _,ok,sev,_ in madz)
        if not bloqueo:
            cur.execute("UPDATE periodo_estado SET estado='VALIDADO',validado_en=now() WHERE cliente_id=%s AND periodo=%s", (cli, per))
        conn.commit()
        return per, integ, madz
    except Exception:
        conn.rollback(); raise
    finally:
        cur.close(); conn.close()

# ---------------------------------------------------------------------------
# INTERFAZ
# ---------------------------------------------------------------------------
st.title("📊 Tablero CFO Vastion")
st.caption("Carga mensual · arrastra los archivos del cliente y valida")

try:
    conn = get_conn(); cur = conn.cursor()
    cur.execute("SELECT id, nombre FROM cliente ORDER BY nombre")
    clientes = cur.fetchall(); cur.close(); conn.close()
except Exception as e:
    st.error(f"No se pudo conectar a la base. Revisa los secrets. ({e})"); st.stop()

if not clientes:
    st.warning("No hay clientes registrados."); st.stop()

nombre_sel = st.selectbox("Cliente", [c[1] for c in clientes])
cli_id = dict((c[1], c[0]) for c in clientes)[nombre_sel]

st.markdown("**Suelta aquí los archivos del mes** (balanza XML obligatoria; catálogo XML la primera vez; CFDI emitidos y recibidos):")
subidos = st.file_uploader("Arrastra los archivos", accept_multiple_files=True,
                           type=['xml','xlsx'], label_visibility='collapsed')

archivos = {}
if subidos:
    for f in subidos:
        data = f.getvalue(); tipo = clasificar(f.name, data)
        archivos[tipo] = (f.name, data)
    etiquetas = {'BALANZA':'Balanza','CATALOGO':'Catálogo','CFDI_EMITIDO':'CFDI emitidos','CFDI_RECIBIDO':'CFDI recibidos'}
    detectados = [etiquetas.get(t, t) for t in archivos]
    st.info("Detecté: " + ", ".join(detectados))
    if 'BALANZA' not in archivos:
        st.error("Falta la balanza (XML). Es obligatoria.")

if st.button("Cargar y validar", type="primary", disabled=not (subidos and 'BALANZA' in archivos)):
    with st.spinner("Cargando y validando…"):
        try:
            per, integ, madz = procesar(cli_id, archivos)
        except Exception as e:
            st.error(f"Error al procesar: {e}"); st.stop()

    bloqueo = any((not ok) and sev=='BLOQUEANTE' for _,ok,sev,_ in integ) or any((not ok) and sev=='BLOQUEANTE' for _,ok,sev,_ in madz)
    adv = any((not ok) and sev=='ADVERTENCIA' for _,ok,sev,_ in integ) or any((not ok) and sev=='ADVERTENCIA' for _,ok,sev,_ in madz)

    st.divider()
    if bloqueo:
        st.error(f"## 🔴 RETENIDO · {per}\nEste periodo no entra al tablero. Corregir en contabilidad antes de reportar.")
    elif adv:
        st.warning(f"## 🟡 VALIDADO con advertencias · {per}\nEl periodo entra, pero revisa las señales amarillas.")
    else:
        st.success(f"## 🟢 VALIDADO · {per}\nEl periodo entra al tablero.")

    st.subheader("Integridad")
    for p, ok, sev, det in integ:
        icono = "✅" if ok else ("🟡" if sev=='ADVERTENCIA' else "❌")
        st.write(f"{icono}  **{p}** — {det}")
    st.subheader("Madurez analítica")
    for p, ok, sev, det in madz:
        icono = "✅" if ok else ("🟡" if sev=='ADVERTENCIA' else "❌")
        st.write(f"{icono}  **{p}** ({sev}) — {det}")
