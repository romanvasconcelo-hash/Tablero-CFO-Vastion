# ============================================================================
# TABLERO CFO VASTION · App de Roberto  v0.3
# Carga + validación + REPORTE INTERNO de indicadores.
# Tras validar, muestra 3 indicadores clave (uno por eje), el P&L de Gestión
# y los indicadores de apoyo. Vista interna (Roberto/Vastion): muestra todo.
# ============================================================================
import streamlit as st
import psycopg2
from psycopg2.extras import execute_values
import xml.etree.ElementTree as ET
import openpyxl, hashlib, io
from datetime import datetime, date
 
st.set_page_config(page_title="Tablero CFO Vastion", page_icon="📊", layout="centered")
 
def get_conn():
    s = st.secrets["db"]
    return psycopg2.connect(host=s["host"], port=s["port"], dbname=s["dbname"],
                            user=s["user"], password=s["password"])
 
# ---------------------------------------------------------------------------
# PARSERS
# ---------------------------------------------------------------------------
def _ns(root): return root.tag[root.tag.find('{')+1:root.tag.find('}')]
def _sha(data): return hashlib.sha256(data).hexdigest()
 
def clasificar(nombre, data):
    if nombre.lower().endswith('.xml'):
        root = ET.fromstring(data); tag = root.tag.lower()
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
# CARGA + VALIDACIÓN
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
    conn = get_conn(); conn.autocommit = False; cur = conn.cursor()
    try:
        if 'CATALOGO' in archivos:
            nom, data = archivos['CATALOGO']; cat_rows = parse_catalogo(data)
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
        if not rcat: raise RuntimeError("No hay catálogo para este cliente. Sube el catálogo XML la primera vez.")
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
        return per, bal_id, integ, madz
    except Exception:
        conn.rollback(); raise
    finally:
        cur.close(); conn.close()
 
# ---------------------------------------------------------------------------
# INDICADORES (reporte interno)
# ---------------------------------------------------------------------------
def cargar_indicadores(bal_id, cli, per):
    conn = get_conn(); cur = conn.cursor(); ind = {}
    try:
        cur.execute("SELECT concepto,monto,pct FROM fn_pl_gestion(%s) ORDER BY orden", (bal_id,))
        pl = cur.fetchall()
        ind['pl'] = [(c, float(m), (float(p) if p is not None else None)) for c,m,p in pl]
        d = {c:(float(m), (float(p) if p is not None else None)) for c,m,p in pl}
        ing = d.get('Ingresos',(0,0))[0]
        lab = d.get('(-) Eficiencia Laboral',(0,0))[0]
        mb  = d.get('= MARGEN BRUTO DE GESTIÓN',(0,0))[0]
        ind['pretax_pct'] = d.get('= UTILIDAD ANTES DE IMPUESTOS',(0,None))[1]
        ind['mb_pct']     = d.get('= MARGEN BRUTO DE GESTIÓN',(0,None))[1]
        ind['nomina_pct'] = round(lab/ing*100,1) if ing else None
        ind['gpld']       = round(mb/lab,2) if lab and mb>0 else None
        cur.execute("SELECT concepto,valor FROM fn_cash_lag(%s)", (bal_id,))
        cl = {c:(float(v) if v is not None else None) for c,v in cur.fetchall()}
        ind['cash_lag'] = next((v for k,v in cl.items() if k.startswith('Cash Lag')), None)
        ind['caja']     = cl.get('Caja fin de mes')
        cur.execute("SELECT concepto,valor FROM fn_eiva(%s,%s)", (cli, per))
        ev = {c:(float(v) if v is not None else None) for c,v in cur.fetchall()}
        ind['eiva_pct']  = ev.get('EIVA % (acred/tras)')
        ind['iva_neto']  = ev.get('IVA neto (+cargo / -favor)')
        return ind
    finally:
        cur.close(); conn.close()
 
def sem_pretax(p):
    if p is None: return "—"
    if p >= 10: return "🟢 Sano (≥10%)"
    if p >= 5:  return "🟡 Mínimo (5–9%)"
    return "🔴 Peligro (<5%)"
 
def money(v):
    return "—" if v is None else f"${v:,.0f}"
 
 
def tendencia(cli):
    conn = get_conn(); cur = conn.cursor(); out = []
    try:
        cur.execute("SELECT periodo, archivo_vigente, estado FROM periodo_estado WHERE cliente_id=%s ORDER BY periodo", (cli,))
        for per, bal, estado in cur.fetchall():
            if bal is None: continue
            cur.execute("SELECT pl.pct FROM fn_pl_gestion(%s) pl WHERE pl.concepto='= UTILIDAD ANTES DE IMPUESTOS'", (bal,))
            r = cur.fetchone(); pretax = float(r[0]) if r and r[0] is not None else None
            cur.execute("SELECT concepto,valor FROM fn_cash_lag(%s)", (bal,))
            cl = {c:(float(v) if v is not None else None) for c,v in cur.fetchall()}
            caja = cl.get('Caja fin de mes')
            cash_lag = next((v for k,v in cl.items() if k.startswith('Cash Lag')), None)
            cur.execute("SELECT 1 FROM fn_validar_madurez(%s) WHERE paso=false LIMIT 1", (bal,)); adv_m = cur.fetchone() is not None
            cur.execute("SELECT 1 FROM fn_validar_periodo(%s) WHERE paso=false LIMIT 1", (bal,)); adv_i = cur.fetchone() is not None
            out.append(dict(periodo=str(per)[:7], pretax=pretax, caja=caja, cash_lag=cash_lag, flag=(adv_m or adv_i)))
        return out
    finally:
        cur.close(); conn.close()
 
 
def datos_reporte_cliente(cli, periodo):
    conn = get_conn(); cur = conn.cursor()
    try:
        cur.execute("SELECT archivo_vigente FROM periodo_estado WHERE cliente_id=%s AND periodo=%s", (cli, periodo))
        r = cur.fetchone()
        if not r or not r[0]: return None
        bal = r[0]
        cur.execute("""SELECT archivo_vigente FROM periodo_estado
                       WHERE cliente_id=%s AND periodo<%s AND archivo_vigente IS NOT NULL
                       ORDER BY periodo DESC LIMIT 1""", (cli, periodo))
        rp = cur.fetchone()
        caja_prev = None
        if rp:
            cur.execute("""SELECT COALESCE(SUM(b.saldo_final),0) FROM insumos_balanza i
                           JOIN raw_balanza b ON b.archivo_id=i.archivo_id AND b.num_cuenta=i.num_cuenta
                           WHERE i.archivo_id=%s AND i.es_hoja AND (i.cod_agrupador LIKE '101%%' OR i.cod_agrupador LIKE '102%%')""", (rp[0],))
            caja_prev = float(cur.fetchone()[0])
    finally:
        cur.close(); conn.close()
    ind = cargar_indicadores(bal, cli, periodo)
    ind['caja_prev'] = caja_prev
    ind['caja_var'] = (ind['caja'] - caja_prev) if (ind['caja'] is not None and caja_prev is not None) else None
    return ind
 
def periodos_cargados(cli):
    conn = get_conn(); cur = conn.cursor()
    try:
        cur.execute("""SELECT periodo FROM periodo_estado WHERE cliente_id=%s AND archivo_vigente IS NOT NULL
                       ORDER BY periodo DESC""", (cli,))
        return [str(r[0]) for r in cur.fetchall()]
    finally:
        cur.close(); conn.close()
 
 
def estados_financieros(archivo_id):
    """ER (NIF) + Balance General + EFE (indirecto) desde la balanza. Lógica validada en datos reales."""
    conn = get_conn(); cur = conn.cursor()
    try:
        cur.execute("""
            SELECT i.cod_agrupador, i.naturaleza::text, i.saldo_final,
                   r.saldo_inicial, r.debe, r.haber
            FROM insumos_balanza i
            JOIN raw_balanza r ON r.archivo_id=i.archivo_id AND r.num_cuenta=i.num_cuenta
            WHERE i.archivo_id=%s AND i.es_hoja AND NOT i.es_orden
        """, (archivo_id,))
        rows = cur.fetchall()
    finally:
        cur.close(); conn.close()
    acc = [dict(ag=(a or ''), nat=(n or ''), sf=float(sf or 0),
                si=float(si or 0), d=float(d or 0), h=float(h or 0))
           for a,n,sf,si,d,h in rows]
    p3 = lambda x: x['ag'][:3]; p1 = lambda x: x['ag'][:1]
    # --- Estado de Resultados (movimiento del mes) ---
    ing = sum(x['h']-x['d'] for x in acc if p3(x) in ('401','402','403'))
    cos = sum(x['d']-x['h'] for x in acc if p3(x) in ('501','502','503'))
    gas = sum(x['d']-x['h'] for x in acc if p3(x) in ('601','604'))
    dep = sum(x['d']-x['h'] for x in acc if p3(x)=='613')
    gf  = sum(x['d']-x['h'] for x in acc if p3(x)=='701')
    pf  = sum(x['h']-x['d'] for x in acc if p3(x)=='702')
    uai = ing-cos-gas-dep-gf+pf
    er = dict(ing=ing, cos=cos, ub=ing-cos, gas=gas, dep=dep, gf=gf, pf=pf, uai=uai)
    # --- Balance General (saldos firmados por naturaleza) ---
    ssf = lambda x: x['sf'] if x['nat']=='D' else -x['sf']
    res_ytd = -sum(ssf(x) for x in acc if p1(x) in ('4','5','6','7'))
    activo  = sum(ssf(x) for x in acc if p1(x)=='1')
    pasivo  = sum(-ssf(x) for x in acc if p1(x)=='2')
    cap_cta = sum(-ssf(x) for x in acc if p1(x)=='3')
    capital = cap_cta + res_ytd
    bg = dict(activo=activo, pasivo=pasivo, cap_cta=cap_cta, res=res_ytd,
              capital=capital, residual=activo-(pasivo+capital))
    # --- EFE (indirecto) ---
    def actividad(ag):
        if ag[:3] in ('101','102','103'): return 'EF'
        if ag.startswith(('202.01','205.06','107.05','251','252','253','254','255','256','257','258','259',
                          '301','302','303','304','305')): return 'FIN'
        if ag[:3] in ('150','151','152','153','154','155','156','157','158','159','171','172','184'): return 'INV'
        if ag[:1] in ('1','2'): return 'OP'
        if ag[:1]=='3': return 'FIN'
        return 'OTRO'
    def ce(x):
        delta = x['sf']-x['si']; return -delta if x['nat']=='D' else delta
    bk = {'OP':0.0,'INV':0.0,'FIN':0.0,'OTRO':0.0}
    for x in acc:
        if p1(x) in ('4','5','6','7'): continue
        k = actividad(x['ag'])
        if k=='EF': continue
        bk[k]+=ce(x)
    operacion = uai + bk['OP']
    dcash = sum(x['sf']-x['si'] for x in acc if p3(x) in ('101','102','103'))
    suma = operacion + bk['INV'] + bk['FIN']
    efe = dict(uai=uai, ct=bk['OP'], op=operacion, inv=bk['INV'], fin=bk['FIN'],
               suma=suma, dcash=dcash, plug=dcash-suma)
    return dict(er=er, bg=bg, efe=efe)
 
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
    st.info("Detecté: " + ", ".join(etiquetas.get(t, t) for t in archivos))
    if 'BALANZA' not in archivos:
        st.error("Falta la balanza (XML). Es obligatoria.")
 
if st.button("Cargar y validar", type="primary", disabled=not (subidos and 'BALANZA' in archivos)):
    with st.spinner("Cargando y validando…"):
        try:
            per, bal_id, integ, madz = procesar(cli_id, archivos)
            ind = cargar_indicadores(bal_id, cli_id, per)
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
 
    # ---- 3 indicadores clave, uno por eje ----
    st.subheader("Indicadores clave")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.caption("EJE 1 · GENERACIÓN")
        st.metric("Pre-Tax Profit %", f"{ind['pretax_pct']}%" if ind['pretax_pct'] is not None else "—")
        st.caption(sem_pretax(ind['pretax_pct']))
    with c2:
        st.caption("EJE 2 · TRASLADO")
        st.metric("Cash Lag", money(ind['cash_lag']))
        if ind['cash_lag'] is None:
            st.caption("Necesita mes previo")
        elif ind['cash_lag'] > 0:
            st.caption("🔴 Caja cayó más que la utilidad")
        else:
            st.caption("🟢 Caja por encima de la utilidad")
    with c3:
        st.caption("EJE 3 · FISCAL")
        st.metric("EIVA %", f"{ind['eiva_pct']}%" if ind['eiva_pct'] is not None else "—")
        if ind['iva_neto'] is not None:
            st.caption(("IVA a cargo " + money(ind['iva_neto'])) if ind['iva_neto'] > 0
                       else ("IVA a favor " + money(abs(ind['iva_neto']))))
 
    # ---- indicadores de apoyo ----
    st.subheader("Apoyo")
    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Margen Bruto %", f"{ind['mb_pct']}%" if ind['mb_pct'] is not None else "—")
    a2.metric("Nómina / Ingresos", f"{ind['nomina_pct']}%" if ind['nomina_pct'] is not None else "—")
    a3.metric("GPLD", f"{ind['gpld']}x" if ind['gpld'] is not None else "n/a")
    a4.metric("Caja fin de mes", money(ind['caja']))
 
    # ---- P&L de Gestión ----
    with st.expander("P&L de Gestión (no constituye estado de resultados NIF)"):
        st.dataframe(
            [{"Concepto": c, "Monto": f"{m:,.0f}", "%": (f"{p}%" if p is not None else "")} for c,m,p in ind['pl']],
            use_container_width=True, hide_index=True)
 
    # ---- detalle de compuertas ----
    with st.expander("Detalle de validación"):
        st.markdown("**Integridad**")
        for p, ok, sev, det in integ:
            st.write((("✅" if ok else ("🟡" if sev=='ADVERTENCIA' else "❌"))) + f"  **{p}** — {det}")
        st.markdown("**Madurez analítica**")
        for p, ok, sev, det in madz:
            st.write((("✅" if ok else ("🟡" if sev=='ADVERTENCIA' else "❌"))) + f"  **{p}** ({sev}) — {det}")
 
 
# ---------------------------------------------------------------------------
# COMPARATIVO DE MESES CARGADOS
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Tendencia de los meses cargados")
if st.button("Ver tendencia"):
    datos = tendencia(cli_id)
    if not datos:
        st.info("Este cliente no tiene meses cargados todavía.")
    else:
        filas = [{"Mes": d["periodo"],
                  "Caja": money(d["caja"]),
                  "Pre-Tax %": (f"{d['pretax']}%" if d["pretax"] is not None else "—"),
                  "Cash Lag": money(d["cash_lag"]),
                  "Comparable": ("⚠️ con observaciones" if d["flag"] else "✓ limpio")} for d in datos]
        st.dataframe(filas, use_container_width=True, hide_index=True)
        import pandas as pd
        caja_df = pd.DataFrame([{"Mes": d["periodo"], "Caja": d["caja"]} for d in datos if d["caja"] is not None]).set_index("Mes")
        if not caja_df.empty:
            st.caption("Caja al cierre por mes")
            st.line_chart(caja_df)
        if any(d["flag"] for d in datos):
            st.caption("⚠️ Los meses marcados traen observaciones de contabilidad (en recontabilización). La tendencia puede no reflejar la operación real hasta que se ajusten.")
 
 
# ---------------------------------------------------------------------------
# REPORTE DEL CLIENTE (vista dueño)
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Reporte del cliente")
_periodos = periodos_cargados(cli_id)
if not _periodos:
    st.caption("Carga al menos un mes para generar el reporte del cliente.")
else:
    mes_rep = st.selectbox("Mes a reportar", _periodos, key="rep_mes")
    if st.button("Generar reporte del cliente"):
        ind = datos_reporte_cliente(cli_id, mes_rep)
        if ind is None:
            st.error("No se encontró ese mes.")
        else:
            st.markdown(f"### {nombre_sel}  ·  {mes_rep[:7]}")
            # caja al centro
            if ind["caja_var"] is not None:
                st.metric("Tu caja al cierre del mes", money(ind["caja"]), delta=f"{ind['caja_var']:,.0f} vs mes anterior")
                direccion = "bajó" if ind["caja_var"] < 0 else "subió"
                st.write(f"Tu caja {direccion} **{money(abs(ind['caja_var']))}** respecto al mes anterior.")
            else:
                st.metric("Tu caja al cierre del mes", money(ind["caja"]))
                st.caption("Sin mes anterior para comparar.")
            # tres indicadores en lenguaje de dueño
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Rentabilidad del mes", f"{ind['pretax_pct']}%" if ind["pretax_pct"] is not None else "—")
                st.caption(sem_pretax(ind["pretax_pct"]))
            with c2:
                st.metric("Margen de tu operación", f"{ind['mb_pct']}%" if ind["mb_pct"] is not None else "—")
            with c3:
                st.metric("Brecha utilidad vs. caja", money(ind["cash_lag"]))
                if ind["cash_lag"] is not None and ind["cash_lag"] > 0:
                    st.caption("🔴 La utilidad no llegó a caja")
                elif ind["cash_lag"] is not None:
                    st.caption("🟢 La caja subió con la utilidad")
            # lectura del mes (la escribe Roberto)
            st.markdown("**Lo que esto significa** (escríbelo para el dueño)")
            st.text_area("Lectura del mes", key="rep_lectura",
                         placeholder="Ej.: El negocio es rentable, pero la utilidad se quedó en cobranza. Prioridad: cobrar, no vender.",
                         label_visibility="collapsed")
            st.caption("Los números los pone el sistema. La lectura la escribe el CFO.")
 
 
# ---------------------------------------------------------------------------
# ESTADOS FINANCIEROS FORMALES (NIF)
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Estados financieros formales (NIF)")
st.caption("Respaldo formal del reporte. Para clientes con observaciones, son diagnóstico interno hasta la recontabilización.")
_pf = periodos_cargados(cli_id)
if not _pf:
    st.caption("Carga al menos un mes para generar los estados financieros.")
else:
    mes_ef = st.selectbox("Mes", _pf, key="ef_mes")
    if st.button("Generar estados financieros"):
        conn=get_conn(); cur=conn.cursor()
        try:
            cur.execute("SELECT archivo_vigente FROM periodo_estado WHERE cliente_id=%s AND periodo=%s",(cli_id,mes_ef))
            rr=cur.fetchone(); bal=rr[0] if rr else None
        finally:
            cur.close(); conn.close()
        if not bal:
            st.error("No se encontró la balanza de ese mes.")
        else:
            ef = estados_financieros(bal)
            f = lambda v: f"${v:,.0f}" if v is not None else "—"
            er, bg, efe = ef["er"], ef["bg"], ef["efe"]
            ingp = (lambda d: f" ({d/er['ing']*100:.1f}%)" if er['ing'] else "")
 
            st.markdown(f"#### Estado de Resultados · {mes_ef[:7]}")
            st.markdown(
                f"| Concepto | Monto |\n|---|--:|\n"
                f"| Ingresos | {f(er['ing'])} |\n"
                f"| (−) Costo de ventas | {f(er['cos'])} |\n"
                f"| **= Utilidad bruta** | **{f(er['ub'])}**{ingp(er['ub'])} |\n"
                f"| (−) Gastos de operación | {f(er['gas'])} |\n"
                f"| (−) Depreciación | {f(er['dep'])} |\n"
                f"| (−) Gasto financiero | {f(er['gf'])} |\n"
                f"| (+) Producto financiero | {f(er['pf'])} |\n"
                f"| **= Utilidad antes de impuestos** | **{f(er['uai'])}**{ingp(er['uai'])} |")
 
            st.markdown(f"#### Balance General · {mes_ef[:7]}")
            st.markdown(
                f"| Concepto | Monto |\n|---|--:|\n"
                f"| ACTIVO | {f(bg['activo'])} |\n"
                f"| PASIVO | {f(bg['pasivo'])} |\n"
                f"| Capital (cuentas) | {f(bg['cap_cta'])} |\n"
                f"| (+) Resultado del ejercicio | {f(bg['res'])} |\n"
                f"| **= Capital contable** | **{f(bg['capital'])}** |\n"
                f"| Pasivo + Capital | {f(bg['pasivo']+bg['capital'])} |")
            if abs(bg['residual']) < 1:
                st.success(f"Balance cuadrado (residual {f(bg['residual'])}).")
            else:
                st.warning(f"Residual de cuadre: {f(bg['residual'])} — saldos desfasados (revisar R-EC-10 / naturaleza).")
 
            st.markdown(f"#### Estado de Flujo de Efectivo · {mes_ef[:7]} (indirecto)")
            st.markdown(
                f"| Concepto | Monto |\n|---|--:|\n"
                f"| Utilidad antes de impuestos | {f(efe['uai'])} |\n"
                f"| (+/−) Cambios en capital de trabajo | {f(efe['ct'])} |\n"
                f"| **= Flujo de operación** | **{f(efe['op'])}** |\n"
                f"| Flujo de inversión | {f(efe['inv'])} |\n"
                f"| Flujo de financiamiento | {f(efe['fin'])} |\n"
                f"| **= Variación de caja calculada** | **{f(efe['suma'])}** |\n"
                f"| Variación real de caja | {f(efe['dcash'])} |\n"
                f"| Partida por identificar | {f(efe['plug'])} |")
            if abs(efe['plug']) < 1:
                st.success("EFE cuadrado (partida por identificar $0).")
            else:
                st.warning(f"Partida por identificar: {f(efe['plug'])} — clasificación incompleta.")
            st.text_area("Lectura del mes", key="rep_lectura",
                         placeholder="Ej.: El negocio es rentable, pero la utilidad se quedó en cobranza. Prioridad: cobrar, no vender.",
                         label_visibility="collapsed")
            st.caption("Los números los pone el sistema. La lectura la escribe el CFO.")
