import os, re, io, sqlite3, smtplib, threading, secrets
from datetime import datetime, date, timezone, timedelta
from email.message import EmailMessage
from flask import Flask, render_template, request, jsonify, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, 'database', 'rf.db')
XLSX = os.path.join(BASE, 'planilhas', 'clientes.xlsx')
if os.path.exists(os.path.join(BASE, '.env')):
    for l in open(os.path.join(BASE, '.env'), encoding='utf-8'):
        if '=' in l and not l.lstrip().startswith('#'):
            k, v = l.strip().split('=', 1); os.environ.setdefault(k, v)
env = lambda k, d='': os.environ.get(k, d)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024
app.jinja_env.filters['brl'] = lambda n: f"{n:,}".replace(',', '.') + ',00'

def conn():
    c = sqlite3.connect(DB); c.row_factory = sqlite3.Row; return c

def init_db():
    with conn() as c, open(os.path.join(BASE, 'schema.sql'), encoding='utf-8') as f:
        c.executescript(f.read())

# ---------- validação ----------
def cpf_ok(c):
    if len(c) != 11 or c == c[0] * 11: return False
    for n in (9, 10):
        s = sum(int(c[i]) * (n + 1 - i) for i in range(n))
        if int(c[n]) != (s * 10 % 11) % 10: return False
    return True

def validar(d):
    t = lambda k: str(d.get(k, '')).strip()
    dig = lambda k: re.sub(r'\D', '', t(k))
    v = dict(nome=' '.join(t('nome').split()), cpf=dig('cpf'), cnpj=dig('cnpj'), nascimento=t('nascimento'),
             email=t('email').lower(), telefone=dig('telefone'), endereco=t('endereco'), cidade=t('cidade'),
             uf=t('uf').upper(), descricao=t('descricao'), plano_id=t('plano_id'))
    e = {}
    if len(v['nome'].split()) < 2 or len(v['nome']) > 120: e['nome'] = 'Informe o nome completo.'
    if not cpf_ok(v['cpf']): e['cpf'] = 'CPF inválido.'
    if v['cnpj'] and len(v['cnpj']) != 14: e['cnpj'] = 'CNPJ deve ter 14 dígitos.'
    try:
        idade = (date.today() - date.fromisoformat(v['nascimento'])).days / 365.25
        if not 18 <= idade <= 120: raise ValueError
    except ValueError: e['nascimento'] = 'Data de nascimento inválida (mínimo 18 anos).'
    if not re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', v['email']): e['email'] = 'E-mail inválido.'
    if len(v['telefone']) not in (10, 11): e['telefone'] = 'Telefone com DDD, 10 ou 11 dígitos.'
    if len(v['endereco']) < 5: e['endereco'] = 'Informe o endereço.'
    if len(v['cidade']) < 2: e['cidade'] = 'Informe a cidade.'
    if not re.fullmatch(r'[A-Z]{2}', v['uf']): e['uf'] = 'UF inválida.'
    if len(v['descricao']) < 10 or len(v['descricao']) > 1000: e['descricao'] = 'Descreva o serviço (10 a 1000 caracteres).'
    if not d.get('consent'): e['consent'] = 'É necessário autorizar o uso dos dados.'
    return v, e

# ---------- planilha ----------
def safe(x):  # evita injeção de fórmulas no Excel
    return "'" + x if isinstance(x, str) and x[:1] in '=+-@' else x

def gerar_planilha():
    with conn() as c: rows = c.execute('SELECT * FROM vw_planilha').fetchall()
    wb = Workbook(); ws = wb.active; ws.title = 'Clientes'
    cols = [('Protocolo', 13), ('Data/hora do envio', 19), ('Nome completo', 32), ('CPF', 16), ('CNPJ', 19),
            ('Nascimento', 13), ('E-mail', 30), ('Telefone', 17), ('Endereço', 38), ('Cidade', 20), ('UF', 6),
            ('Plano de interesse', 34), ('Valor (R$)', 13), ('Descrição do serviço', 55), ('Status', 14)]
    ws.append([c for c, _ in cols])
    thin = Side(style='thin', color='DDE2E9')
    for i, (c, w) in enumerate(cols, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
        h = ws.cell(1, i); h.font = Font(bold=True, color='FFFFFF'); h.fill = PatternFill('solid', fgColor='0A62C4')
        h.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 26
    for r in rows:
        cpf, cnpj, tel = r['cpf'], r['cnpj'] or '', r['telefone']
        ws.append([r['protocolo'], datetime.fromisoformat(r['criado_em']).strftime('%d/%m/%Y %H:%M'), safe(r['nome']),
                   f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}",
                   f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}" if cnpj else '',
                   datetime.fromisoformat(r['nascimento']).strftime('%d/%m/%Y'), safe(r['email']),
                   f"({tel[:2]}) {tel[2:-4]}-{tel[-4:]}", safe(r['endereco']), safe(r['cidade']), r['uf'],
                   r['plano'], r['preco'], safe(r['descricao']), r['status']])
    for n, row in enumerate(ws.iter_rows(min_row=2), 2):
        for c in row:
            c.alignment = Alignment(vertical='top', wrap_text=True); c.border = Border(bottom=thin)
            if n % 2 == 0: c.fill = PatternFill('solid', fgColor='EEF1F5')
        row[12].number_format = '#,##0.00'
    ws.freeze_panes = 'D2'; ws.auto_filter.ref = ws.dimensions
    rs = wb.create_sheet('Resumo'); rs.column_dimensions['A'].width = 36; rs.column_dimensions['B'].width = 12
    rs.append(['Total de clientes', len(rows)]); rs.append([])
    for titulo, chave in (('Por status', 'status'), ('Por plano', 'plano')):
        rs.append([titulo, 'Qtd'])
        for cel in rs[rs.max_row]: cel.font = Font(bold=True, color='FFFFFF'); cel.fill = PatternFill('solid', fgColor='0A62C4')
        cont = {}
        for r in rows: cont[r[chave]] = cont.get(r[chave], 0) + 1
        for k, q in cont.items(): rs.append([k, q])
        rs.append([])
    rs['A1'].font = rs['B1'].font = Font(bold=True, size=13)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def enviar(lead, xls):
    try:
        os.makedirs(os.path.dirname(XLSX), exist_ok=True); open(XLSX, 'wb').write(xls)
        if not (env('SMTP_HOST') and env('OWNER_EMAIL')): return
        m = EmailMessage(); m['Subject'] = f"Novo cliente {lead['protocolo']} - {lead['nome']}"
        m['From'] = env('SMTP_USER'); m['To'] = env('OWNER_EMAIL')
        m.set_content(f"Novo cadastro no site RF Soluções Financeiras\n\nProtocolo: {lead['protocolo']}\nNome: {lead['nome']}\n"
                      f"Telefone: {lead['telefone']}\nPlano: {lead['plano']}\n\nA planilha completa e organizada está em anexo.")
        m.add_attachment(xls, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename='clientes_RF.xlsx')
        with smtplib.SMTP(env('SMTP_HOST'), int(env('SMTP_PORT', '587')), timeout=20) as s:
            s.starttls(); s.login(env('SMTP_USER'), env('SMTP_PASS')); s.send_message(m)
    except Exception as ex:
        app.logger.error('Falha ao enviar planilha: %s', ex)

# ---------- rotas ----------
@app.get('/')
def index():
    with conn() as c: planos = c.execute('SELECT * FROM planos ORDER BY ordem').fetchall()
    return render_template('index.html', planos=planos, wa=re.sub(r'\D', '', env('WHATSAPP', '5581982431435')))

@app.post('/api/leads')
def novo_lead():
    d = request.get_json(silent=True) or {}
    if d.get('site'): return jsonify(ok=True, protocolo='RF-000000')  # honeypot (robôs)
    v, e = validar(d)
    if e: return jsonify(ok=False, erros=e), 422
    agora = datetime.now(timezone(timedelta(hours=-3))).replace(tzinfo=None, microsecond=0).isoformat(' ')
    with conn() as c:
        p = c.execute('SELECT id, nome FROM planos WHERE id=?', (v['plano_id'],)).fetchone() if v['plano_id'].isdigit() else None
        cur = c.execute('INSERT INTO leads(criado_em,nome,cpf,cnpj,nascimento,email,telefone,endereco,cidade,uf,plano_id,descricao) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)',
                        (agora, v['nome'], v['cpf'], v['cnpj'] or None, v['nascimento'], v['email'], v['telefone'], v['endereco'],
                         v['cidade'], v['uf'], p['id'] if p else None, v['descricao']))
        prot = f"RF-{cur.lastrowid:06d}"; c.execute('UPDATE leads SET protocolo=? WHERE id=?', (prot, cur.lastrowid))
    xls = gerar_planilha()
    threading.Thread(target=enviar, args=({'protocolo': prot, 'nome': v['nome'], 'telefone': v['telefone'], 'plano': p['nome'] if p else 'Ainda não sei'}, xls), daemon=True).start()
    return jsonify(ok=True, protocolo=prot)

@app.get('/admin/planilha')
def baixar():
    a = request.authorization
    pw = env('ADMIN_PASSWORD')
    if not pw or not a or not (secrets.compare_digest(a.username, env('ADMIN_USER', 'admin')) and secrets.compare_digest(a.password or '', pw)):
        return Response('Acesso restrito', 401, {'WWW-Authenticate': 'Basic realm="RF"'})
    return send_file(io.BytesIO(gerar_planilha()), as_attachment=True, download_name='clientes_RF.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def init_db():
    os.makedirs(os.path.dirname(DB), exist_ok=True)
    with conn() as c, open(os.path.join(BASE, 'schema.sql'), encoding='utf-8') as f:
        c.executescript(f.read())