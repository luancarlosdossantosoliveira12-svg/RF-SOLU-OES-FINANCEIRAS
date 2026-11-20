const f = document.getElementById('lead'), msg = document.getElementById('msg');
const mask = (el, fn) => el.addEventListener('input', () => el.value = fn(el.value.replace(/\D/g, '')));
mask(f.cpf, d => d.slice(0,11).replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d{1,2})$/,'$1-$2'));
mask(f.cnpj, d => d.slice(0,14).replace(/(\d{2})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1.$2').replace(/(\d{3})(\d)/,'$1/$2').replace(/(\d{4})(\d{1,2})$/,'$1-$2'));
mask(f.telefone, d => d.slice(0,11).replace(/(\d{2})(\d)/,'($1) $2').replace(/(\d{4,5})(\d{4})$/,'$1-$2'));
f.uf.addEventListener('input', () => f.uf.value = f.uf.value.replace(/[^a-z]/gi,'').toUpperCase());

document.querySelectorAll('[data-plano]').forEach(a => a.addEventListener('click', () => f.plano_id.value = a.dataset.plano));
document.querySelectorAll('.plan').forEach(p => p.addEventListener('pointermove', e => {
  const r = p.getBoundingClientRect(); p.style.setProperty('--x', e.clientX - r.left + 'px'); p.style.setProperty('--y', e.clientY - r.top + 'px');
}));

function limpaErros() { f.querySelectorAll('.fe').forEach(x => x.remove()); f.querySelectorAll('.err').forEach(x => x.classList.remove('err')); }
function erros(e) {
  limpaErros();
  Object.entries(e).forEach(([k, t]) => {
    const el = f.elements[k]; if (!el) return;
    el.classList.add('err'); const s = document.createElement('span'); s.className = 'fe'; s.textContent = t;
    (el.closest('label') || el).appendChild(s);
  });
  const p = f.querySelector('.err, .fe'); if (p) p.scrollIntoView({behavior: 'smooth', block: 'center'});
}

f.addEventListener('submit', async ev => {
  ev.preventDefault(); msg.textContent = ''; msg.className = '';
  const wa = f.dataset.wa;
  const d = Object.fromEntries(new FormData(f)); d.consent = f.consent.checked;
  const btn = f.querySelector('button'); btn.disabled = true; btn.textContent = 'Enviando...';
  try {
    const r = await fetch('/api/leads', {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(d)});
    const j = await r.json();
    if (!r.ok) { erros(j.erros || {}); msg.className = 'bad'; msg.textContent = 'Revise os campos destacados.'; return; }
    f.innerHTML = `<div class="done"><div class="tick"><svg viewBox="0 0 52 52"><path d="M12 27l9 9 19-20"/></svg></div>
      <div class="big">Dados enviados!</div><p>Seu protocolo é <b>${j.protocolo}</b>. Entraremos em contato em breve.</p>
      <a class="btn wa" target="_blank" rel="noopener" href="https://wa.me/${wa}?text=${encodeURIComponent('Olá, enviei meus dados no site. Protocolo ' + j.protocolo)}">Falar com a RF no WhatsApp</a></div>`;
  } catch { msg.className = 'bad'; msg.textContent = 'Não foi possível enviar agora. Tente novamente.'; }
  finally { btn.disabled = false; btn.textContent = 'Enviar meus dados'; }
});