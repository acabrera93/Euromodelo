/* ===========================================================
   XX Euromodelo Joven 2026 — main.js
   Funciones compartidas por todas las páginas del sitio.
   IMPORTANTE: este sitio usa localStorage (no window.storage),
   porque está pensado para alojarse en un servidor propio, no
   dentro del panel de artefactos de Claude.ai.
   =========================================================== */

/* ---------- Menú móvil ---------- */
document.addEventListener('DOMContentLoaded', () => {
  const menuToggle = document.getElementById('menuToggle');
  const navLinks = document.getElementById('navLinks');
  if (menuToggle && navLinks) {
    menuToggle.addEventListener('click', () => navLinks.classList.toggle('open'));
  }
});

/* ---------- Header: menú desplegable "Euromodelo Joven 2026" ---------- */
document.addEventListener('DOMContentLoaded', () => {
  const toggle = document.getElementById('navDropdownToggle');
  const menu = document.getElementById('navDropdownMenu');
  if (!toggle || !menu) return;
  toggle.setAttribute('aria-expanded', 'false');
  toggle.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = menu.classList.toggle('open');
    toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
  });
  document.addEventListener('click', (e) => {
    if (menu.classList.contains('open') && !menu.contains(e.target) && e.target !== toggle) {
      menu.classList.remove('open');
      toggle.setAttribute('aria-expanded', 'false');
    }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      menu.classList.remove('open');
      toggle.setAttribute('aria-expanded', 'false');
    }
  });
});

/* ---------- Modal genérico (fábrica de controladores) ---------- */
function createModalController(id) {
  const el = document.getElementById(id);
  const controller = {
    el,
    open() { if (el) el.classList.add('open'); },
    close() { if (el) el.classList.remove('open'); },
  };
  if (el) {
    const closeBtn = el.querySelector('.modal-close');
    if (closeBtn) closeBtn.addEventListener('click', () => controller.close());
    el.addEventListener('click', (e) => { if (e.target === el) controller.close(); });
  }
  return controller;
}
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    document.querySelectorAll('.modal-overlay.open').forEach(m => m.classList.remove('open'));
  }
});

function renderDetailModal(modalController, item) {
  const root = modalController.el;
  const iconEl = root.querySelector('.js-modal-icon');
  const titleEl = root.querySelector('.js-modal-title');
  const introEl = root.querySelector('.js-modal-intro');
  const topicsEl = root.querySelector('.js-modal-topics');
  if (iconEl) {
    iconEl.textContent = item.code || '';
    iconEl.style.background = item.color || '';
    iconEl.style.color = item.color ? '#FFFFFF' : '';
  }
  if (titleEl) titleEl.textContent = item.title || '';
  if (introEl) introEl.textContent = item.intro || item.desc || '';
  if (topicsEl) {
    if (item.topics) {
      topicsEl.innerHTML = item.topics.map((t, i) =>
        `<li><span class="t-num">${String(i + 1).padStart(2, '0')}</span><span><b>${t[0]}.</b> ${t[1]}</span></li>`
      ).join('');
    } else if (item.paragraphs) {
      topicsEl.innerHTML = item.paragraphs.map(p => `<li><span>${p}</span></li>`).join('');
    } else {
      topicsEl.innerHTML = '';
    }
  }
  modalController.open();
}

/* ---------- Utilidades ---------- */
function refCode(prefix) {
  return prefix + '-' + Math.random().toString(36).slice(2, 7).toUpperCase();
}

function randomAlphaNum(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let out = '';
  for (let i = 0; i < len; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

function slugifyName(str) {
  return (str || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/* ---------- Backend (Google Apps Script) ---------- */
const EUROMODELO_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbztpSnDRFqR1JAvh4jmNyLcUo05Ur3PoRckFeavmP6PqezdYTW0QNTz172gG20U6vYsPg/exec';

/* ---------- Autenticación (piloto) ----------
   El usuario es el correo del participante. localStorage guarda una copia local
   (euromodelo_users) para acceso rápido en el mismo dispositivo; el backend
   (Sheet "preinscripciones", vía Apps Script) es la fuente de verdad para poder
   iniciar sesión desde otro dispositivo, y para recuperar la contraseña.
   euromodelo_users = { email: { password, mustChangePassword, nombre, tipoDocumento,
                                  numDocumento, ciudad, institucion, autorizacion, createdAt,
                                  inscripcion: null | {...} } }
   euromodelo_currentUser = "email" | null
   -------------------------------------------------- */
const AUTH_USERS_KEY = 'euromodelo_users';
const AUTH_CURRENT_KEY = 'euromodelo_currentUser';

function getUsers() {
  try {
    return JSON.parse(localStorage.getItem(AUTH_USERS_KEY)) || {};
  } catch (e) { return {}; }
}
function saveUsers(users) {
  localStorage.setItem(AUTH_USERS_KEY, JSON.stringify(users));
}
function createUser(data) {
  const users = getUsers();
  const username = (data.email || '').trim().toLowerCase();
  const password = randomAlphaNum(6);
  users[username] = { ...data, password, mustChangePassword: true, createdAt: new Date().toISOString(), inscripcion: null };
  saveUsers(users);
  return { username, password };
}
async function loginUser(usernameOrEmail, password) {
  const uname = (usernameOrEmail || '').trim().toLowerCase();
  const users = getUsers();
  const local = users[uname];
  if (local && local.password === password) {
    localStorage.setItem(AUTH_CURRENT_KEY, uname);
    return true;
  }
  // No coincide en este dispositivo (o es la primera vez aquí): verificar contra el backend.
  try {
    const res = await fetch(EUROMODELO_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ form: 'login', email: uname, password }),
    });
    const result = await res.json();
    if (result && result.ok && result.user) {
      users[uname] = {
        ...(local || {}),
        ...result.user,
        password,
        mustChangePassword: local ? !!local.mustChangePassword : false,
      };
      saveUsers(users);
      localStorage.setItem(AUTH_CURRENT_KEY, uname);
      return true;
    }
  } catch (e) {
    console.error('No se pudo verificar el usuario contra el backend:', e);
  }
  return false;
}
// Se intenta solo cuando loginUser ya falló: el mismo botón de "Iniciar sesión" de los
// participantes sirve también para el staff con cuenta de administrador. Si el correo/contraseña
// coincide con una cuenta admin (pestaña "admins" en la Sheet), guarda las credenciales bajo la
// misma clave de sessionStorage que ya usa admin.html, para que su auto-login la reconozca sin
// pedir la clave otra vez al llegar ahí.
const ADMIN_CREDS_STORAGE_KEY = 'euromodelo_admin_creds';
async function attemptAdminLogin(email, password) {
  try {
    const res = await fetch(EUROMODELO_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ form: 'admin_login', email, password }),
    });
    const result = await res.json();
    if (result && result.ok) {
      sessionStorage.setItem(ADMIN_CREDS_STORAGE_KEY, JSON.stringify({ email, password }));
      return true;
    }
  } catch (e) {
    console.error('No se pudo verificar como administrador:', e);
  }
  return false;
}
function logoutUser() {
  localStorage.removeItem(AUTH_CURRENT_KEY);
}
// A diferencia de loginUser, siempre consulta el backend (no usa el atajo local):
// sirve para refrescar el perfil con datos que solo vive en la Sheet, como el rol,
// la comisión o el partido ya asignados por el staff.
async function refreshUserFromServer(username) {
  const users = getUsers();
  const local = users[username];
  if (!local || !local.password) return null;
  try {
    const res = await fetch(EUROMODELO_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ form: 'login', email: username, password: local.password }),
    });
    const result = await res.json();
    if (result && result.ok && result.user) {
      users[username] = { ...local, ...result.user };
      saveUsers(users);
      return { username, ...users[username] };
    }
  } catch (e) {
    console.error('No se pudo actualizar el perfil desde el backend:', e);
  }
  return null;
}
function currentUsername() {
  return localStorage.getItem(AUTH_CURRENT_KEY);
}
function currentUserData() {
  const uname = currentUsername();
  if (!uname) return null;
  const users = getUsers();
  return users[uname] ? { username: uname, ...users[uname] } : null;
}
function saveInscripcion(username, inscripcionData) {
  const users = getUsers();
  if (users[username]) {
    users[username].inscripcion = inscripcionData;
    saveUsers(users);
  }
}
function changePassword(username, newPassword) {
  const users = getUsers();
  if (users[username]) {
    users[username].password = newPassword;
    users[username].mustChangePassword = false;
    saveUsers(users);
  }
  try {
    fetch(EUROMODELO_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ form: 'update_password', email: username, newPassword }),
    });
  } catch (e) {
    console.error('No se pudo sincronizar la nueva contraseña con el backend:', e);
  }
}
function requestPasswordReset(email) {
  try {
    fetch(EUROMODELO_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ form: 'forgot_password', email: (email || '').trim().toLowerCase() }),
    });
  } catch (e) {
    console.error('No se pudo solicitar la recuperación de contraseña:', e);
  }
}

/* ---------- Botones con estado de carga ----------
   Se usa en cualquier botón que dispare una espera real (fetch con respuesta leída, o una
   navegación a otra página): reemplaza el contenido del botón por un spinner + texto, y lo
   deshabilita para evitar doble click. clearButtonLoading lo devuelve a su estado original —
   no hace falta llamarla si la acción termina en una navegación (el botón se queda "cargando"
   hasta que la página siguiente reemplaza todo). -------------------------------------------- */
function setButtonLoading(btn, loadingText) {
  if (!btn || btn.dataset.loading === 'true') return;
  btn.dataset.loading = 'true';
  btn.dataset.originalHtml = btn.innerHTML;
  btn.disabled = true;
  btn.classList.add('btn-loading');
  btn.innerHTML = '<span class="btn-spinner"></span>' + (loadingText || 'Cargando…');
}
function clearButtonLoading(btn) {
  if (!btn || btn.dataset.loading !== 'true') return;
  btn.disabled = false;
  btn.classList.remove('btn-loading');
  if (btn.dataset.originalHtml !== undefined) btn.innerHTML = btn.dataset.originalHtml;
  delete btn.dataset.loading;
  delete btn.dataset.originalHtml;
}

/* ---------- Header: panel de inicio de sesión ---------- */
function initLoginPanel() {
  const loginBtn = document.getElementById('loginBtn');
  const loginPanel = document.getElementById('loginPanel');
  const loginForm = document.getElementById('loginForm');
  const authArea = document.getElementById('authArea');
  if (!authArea) return;

  const user = currentUserData();
  if (user) {
    authArea.innerHTML = `
      <div class="login-user-chip" id="userChip">
        <span>&#128100;</span><span>${user.nombre ? user.nombre.split(' ')[0] : user.username}</span>
      </div>
      <div class="login-panel" id="loginPanel">
        <h4>Hola, ${user.nombre || user.username}</h4>
        <p class="login-note">Sesión iniciada como <b>${user.username}</b>.</p>
        <a href="perfil.html" style="display:block; text-align:center; background:var(--blue-revel); color:#fff; padding:9px; border-radius:8px; font-weight:600; font-size:13.5px; margin-bottom:8px;">Ver mi perfil</a>
        <button id="logoutBtn">Cerrar sesión</button>
      </div>`;
    document.getElementById('userChip').addEventListener('click', () => {
      document.getElementById('loginPanel').classList.toggle('open');
    });
    document.getElementById('logoutBtn').addEventListener('click', () => {
      logoutUser();
      window.location.reload();
    });
    return;
  }

  if (loginBtn && loginPanel) {
    loginBtn.addEventListener('click', () => loginPanel.classList.toggle('open'));
  }
  if (loginForm) {
    loginForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const fd = new FormData(loginForm);
      const username = fd.get('username').trim();
      const password = fd.get('password').trim();
      const submitBtn = loginForm.querySelector('button[type="submit"]');
      if (submitBtn) setButtonLoading(submitBtn);
      const ok = await loginUser(username, password);
      if (ok) {
        window.location.href = 'perfil.html';
        return;
      }
      const isAdmin = await attemptAdminLogin(username, password);
      if (isAdmin) {
        window.location.href = 'admin.html';
        return;
      }
      if (submitBtn) clearButtonLoading(submitBtn);
      const errEl = document.getElementById('loginError');
      if (errEl) errEl.style.display = 'block';
    });
  }

  const forgotLink = document.getElementById('forgotPasswordLink');
  const forgotForm = document.getElementById('forgotPasswordForm');
  const forgotSent = document.getElementById('forgotPasswordSent');
  if (forgotLink && forgotForm) {
    forgotLink.addEventListener('click', (e) => {
      e.preventDefault();
      if (loginForm) loginForm.style.display = 'none';
      forgotLink.style.display = 'none';
      forgotForm.style.display = 'block';
    });
    forgotForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(forgotForm);
      requestPasswordReset(fd.get('email'));
      forgotForm.style.display = 'none';
      if (forgotSent) forgotSent.style.display = 'block';
    });
  }
}

/* ---------- Hemiciclo interactivo (SVG) ---------- */
function buildHemiciclo(containerId, parties, onDotClick) {
  const container = document.getElementById(containerId);
  if (!container) return;
  const total = parties.reduce((s, p) => s + p.seats, 0);
  const cx = 300, cy = 300;
  const rows = 9, innerR = 86, rowGap = 21;
  const angleStart = 180, angleEnd = 0;

  let cum = 0;
  const bounds = parties.map(p => {
    const startAngle = angleStart - (cum / total) * (angleStart - angleEnd);
    cum += p.seats;
    const endAngle = angleStart - (cum / total) * (angleStart - angleEnd);
    return { ...p, startAngle, endAngle };
  });
  function partyForAngle(angle) {
    for (const b of bounds) {
      if (angle <= b.startAngle + 0.01 && angle >= b.endAngle - 0.01) return b;
    }
    return bounds[bounds.length - 1];
  }

  const radii = [];
  let totalWeight = 0;
  for (let i = 0; i < rows; i++) { radii.push(innerR + i * rowGap); totalWeight += radii[i]; }
  const rowCounts = radii.map(r => Math.round(total * r / totalWeight));
  let diff = total - rowCounts.reduce((a, b) => a + b, 0);
  rowCounts[rowCounts.length - 1] += diff;

  let dotsSVG = '';
  radii.forEach((r, i) => {
    const n = rowCounts[i];
    for (let j = 0; j < n; j++) {
      const angle = n === 1 ? 90 : angleStart - j * ((angleStart - angleEnd) / (n - 1));
      const rad = angle * Math.PI / 180;
      const x = (cx + r * Math.cos(rad)).toFixed(1);
      const y = (cy - r * Math.sin(rad)).toFixed(1);
      const party = partyForAngle(angle);
      dotsSVG += `<circle cx="${x}" cy="${y}" r="5.3" fill="${party.color}" data-code="${party.code}"><title>${party.title} — ${party.seats} escaños</title></circle>`;
    }
  });

  container.innerHTML = `
    <svg viewBox="0 0 600 330" class="hemiciclo-svg">
      ${dotsSVG}
      <text x="300" y="298" text-anchor="middle" class="hemiciclo-total-n">${total}</text>
      <text x="300" y="316" text-anchor="middle" class="hemiciclo-total-l">escaños simulados en el Euromodelo</text>
    </svg>
    <div class="hemiciclo-legend">
      ${parties.map(p => `<div class="leg-chip" data-code="${p.code}"><span class="leg-dot" style="background:${p.color}"></span>${p.code} · ${p.seats}</div>`).join('')}
    </div>`;

  if (onDotClick) {
    container.querySelectorAll('circle, .leg-chip').forEach(el => {
      el.addEventListener('click', () => onDotClick(el.dataset.code));
    });
  }
}

/* ---------- Utilidades de aleatorización y desempate de brújulas ---------- */
function shuffleArray(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/* Elige la categoría ganadora de un tally; en caso de empate, al azar
   (nunca con un orden de prioridad fijo, para no sesgar el resultado
   agregado del cohorte hacia una categoría en particular). */
function pickQuizWinner(tally) {
  const max = Math.max(...Object.values(tally));
  const winners = Object.keys(tally).filter(k => tally[k] === max);
  return winners[Math.floor(Math.random() * winners.length)];
}

/* ---------- Quiz genérico (usado por roles.html, comisiones.html y partidos.html) ----------
   config = {
     modalId, emoji, title, intro,
     questions: [{q, options:[{label, value}]}],
     shuffleQuestions: bool (opcional),
     onFinish: (answers, body) => { ...renderResult... }
   }
   Las opciones de cada pregunta (y, si se pide, el orden de las preguntas)
   se mezclan de nuevo cada vez que el participante abre o reinicia el test.
   -------------------------------------------------------------------- */
function initQuiz(config) {
  let current = 0;
  let answers = [];
  let questions = [];
  const modal = document.getElementById(config.modalId);
  const body = modal.querySelector('.quiz-body');

  function buildQuestions() {
    let qs = config.questions.map(q => ({ q: q.q, options: shuffleArray(q.options) }));
    if (config.shuffleQuestions) qs = shuffleArray(qs);
    return qs;
  }

  function renderIntro() {
    body.innerHTML = `
      <div class="quiz-intro">
        ${config.emoji ? `<div class="q-emoji">${config.emoji}</div>` : ''}
        <p class="quiz-question">${config.title || ''}</p>
        <p style="font-size:13.5px; color:#54637C; line-height:1.6; margin:0 0 18px;">${config.intro}</p>
        <button class="btn-primary" id="quizStartBtn">Comenzar &rarr;</button>
      </div>`;
    document.getElementById('quizStartBtn').addEventListener('click', renderQuestion);
  }

  function renderQuestion() {
    const q = questions[current];
    const pct = Math.round((current / questions.length) * 100);
    body.innerHTML = `
      <div class="quiz-progress"><div class="quiz-progress-bar" style="width:${pct}%"></div></div>
      <p class="quiz-question">${current + 1}. ${q.q}</p>
      <div class="quiz-options">
        ${q.options.map((o, i) => `<button class="quiz-option" data-i="${i}">${o.label}</button>`).join('')}
      </div>
      <div class="quiz-nav">
        <button class="quiz-back" id="quizBack" ${current === 0 ? 'disabled' : ''}>&larr; Anterior</button>
        <span></span>
      </div>`;
    body.querySelectorAll('.quiz-option').forEach(btn => {
      btn.addEventListener('click', () => {
        answers[current] = q.options[+btn.dataset.i].value;
        current++;
        if (current >= questions.length) {
          renderResult();
        } else {
          renderQuestion();
        }
      });
    });
    const backBtn = document.getElementById('quizBack');
    if (backBtn) backBtn.addEventListener('click', () => { if (current > 0) { current--; renderQuestion(); } });
  }

  function renderResult() {
    config.onFinish(answers, body);
  }

  function restart() {
    current = 0;
    answers = [];
    questions = buildQuestions();
    if (config.intro) renderIntro(); else renderQuestion();
  }

  config._restart = restart;
  restart();
}
