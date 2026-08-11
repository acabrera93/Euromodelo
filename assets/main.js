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

/* ---------- Autenticación local (piloto) ----------
   Estructura en localStorage:
   euromodelo_users = { username: { password, nombre, tipoDocumento, numDocumento,
                                     ciudad, institucion, autorizacion, createdAt,
                                     inscripcion: null | {...} } }
   euromodelo_currentUser = "username" | null
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
  let username = slugifyName(data.nombre).slice(0, 6) + (data.numDocumento || '').slice(-3);
  if (!username || username.length < 4) username = 'euro' + randomAlphaNum(4).toLowerCase();
  let finalUsername = username;
  let n = 1;
  while (users[finalUsername]) { finalUsername = username + n; n++; }
  const password = randomAlphaNum(6);
  users[finalUsername] = { ...data, password, createdAt: new Date().toISOString(), inscripcion: null };
  saveUsers(users);
  return { username: finalUsername, password };
}
function loginUser(username, password) {
  const users = getUsers();
  const u = users[username];
  if (u && u.password === password) {
    localStorage.setItem(AUTH_CURRENT_KEY, username);
    return true;
  }
  return false;
}
function logoutUser() {
  localStorage.removeItem(AUTH_CURRENT_KEY);
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
    loginForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const fd = new FormData(loginForm);
      const ok = loginUser(fd.get('username').trim(), fd.get('password').trim());
      const errEl = document.getElementById('loginError');
      if (ok) {
        window.location.href = 'perfil.html';
      } else if (errEl) {
        errEl.style.display = 'block';
      }
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

/* ---------- Quiz genérico (usado por roles.html y partidos.html) ----------
   config = {
     modalId, title, emoji, questions: [{q, options:[{label, value}]}],
     onFinish: (tally) => { ...renderResult... }
   }
   -------------------------------------------------------------------- */
function initQuiz(config) {
  let current = 0;
  const answers = [];
  const modal = document.getElementById(config.modalId);
  const body = modal.querySelector('.quiz-body');

  function renderQuestion() {
    const q = config.questions[current];
    const pct = Math.round((current / config.questions.length) * 100);
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
        if (current >= config.questions.length) {
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
    answers.length = 0;
    renderQuestion();
  }

  config._restart = restart;
  restart();
}
