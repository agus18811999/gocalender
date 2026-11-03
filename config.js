/* CALENDRA — shared, public browser configuration. Do not put secrets here. */
(() => {
  'use strict';
  const CONFIG = Object.freeze({
    APP_NAME: 'Calendra',
    GAS_WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzoBrI8acvtl2WmQ4HyLX10_OegPkrc8PW7VOXZ1z4aTJkKnV6UQVNAj73S65k1LxvM/exec',
    GOOGLE_CLIENT_ID: '677539686091-le0onnuqtuf04bpt5jebng819rivqlff.apps.googleusercontent.com',
    FIREBASE_CONFIG: Object.freeze({
      apiKey: 'AIzaSyAk1VaR81jMGwNHRkZ0_kTJeNWdJxCHANc',
      authDomain: 'calendra-301e2.firebaseapp.com',
      projectId: 'calendra-301e2',
      appId: '1:677539686091:web:2f369ad95f0c71a83add4e'
    }),
    CALENDAR_API_BASE: 'https://www.googleapis.com/calendar/v3',
    CALENDAR_SCOPES: [
      'https://www.googleapis.com/auth/calendar.events',
      'https://www.googleapis.com/auth/calendar.calendarlist.readonly',
      'https://www.googleapis.com/auth/calendar.freebusy'
    ],
    CALENDAR_ADMIN_SCOPE: 'https://www.googleapis.com/auth/calendar.calendars',
    LOCALE: 'id-ID',
    DEFAULT_TIMEZONE: 'Asia/Jakarta',
    RPC_CHANNEL: 'CALENDRA_METADATA_RPC_V1'
  });
  const configured = () => !/YOUR_|\.\.\./.test(CONFIG.FIREBASE_CONFIG.apiKey) && !/YOUR_/.test(CONFIG.FIREBASE_CONFIG.projectId);
  const gasConfigured = () => /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec(?:\?.*)?$/.test(CONFIG.GAS_WEB_APP_URL);
  const googleConfigured = () => /^\d+-[a-zA-Z0-9_-]+\.apps\.googleusercontent\.com$/.test(CONFIG.GOOGLE_CLIENT_ID);
  const $ = (selector, root = document) => root.querySelector(selector);
  const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];
  const escapeHTML = (value = '') => String(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmtDate = (input, opt = {day:'numeric',month:'long',year:'numeric'}) => new Intl.DateTimeFormat(CONFIG.LOCALE,opt).format(new Date(input));
  const fmtTime = input => new Intl.DateTimeFormat(CONFIG.LOCALE,{hour:'2-digit',minute:'2-digit'}).format(new Date(input));
  const uid = () => crypto.randomUUID?.() || 'id' + Date.now().toString(36) + Math.random().toString(36).slice(2);
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const errorMessage = e => ({'auth/invalid-credential':'Email atau password tidak cocok.','auth/email-already-in-use':'Email sudah digunakan.','auth/weak-password':'Password terlalu lemah.','auth/popup-closed-by-user':'Login Google dibatalkan.','auth/too-many-requests':'Terlalu banyak percobaan. Coba lagi nanti.','auth/network-request-failed':'Koneksi terganggu. Periksa internet.'}[e?.code] || e?.message || 'Terjadi kesalahan. Coba lagi.');

  /* HtmlService bridge: no Firebase tokens in URLs, JSONP, or localStorage.
     iframe support is browser/GAS-deployment-dependent; failures are reported, never faked. */
  function makeMetadataBridge() {
    let frame, endpoint, alive = false, initializing, seq = 0;
    const nonce = uid() + uid();
    const pending = new Map();
    function deliver(event) {
      const msg = event.data;
      if (!frame || !msg || msg.channel !== CONFIG.RPC_CHANNEL || msg.nonce !== nonce) return;
      if (event.origin !== 'null' && !/^https:\/\/(?:[a-z0-9-]+\.)*googleusercontent\.com$/.test(event.origin) && event.origin !== 'https://script.google.com') return;
      if (msg.type === 'ready') {
        endpoint = event.source;
        alive = true;
        const task = pending.get('init');
        if (task) { clearTimeout(task.timer); pending.delete('init'); task.resolve(true); }
      } else if (msg.type === 'response') {
        const task = pending.get(msg.id);
        if (task) { clearTimeout(task.timer); pending.delete(msg.id); (msg.result?.success ? task.resolve : task.reject)(msg.result?.success ? msg.result.data : Error(msg.result?.error?.message || 'Metadata backend gagal.')); }
      }
    }
    async function init() {
      if (alive) return true;
      if (initializing) return initializing;
      if (!gasConfigured()) throw Error('URL Google Apps Script belum diatur di config.js.');
      initializing = new Promise((resolve,reject) => {
        frame = document.createElement('iframe');
        frame.style.cssText='position:absolute;width:1px;height:1px;border:0;opacity:0;pointer-events:none;';
        frame.title='Jalur aman metadata';
        frame.setAttribute('aria-hidden','true');
        const timer = setTimeout(() => {pending.delete('init'); reject(Error('Bridge Google Apps Script tidak dapat dimuat. Periksa izin iframe, deployment, atau kebijakan browser.'));}, 11000);
        pending.set('init',{resolve,reject,timer});
        window.addEventListener('message',deliver);
        frame.addEventListener('load', () => { try {frame.contentWindow.postMessage({channel:CONFIG.RPC_CHANNEL,type:'ping',nonce},'*');} catch {} });
        frame.src = CONFIG.GAS_WEB_APP_URL + (CONFIG.GAS_WEB_APP_URL.includes('?')?'&':'?') + 'bridge=1&nonce=' + encodeURIComponent(nonce);
        document.body.append(frame);
      });
      try { return await initializing; } finally { initializing = null; }
    }
    async function call(action, payload, idToken) {
      await init();
      if (!idToken) throw Error('Sesi Firebase tidak valid.');
      const id = 'q' + (++seq);
      return new Promise((resolve,reject) => {
        const timer = setTimeout(() => {pending.delete(id);reject(Error('Backend tidak merespons dalam 18 detik.'));}, 18000);
        pending.set(id,{resolve,reject,timer});
        endpoint.postMessage({channel:CONFIG.RPC_CHANNEL,type:'request',nonce,id,action,payload:payload||{},idToken},'*');
      });
    }
    function dispose() {
      for (const [,p] of pending) { clearTimeout(p.timer); p.reject?.(Error('Bridge ditutup.')); }
      pending.clear(); frame?.remove(); frame = null; endpoint = null; alive = false; window.removeEventListener('message',deliver);
    }
    return {init,call,dispose,get connected(){return alive;}};
  }
  window.CALENDRA = Object.freeze({CONFIG,configured,gasConfigured,googleConfigured,$,$$,escapeHTML,fmtDate,fmtTime,uid,sleep,errorMessage,makeMetadataBridge});
})();
