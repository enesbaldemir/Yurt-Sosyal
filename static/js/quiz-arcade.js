/* =====================================================================
   QUIZ ARCADE — paylaşılan JS yardımcıları
   window.QuizArcade altında dışa açılır; sayfa scriptleri bunu kullanır.
   ===================================================================== */
(function (window) {
    'use strict';

    function prefersReducedMotion() {
        return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    }

    // ---------- Ripple efekti (tıklanan butonun içine) ----------
    function ripple(el, evt) {
        if (prefersReducedMotion() || !el) return;
        var rect = el.getBoundingClientRect();
        var size = Math.max(rect.width, rect.height);
        var x = (evt && evt.clientX != null) ? evt.clientX - rect.left : rect.width / 2;
        var y = (evt && evt.clientY != null) ? evt.clientY - rect.top : rect.height / 2;
        var span = document.createElement('span');
        span.className = 'qa-ripple';
        span.style.width = span.style.height = size + 'px';
        span.style.left = (x - size / 2) + 'px';
        span.style.top = (y - size / 2) + 'px';
        el.appendChild(span);
        setTimeout(function () { span.remove(); }, 600);
    }

    // ---------- Sayı count-up animasyonu ----------
    function countUp(el, target, opts) {
        if (!el) return;
        opts = opts || {};
        var duration = prefersReducedMotion() ? 0 : (opts.duration || 900);
        var start = opts.from != null ? opts.from : 0;
        target = Number(target) || 0;
        if (duration <= 0) { el.textContent = (opts.prefix || '') + target + (opts.suffix || ''); return; }
        var startTime = null;
        function step(ts) {
            if (!startTime) startTime = ts;
            var p = Math.min(1, (ts - startTime) / duration);
            var eased = 1 - Math.pow(1 - p, 3);
            var val = Math.round(start + (target - start) * eased);
            el.textContent = (opts.prefix || '') + val + (opts.suffix || '');
            if (p < 1) requestAnimationFrame(step);
        }
        requestAnimationFrame(step);
    }

    // ---------- Uçan puan etiketi (+850 gibi) ----------
    function floatScore(originEl, text, color) {
        if (!originEl) return;
        var rect = originEl.getBoundingClientRect();
        var el = document.createElement('div');
        el.className = 'qa-floater';
        el.textContent = text;
        if (color) el.style.color = color;
        el.style.left = (rect.left + rect.width / 2 - 30) + 'px';
        el.style.top = (rect.top - 10) + 'px';
        document.body.appendChild(el);
        setTimeout(function () { el.remove(); }, 1150);
    }

    // ---------- Hafif canvas confetti (kütüphanesiz) ----------
    var confettiCanvas = null, confettiCtx = null, confettiParticles = [], confettiRaf = null;
    var CONFETTI_COLORS = ['#7C3AED', '#2563EB', '#EC4899', '#F97316', '#FACC15', '#22C55E'];

    function ensureCanvas() {
        if (confettiCanvas) return confettiCanvas;
        confettiCanvas = document.createElement('canvas');
        confettiCanvas.className = 'qa-confetti-canvas';
        document.body.appendChild(confettiCanvas);
        confettiCtx = confettiCanvas.getContext('2d');
        function resize() {
            confettiCanvas.width = window.innerWidth;
            confettiCanvas.height = window.innerHeight;
        }
        resize();
        window.addEventListener('resize', resize);
        return confettiCanvas;
    }

    function confetti(opts) {
        if (prefersReducedMotion()) return;
        opts = opts || {};
        var count = Math.min(opts.count || 70, 140); // mobilde de performansı korumak için sınırlı
        ensureCanvas();
        var w = confettiCanvas.width;
        for (var i = 0; i < count; i++) {
            confettiParticles.push({
                x: w / 2 + (Math.random() - 0.5) * w * 0.5,
                y: -10 - Math.random() * 40,
                vx: (Math.random() - 0.5) * 6,
                vy: 2 + Math.random() * 4,
                size: 5 + Math.random() * 5,
                color: CONFETTI_COLORS[i % CONFETTI_COLORS.length],
                rot: Math.random() * Math.PI,
                vr: (Math.random() - 0.5) * 0.3,
                life: 0,
                maxLife: 90 + Math.random() * 40
            });
        }
        if (!confettiRaf) tickConfetti();
    }

    function tickConfetti() {
        if (!confettiCtx) { confettiRaf = null; return; }
        confettiCtx.clearRect(0, 0, confettiCanvas.width, confettiCanvas.height);
        confettiParticles = confettiParticles.filter(function (p) {
            p.x += p.vx; p.y += p.vy; p.vy += 0.12; p.rot += p.vr; p.life++;
            var alpha = 1 - Math.max(0, (p.life - p.maxLife * 0.7) / (p.maxLife * 0.3));
            confettiCtx.save();
            confettiCtx.globalAlpha = Math.max(0, alpha);
            confettiCtx.translate(p.x, p.y);
            confettiCtx.rotate(p.rot);
            confettiCtx.fillStyle = p.color;
            confettiCtx.fillRect(-p.size / 2, -p.size / 2, p.size, p.size * 0.6);
            confettiCtx.restore();
            return p.life < p.maxLife && p.y < confettiCanvas.height + 30;
        });
        if (confettiParticles.length) {
            confettiRaf = requestAnimationFrame(tickConfetti);
        } else {
            confettiRaf = null;
        }
    }

    // ---------- Bağlantı durumu rozeti (Socket.IO) ----------
    function attachConnectionBanner(socket) {
        if (!socket) return;
        var el = document.createElement('div');
        el.className = 'qa-connection';
        el.innerHTML = '<span class="spin"></span><span class="qa-connection-text">Bağlantı kuruluyor…</span>';
        document.body.appendChild(el);
        var textEl = el.querySelector('.qa-connection-text');
        var hideTimer = null;

        function show(text, ok, autoHide) {
            clearTimeout(hideTimer);
            textEl.textContent = text;
            el.classList.toggle('ok', !!ok);
            el.querySelector('.spin').style.display = ok ? 'none' : 'inline-block';
            el.classList.add('show');
            if (autoHide) hideTimer = setTimeout(function () { el.classList.remove('show'); }, 2200);
        }

        socket.on('disconnect', function () { show('Bağlantı yenileniyor…', false, false); });
        socket.on('reconnect_attempt', function () { show('Bağlantı yenileniyor…', false, false); });
        socket.on('connect', function () {
            // İlk bağlantıda rozet göstermeye gerek yok, sadece kopup geri gelince bilgi ver.
            if (el.dataset.everDisconnected === '1') show('Bağlantı sağlandı', true, true);
        });
        socket.on('disconnect', function () { el.dataset.everDisconnected = '1'; });
        return el;
    }

    // ---------- Basit toast ----------
    function toast(msg) {
        var el = document.createElement('div');
        el.className = 'qa-connection ok show';
        el.style.background = 'var(--qa-purple)';
        el.textContent = msg;
        document.body.appendChild(el);
        setTimeout(function () { el.remove(); }, 2000);
    }

    // ---------- Timer ring güncelleme ----------
    function setRing(ringEl, remaining, total) {
        if (!ringEl) return;
        var fg = ringEl.querySelector('.qa-ring-fg');
        var num = ringEl.querySelector('.qa-ring-num');
        if (!fg) return;
        var r = fg.r.baseVal.value;
        var circumference = 2 * Math.PI * r;
        var ratio = total > 0 ? Math.max(0, remaining / total) : 0;
        fg.style.strokeDasharray = circumference;
        fg.style.strokeDashoffset = circumference * (1 - ratio);
        var color = '#22C55E';
        if (ratio <= 0.2) color = '#EF4444';
        else if (ratio <= 0.4) color = '#F97316';
        else if (ratio <= 0.65) color = '#FACC15';
        fg.style.stroke = color;
        if (num) num.textContent = Math.max(0, Math.ceil(remaining));
        ringEl.classList.toggle('qa-urgent', remaining <= 5 && remaining > 0);
    }

    window.QuizArcade = {
        ripple: ripple,
        countUp: countUp,
        floatScore: floatScore,
        confetti: confetti,
        attachConnectionBanner: attachConnectionBanner,
        toast: toast,
        setRing: setRing,
        prefersReducedMotion: prefersReducedMotion
    };
})(window);
