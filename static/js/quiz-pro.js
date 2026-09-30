/* =====================================================================
   QUIZ PRO — ortak yardımcılar (window.KQ)
   Şekiller, animasyonlu sıralama (FLIP), sayı sayacı, konfeti, sesler.
   ===================================================================== */
(function (window, document) {
    'use strict';

    var reduced = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    // ---------- Şekiller: A üçgen, B elmas, C daire, D kare ----------
    var SHAPES = {
        a: '<svg viewBox="0 0 32 32" aria-hidden="true"><path d="M16 3.5 30 28.5H2z"/></svg>',
        b: '<svg viewBox="0 0 32 32" aria-hidden="true"><path d="M16 1.5 30.5 16 16 30.5 1.5 16z"/></svg>',
        c: '<svg viewBox="0 0 32 32" aria-hidden="true"><circle cx="16" cy="16" r="14"/></svg>',
        d: '<svg viewBox="0 0 32 32" aria-hidden="true"><rect x="3" y="3" width="26" height="26" rx="2"/></svg>'
    };
    var SHAPE_NAMES = { a: 'Üçgen', b: 'Elmas', c: 'Daire', d: 'Kare' };
    function shape(letter) { return '<span class="kq-shape">' + (SHAPES[letter] || '') + '</span>'; }

    function esc(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
            return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
        });
    }

    // ---------- Toast ----------
    var toastEl = null, toastTimer = null;
    function toast(msg, isError) {
        if (!toastEl) {
            toastEl = document.createElement('div');
            toastEl.className = 'kq-toast';
            toastEl.setAttribute('role', 'status');
            document.body.appendChild(toastEl);
        }
        toastEl.textContent = msg;
        toastEl.classList.toggle('err', !!isError);
        toastEl.classList.add('show');
        clearTimeout(toastTimer);
        toastTimer = setTimeout(function () { toastEl.classList.remove('show'); }, 2600);
    }

    // ---------- Sayı sayacı ----------
    function countUp(el, from, to, ms) {
        if (!el) return;
        from = Number(from) || 0; to = Number(to) || 0;
        if (reduced || from === to) { el.textContent = to.toLocaleString('tr-TR'); return; }
        var start = null, dur = ms || 900;
        function step(ts) {
            if (!start) start = ts;
            var p = Math.min(1, (ts - start) / dur), e = 1 - Math.pow(1 - p, 3);
            el.textContent = Math.round(from + (to - from) * e).toLocaleString('tr-TR');
            if (p < 1) requestAnimationFrame(step);
        }
        requestAnimationFrame(step);
    }

    // ---------- Animasyonlu sıralama (FLIP) ----------
    // rows: [{id, nickname, score, prev_score, gained, rank, prev_rank, streak}]
    // Satırlar ÖNCEKİ sıralarında ve eski puanlarıyla çizilir, sonra yeni yerlerine kayar.
    function renderBoard(container, rows, opts) {
        opts = opts || {};
        var limit = opts.limit || 5, rowH = opts.rowHeight || (window.innerWidth < 768 ? 62 : 74);
        var meId = opts.meId != null ? String(opts.meId) : null;
        container.innerHTML = '';
        container.classList.remove('is-settled');
        container.classList.add('kq-lb');

        // Yeni sırada ilk N + (varsa) önceki sıralamada ilk N'de olup düşenler
        var shown = rows.filter(function (r) { return r.rank <= limit; });
        var byPrev = rows.slice().sort(function (a, b) { return a.prev_rank - b.prev_rank; });
        container.style.height = (Math.min(limit, shown.length) * rowH) + 'px';

        var els = shown.map(function (r) {
            var el = document.createElement('div');
            el.className = 'kq-lb-row' + (r.rank === 1 ? ' is-first' : '') + (meId && String(r.id) === meId ? ' is-me' : '');
            var delta = (r.prev_rank || r.rank) - r.rank;
            el.innerHTML =
                '<span class="kq-lb-rank">' + r.rank + '</span>' +
                '<span class="kq-lb-name">' + esc(r.nickname) + '</span>' +
                (r.streak >= 2 ? '<span class="kq-lb-streak">🔥 ' + r.streak + '</span>' : '') +
                (r.gained ? '<span class="kq-lb-gain">+' + r.gained + '</span>' : '') +
                (delta ? '<span class="kq-lb-delta ' + (delta > 0 ? 'up' : 'down') + '">' + (delta > 0 ? '▲ ' : '▼ ') + Math.abs(delta) + '</span>' : '') +
                '<span class="kq-lb-score">' + Number(r.prev_score != null ? r.prev_score : r.score).toLocaleString('tr-TR') + '</span>';
            // Başlangıç konumu: önceki sıra (ilk N dışından geliyorsa alttan kayarak girer)
            var startIdx = Math.min((r.prev_rank || r.rank) - 1, limit);
            el.style.transform = 'translateY(' + (startIdx * rowH) + 'px)';
            if (startIdx >= limit) el.style.opacity = '0';
            container.appendChild(el);
            return { el: el, r: r };
        });

        var delay = reduced ? 0 : (opts.delay != null ? opts.delay : 650);
        setTimeout(function () {
            els.forEach(function (x) {
                x.el.style.opacity = '1';
                x.el.style.transition = 'transform .9s cubic-bezier(.2,.9,.25,1.1), opacity .5s';
                x.el.style.transform = 'translateY(' + ((x.r.rank - 1) * rowH) + 'px)';
                countUp(x.el.querySelector('.kq-lb-score'), x.r.prev_score != null ? x.r.prev_score : x.r.score, x.r.score, 1100);
            });
            setTimeout(function () { container.classList.add('is-settled'); }, 350);
        }, delay);
        return byPrev;
    }

    // ---------- Konfeti ----------
    var cv = null, ctx = null, parts = [], raf = null;
    var COLORS = ['#E21B3C', '#1368CE', '#D89E00', '#26890C', '#FFFFFF', '#8B4DFF'];
    function confetti(count, originY) {
        if (reduced) return;
        if (!cv) {
            cv = document.createElement('canvas');
            cv.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:3000;';
            document.body.appendChild(cv);
            ctx = cv.getContext('2d');
            var rs = function () { cv.width = innerWidth; cv.height = innerHeight; };
            rs(); addEventListener('resize', rs);
        }
        count = Math.min(count || 120, 260);
        for (var i = 0; i < count; i++) {
            parts.push({
                x: cv.width * (0.1 + Math.random() * 0.8), y: originY != null ? originY : -20 - Math.random() * 80,
                vx: (Math.random() - 0.5) * 7, vy: originY != null ? -6 - Math.random() * 8 : 2 + Math.random() * 4,
                s: 6 + Math.random() * 7, c: COLORS[i % COLORS.length], r: Math.random() * 6, vr: (Math.random() - 0.5) * 0.35, life: 0,
                max: 140 + Math.random() * 60
            });
        }
        if (!raf) tick();
    }
    function tick() {
        ctx.clearRect(0, 0, cv.width, cv.height);
        parts = parts.filter(function (p) {
            p.x += p.vx; p.y += p.vy; p.vy += 0.15; p.vx *= 0.995; p.r += p.vr; p.life++;
            ctx.save();
            ctx.globalAlpha = Math.max(0, 1 - Math.max(0, p.life - p.max * 0.7) / (p.max * 0.3));
            ctx.translate(p.x, p.y); ctx.rotate(p.r); ctx.fillStyle = p.c;
            ctx.fillRect(-p.s / 2, -p.s / 2, p.s, p.s * 0.55);
            ctx.restore();
            return p.life < p.max && p.y < cv.height + 40;
        });
        raf = parts.length ? requestAnimationFrame(tick) : null;
    }

    // ---------- Küçük sesler (WebAudio, dosya gerektirmez) ----------
    var audioCtx = null, muted = false;
    try { muted = localStorage.getItem('kq-muted') === '1'; } catch (e) {}
    function ac() {
        if (!audioCtx) {
            var C = window.AudioContext || window.webkitAudioContext;
            if (!C) return null;
            audioCtx = new C();
        }
        if (audioCtx.state === 'suspended') audioCtx.resume();
        return audioCtx;
    }
    function tone(freq, dur, type, vol, when) {
        if (muted) return;
        var a = ac(); if (!a) return;
        var t = a.currentTime + (when || 0);
        var o = a.createOscillator(), g = a.createGain();
        o.type = type || 'sine'; o.frequency.setValueAtTime(freq, t);
        g.gain.setValueAtTime(0.0001, t);
        g.gain.exponentialRampToValueAtTime(vol || 0.12, t + 0.01);
        g.gain.exponentialRampToValueAtTime(0.0001, t + dur);
        o.connect(g); g.connect(a.destination); o.start(t); o.stop(t + dur + 0.02);
    }
    var sfx = {
        tick: function () { tone(880, 0.07, 'square', 0.05); },
        pop: function () { tone(520, 0.08, 'triangle', 0.1); tone(780, 0.1, 'triangle', 0.08, 0.05); },
        correct: function () { tone(660, 0.12, 'triangle', 0.14); tone(880, 0.14, 'triangle', 0.14, 0.1); tone(1320, 0.25, 'triangle', 0.12, 0.2); },
        wrong: function () { tone(300, 0.2, 'sawtooth', 0.08); tone(200, 0.3, 'sawtooth', 0.08, 0.15); },
        reveal: function () { tone(392, 0.15, 'triangle', 0.12); tone(523, 0.15, 'triangle', 0.12, 0.12); tone(659, 0.3, 'triangle', 0.12, 0.24); },
        fanfare: function () { [523, 659, 784, 1047].forEach(function (f, i) { tone(f, 0.28, 'triangle', 0.13, i * 0.14); }); tone(1047, 0.8, 'triangle', 0.1, 0.6); },
        drum: function () { for (var i = 0; i < 8; i++) tone(120 + i * 6, 0.09, 'square', 0.05, i * 0.1); }
    };
    function setMuted(v) { muted = !!v; try { localStorage.setItem('kq-muted', muted ? '1' : '0'); } catch (e) {} }
    function isMuted() { return muted; }

    // ---------- Soru kutucukları ----------
    function tileHtml(letter, text, img, extra) {
        return '<span class="kq-shape">' + SHAPES[letter] + '</span>' +
            (img ? '<img src="/static/uploads/quiz/' + encodeURIComponent(img) + '" alt="">' : '') +
            '<span class="txt">' + esc(text) + '</span>' + (extra || '') + '<span class="mark"><i class="bi bi-check-lg"></i></span>';
    }
    function filledLetters(q) {
        return ['a', 'b', 'c', 'd'].filter(function (k) { return (q['option_' + k] || '').trim() || q['option_' + k + '_image']; });
    }

    // ---------- Geri sayım halkası ----------
    function ringTimer(el, totalMs, remainingMs, onTick, onDone) {
        var circle = el.querySelector('circle'), num = el.querySelector('.num') || el;
        var r = circle ? circle.r.baseVal.value : 0, C = 2 * Math.PI * r;
        if (circle) { circle.style.strokeDasharray = C; }
        var end = Date.now() + remainingMs, lastSec = null, stopped = false;
        function frame() {
            if (stopped) return;
            var left = Math.max(0, end - Date.now()), sec = Math.ceil(left / 1000);
            if (sec !== lastSec) {
                lastSec = sec;
                num.textContent = sec;
                el.classList.toggle('urgent', sec <= 5 && sec > 0);
                if (onTick) onTick(sec);
                if (circle) circle.style.strokeDashoffset = C * (1 - (left / totalMs));
            }
            if (left <= 0) { if (onDone) onDone(); return; }
            requestAnimationFrame(frame);
        }
        frame();
        return { stop: function () { stopped = true; } };
    }

    function progressBar(el, ms) {
        var bar = el.querySelector('i'); if (!bar) return;
        bar.style.transition = 'none'; bar.style.transform = 'scaleX(1)';
        void bar.offsetWidth;
        bar.style.transition = 'transform ' + (ms / 1000) + 's linear';
        bar.style.transform = 'scaleX(0)';
    }

    function showScreen(root, id) {
        root.querySelectorAll('[data-screen]').forEach(function (s) { s.classList.toggle('is-active', s.dataset.screen === id); });
    }

    function connectionBadge(socket) {
        var el = document.createElement('div');
        el.className = 'kq-toast';
        el.style.bottom = 'auto'; el.style.top = '72px';
        document.body.appendChild(el);
        var wasDown = false;
        socket.on('disconnect', function () { wasDown = true; el.textContent = 'Bağlantı koptu, yeniden bağlanılıyor…'; el.classList.add('show', 'err'); });
        socket.on('connect', function () {
            if (!wasDown) return;
            el.textContent = 'Bağlantı geri geldi ✓'; el.classList.remove('err');
            setTimeout(function () { el.classList.remove('show'); }, 1800);
        });
    }

    window.KQ = {
        SHAPES: SHAPES, SHAPE_NAMES: SHAPE_NAMES, shape: shape, esc: esc, toast: toast, countUp: countUp,
        renderBoard: renderBoard, confetti: confetti, sfx: sfx, setMuted: setMuted, isMuted: isMuted,
        tileHtml: tileHtml, filledLetters: filledLetters, ringTimer: ringTimer, progressBar: progressBar,
        showScreen: showScreen, connectionBadge: connectionBadge, reduced: reduced
    };
})(window, document);
