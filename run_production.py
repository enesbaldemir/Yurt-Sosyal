"""
ÜRETİM / TÜNEL KULLANIMI İÇİN ALTERNATİF BAŞLATICI
=====================================================

"python app.py" komutu Werkzeug'un geliştirme sunucusunu kullanır. Bu sunucu
küçük ölçekli testler için yeterlidir, ama 150-200 kişi gibi daha kalabalık ve
özellikle bir tünel (localtunnel, ngrok vb.) üzerinden erişim söz konusu
olduğunda zaman zaman yavaşlık ya da "Bad Gateway" hatalarına yol açabilir.

Bu betik, Flask-SocketIO'nun geliştirme sunucusu yerine eventlet'in kendi
WSGI sunucusunu kullanmasını sağlar. eventlet, yüzlerce eşzamanlı bağlantıyı
(sohbet, bildirim, DM gibi WebSocket bağlantıları dahil) tek bir işlemde,
ağır işletim sistemi thread'leri açmadan, çok daha verimli şekilde yönetir.
Bu hem hızı hem de kararlılığı belirgin şekilde artırır.

KURULUM (bir kere):
    pip install eventlet

ÇALIŞTIRMA:
    python run_production.py

Ardından tünelinizi her zamanki gibi aynı porta (varsayılan 5000) açabilirsiniz:
    npx localtunnel --port 5000

İPUCU: Gerçek 150-200 kişilik kullanım öncesi bir gece önceden bu şekilde
çalıştırıp birkaç kişiyle test etmenizi öneririz. Herhangi bir sorunla
karşılaşırsanız her zaman güvenli yedek seçeneğiniz "python app.py"dir
(threading modu) — hiçbir veri kaybı olmadan geri dönebilirsiniz.

NOT: Tünel servisleri (özellikle ücretsiz localtunnel) kendi başlarına da
zaman zaman yavaş/kararsız olabilir; bu betik uygulamanın kendi tarafındaki
darboğazı çözer ama tünel sağlayıcısının kendi limitlerini değiştiremez.
150-200 kişilik gerçek/kritik bir kullanım için mümkünse ufak bir bulut
sunucusuna (ör. bir VPS + kendi alan adınız) kalıcı olarak kurmanız çok
daha sağlıklı olacaktır.
"""
import os

# eventlet.monkey_patch() diğer TÜM importlardan önce çalışmalıdır, bu yüzden
# app.py'yi (ve onun sqlite3/threading gibi standart kütüphane importlarını)
# içeri almadan önce burada patch'i uyguluyoruz.
os.environ.setdefault('YURT_ASYNC_MODE', 'eventlet')

try:
    import eventlet
except ImportError:
    raise SystemExit(
        "eventlet kurulu değil. Önce şunu çalıştırın:\n\n    pip install eventlet\n"
    )

eventlet.monkey_patch()

from app import app, socketio  # noqa: E402  (patch'ten SONRA import edilmeli)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Yurt Sosyal (eventlet modu) http://0.0.0.0:{port} adresinde başlatılıyor...")
    socketio.run(app, host='0.0.0.0', port=port, debug=False, use_reloader=False)
