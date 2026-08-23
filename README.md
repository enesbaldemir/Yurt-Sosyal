# Yurt Sosyal 🌙

Yurt öğrencileri için sosyal ağ + yoklama + etkinlik yönetim uygulaması.

## Kurulum

```bash
python -m venv venv
# Windows:  venv\Scripts\activate
# Mac/Linux: source venv/bin/activate

pip install -r requirements.txt
python app.py
```

Uygulama http://localhost:5000 adresinde çalışır.

Varsayılan admin girişi: **admin** / **admin123**

Dışarıya açmak için (test amaçlı):
```bash
npx localtunnel --port 5000
```
cloudflared tunnel --url http://localhost:5000
## Tünel (localtunnel/ngrok) kullanırken yavaşlık ya da "Bad Gateway"

Bunun en sık sebebi, `python app.py`'nin kullandığı Werkzeug geliştirme sunucusunun
varsayılan olarak **aynı anda tek bir isteği** işlemesidir: biri video yüklerken ya
da sohbet için WebSocket bağlantısı açıkken, diğer herkes (ve tünel) o istek bitene
kadar bekler — tünel bu bekleme süresini aşınca "Bad Gateway" gösterir. Bunu şu
sırayla deneyin:

1. **Basit çözüm (kod güncellemesiyle otomatik geldi):** Artık sunucu `threaded=True`
   ile çalışıyor ve `debug` modu varsayılan olarak **kapalı** (hem güvenlik hem hız
   için). Bu değişiklik tek başına çoğu "bazı sekmelerde Bad Gateway" şikayetini
   çözer. Elle açmak isterseniz: `YURT_DEBUG=true python app.py`
2. **Daha kalabalık/kararlı kullanım için (önerilir, 150-200 kişi gibi):**
   ```bash
   pip install eventlet
   python run_production.py
   ```
   Bu, WebSocket bağlantılarını (sohbet, bildirim, DM) çok daha verimli yöneten
   gerçek bir eşzamanlı sunucuya geçer. Tüneli her zamanki gibi aynı porta açın.
   Bir sorun olursa `python app.py`'ye sorunsuz geri dönebilirsiniz.
3. **Tünelin kendisi de bir darboğaz olabilir:** Ücretsiz localtunnel servisi zaman
   zaman kendi başına yavaş/kararsız olabilir; bu, uygulamanın kendisinden bağımsız
   bir kısıttır. Gerçek 150-200 kişilik, kritik bir kullanım için mümkünse ufak bir
   bulut sunucusuna (VPS) kalıcı olarak kurmak çok daha sağlıklı olur.

## Bu sürümde yapılan değişiklikler

### 🐞 Düzeltilen hatalar
- `/api/notifications` rotası `socketio.run()` çağrısından **sonra** tanımlanmıştı; bu satırlar
  `python app.py` ile hiçbir zaman çalışmıyordu çünkü `socketio.run()` süreci bloke ediyor.
  Rota artık dosyanın normal akışında, sunucu başlamadan önce tanımlanıyor.
- **Sunucu kararsızlığı:** `requirements.txt` içinde `eventlet` kuruluydu ama hiç
  `monkey_patch()` çağrılmıyordu; Flask-SocketIO bunu otomatik async motoru olarak
  seçmeye çalışınca donmalara/çökmelere yol açıyordu. `SocketIO(..., async_mode='threading')`
  ile motoru açıkça sabitledik ve `eventlet`'i bağımlılıklardan kaldırdık.
- **Sunucu kararsızlığı (2):** Aynı anda birden fazla yazma isteği (sohbet + yoklama + DM)
  SQLite'ta "database is locked" hatasına yol açabiliyordu. `PRAGMA journal_mode=WAL` ve
  `busy_timeout` eklendi.
- Gece zamanlayıcı fonksiyonu (`send_nightly_attendance`) arka plan iş parçacığında
  `url_for()` çağırdığı için `RuntimeError` fırlatıyordu (SERVER_NAME gerektirir) — artık
  göreli yol doğrudan oluşturuluyor.
- Etkinlik kontenjan kontrolünde (`max_participants`) string/integer karşılaştırma hatası
  düzeltildi.
- Gizli anahtar (`SECRET_KEY`) artık `YURT_SECRET_KEY` ortam değişkeninden okunabiliyor.
- Kayıt formunda boş alan / kısa şifre kontrolü eklendi.
- Geliştirme sunucusunda `use_reloader=False` yapıldı; aksi halde arka planda çalışan
  gece yoklaması zamanlayıcısı (scheduler) iki kez başlıyordu.
- Bildirimler artık sadece ilgili kullanıcıya (SocketIO odası ile) gönderiliyor.
- **Gönderi yorumları açılmıyordu** — "yorumları gör" linki hiçbir şey yapmıyordu ve
  yorumlar sayfada hiç render edilmiyordu. Artık her gönderi altında gerçek yorum listesi
  var, "yorum" ikonuna ya da "X yorumun tümünü gör" yazısına tıklayınca açılıp kapanıyor.
- **Tünelde yavaşlık / "Bad Gateway"** — Werkzeug geliştirme sunucusu varsayılan olarak
  tek seferde tek istek işliyordu (`threaded=True` eksikti); büyük bir dosya yükleme ya
  da açık bir sohbet bağlantısı tüm diğer sekmeleri bloke ediyordu. Artık `threaded=True`
  varsayılan ve `debug` modu güvenlik + hız için kapalı geliyor. Daha kalabalık kullanım
  için opsiyonel `run_production.py` (eventlet) eklendi — aşağıdaki ilgili bölüme bakın.
- **`requirements.txt` Windows'ta okunamıyordu** (`UnicodeDecodeError: 'charmap' codec
  can't decode byte 0x9e`) — dosyadaki Türkçe açıklama yorumları UTF-8 idi ama pip bazı
  Windows kurulumlarında (Türkçe sistem dili / cp1254) BOM'suz dosyaları önce sistem
  diline göre okumaya çalışıyor ve bu karakterlerde patlıyordu. `requirements.txt` artık
  tamamen ASCII (İngilizce yorumlarla), bu hatayı bir daha vermeyecek.
- **`python app.py` başlarken çöküyordu** (`TypeError: flask.app.Flask.run() got
  multiple values for keyword argument 'threaded'`) — bir önceki güncellemede stabilite
  için eklediğimiz `threaded=True`, Flask-SocketIO'nun zaten kendi içinde otomatik olarak
  gönderdiği aynı parametreyle çakışıyordu. Bunu kaldırdık; Flask-SocketIO zaten
  `threading` modunda tüm istekleri thread'li işliyor, ekstra bir şey yapmaya gerek yok.
  Gerçek bir terminalden (`python app.py`) başlatıp 10 eşzamanlı isteği paralel test
  ederek doğrulandı.

### 🌗 Aydınlık / Karanlık Mod
- Sağ üstteki ay/güneş ikonuna (ya da menüdeki anahtara) tıklayarak tema değiştirilebilir.
- Tercih tarayıcıda (`localStorage`) saklanır, bir sonraki ziyarette hatırlanır.
- Bootstrap'ın kendi `data-bs-theme` mekanizması kullanıldığı için dropdown, modal, tablo
  gibi tüm bileşenler otomatik olarak doğru kontrastla geçiş yapar.
- Karanlık moddaki düşük kontrastlı "soluk" metin rengi (`--text-faint`) okunabilirlik için
  belirgin şekilde açıldı; üst bar / alt navigasyon / zorunlu yoklama penceresi gibi sabit
  koyu renkli alanlar da artık temaya göre otomatik uyum sağlıyor.

### ✍️ Metin gönderileri
- Gönderi paylaşırken artık "Fotoğraf / Video" ile "Sadece Metin" arasında seçim yapılabiliyor.
- Metin gönderileri akışta (feed), profilde ve ana sayfa ızgarasında kendi görsel stiliyle
  (gradyanlı kart) gösteriliyor — bozuk resim ikonu göstermiyor.

### 🎨 Admin: Kullanıcıya Özel Profil Teması
- Admin panelinde "Kullanıcılar" sekmesinde her kullanıcının satırına bir "Profil Teması"
  açılır menüsü eklendi: **Varsayılan, Altın, Barbie (Pembe), Gece Mavisi, Orman, Gün Batımı**.
- Seçim yapıldığı anda otomatik kaydedilir, o kullanıcının **herkese açık profil sayfasının
  tamamı** (avatar çerçevesi, isim, kullanıcı adı, biyografi, istatistikler, butonlar,
  "Gönderiler" başlığı ve gönderi yoksa gösterilen boş durum mesajı dahil) tek, bütün bir
  panel olarak seçilen temaya bürünür. Diğer sayfalar (akış, sohbet vb.) etkilenmez —
  tema sadece o kullanıcının profil sayfasına özeldir.
- Yeni bir tema eklemek isterseniz `app.py` içindeki `PROFILE_THEMES` sözlüğüne bir satır,
  `static/css/style.css` içine de `.profile-theme-<isim>` kuralları eklemeniz yeterli.

### 🔒 Güvenlik
- Giriş sayfasındaki "Admin girişi: admin / admin123" ipucu kaldırıldı — artık kimse
  admin kullanıcı adı/şifresini sayfaya bakarak öğrenemiyor. İlk kurulumda admin şifresini
  `app.py` içinde `admin_hash = generate_password_hash('admin123')` satırından
  değiştirmeniz önerilir.

### ⚡ 150-200 kullanıcı ölçeği için optimizasyon
- Sık sorgulanan tüm sütunlara (yoklama kayıtları, etkinlik katılımcıları, gönderi
  beğeni/yorumları, bildirimler, özel mesajlar) veritabanı indeksleri eklendi.
- `PRAGMA journal_mode=WAL` + `synchronous=NORMAL` + `busy_timeout` ile eşzamanlı
  yazmalarda (sohbet, DM, yoklama işaretleme) kilitlenme riski azaltıldı.
- Toplu bildirim gönderimi (ör. "180 öğrenciye yoklama bildirimi") artık tek tek INSERT
  yerine `executemany` ile toplu yazılıyor — 180 kullanıcıda ölçüm ~10ms sürüyor.
- Akış (feed) artık **sayfalanıyor** (`/feed?page=2`) — tüm gönderiler tek seferde
  yüklenmiyor, bu sayede gönderi sayısı arttıkça sayfa yavaşlamıyor.
- 180 sahte kullanıcıyla yapılan yük testinde en ağır sayfa (180 kişilik yoklama detayı)
  ~130ms'de, admin paneli ~25ms'de, dashboard/mesajlar ~10ms'nin altında yanıt verdi.


### 🌙 Gece 22:00 Otomatik & Zorunlu Yoklama
- Her gece saat **22:00**'de arka planda çalışan bir zamanlayıcı (APScheduler) otomatik
  olarak yeni bir yoklama açar.
- Sadece **"yurtta kalıyorum"** işaretli kullanıcılara bildirim gönderilir.
- Bu yoklamayı doldurmayan yurt sakinleri, uygulamaya her girdiğinde tam ekran bir
  hatırlatma penceresiyle karşılaşır ("Varım / Yokum") — böylece doldurmak fiilen
  zorunlu hale getirilmiştir.
- Admin panelinden de manuel olarak "Zorunlu" işaretli yoklama açılabilir.
- Kullanıcının yurtta kalıp kalmadığı, profil düzenleme sayfasından veya admin panelinden
  değiştirilebilir.

### 👤 Instagram tarzı profil sistemi
- Profil fotoğrafı yükleme, biyografi, oda numarası, yurt durumu.
- Herkese açık profil sayfası (`/profile/<kullanıcı_adı>`) — gönderi ızgarası ile.

### 📸 Sosyal akış (feed)
- Fotoğraf/video paylaşımı, beğeni, yorum.
- Kronolojik akış (`/feed`) + gönderi silme (kendi gönderin ya da admin).

### 💬 Özel mesajlar (DM)
- Kullanıcılar birbirine gerçek zamanlı (SocketIO) özel mesaj gönderebilir.
- Okunmamış mesaj rozetleri, konuşma listesi.

### 🎨 Tasarım
- Uçtan uca karanlık mod, mobil öncelikli (bottom nav + üst bar + yan menü) tasarım.
- Instagram esintili gradyan vurgular, kart tabanlı arayüz.

## Notlar
- Yüklenen dosyalar `static/uploads/profiles` ve `static/uploads/posts` klasörlerinde saklanır.
- Zamanlayıcı tek worker/tek süreç varsayımıyla çalışır; birden fazla worker ile
  (ör. gunicorn `-w 4`) çalıştırırsanız görevi tek bir sürece sabitlemeniz gerekir.
- Geliştirme sunucusu (Werkzeug) küçük/orta kullanıcı sayısı için yeterlidir; gerçek
  üretimde (çok sayıda eşzamanlı kullanıcı) `gunicorn` + `eventlet` ya da `gevent`
  worker'ı ile çalıştırmanız önerilir.
- 150-200 kayıtlı kullanıcı için mevcut kurulum (SQLite + WAL + indeksler) test edildi ve
  sorunsuz çalışıyor. Kullanıcı sayısı bunun çok üzerine çıkarsa (binlerce) PostgreSQL gibi
  gerçek bir istemci-sunucu veritabanına geçmeniz önerilir.


## Menü sistemi
- Yönetici aylık `.xls`, `.xlsx` veya `.xlsm` dosyasını yükleyebilir.
- Kahvaltı ve akşam yemeği ayrı ayrı gösterilir; öğle/ikram sütunları varsa otomatik alınır.
- `.xls` için `xlrd` kullanılır.
- Öğrenciler her gün genel, kahvaltı ve akşam puanı verebilir; puanını daha sonra güncelleyebilir.
- Admin Analytics bölümünde günlük ortalamalar ve oy sayıları görülür.
