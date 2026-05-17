# PsyScreen (Poliklinik Ekran Yöneticisi)

Tkinter tabanlı masaüstü uygulaması: bekleme salonu / ikinci monitörde anons, mola sayacı, hasta çağırma; isteğe bağlı Telegram, Excel ve Notion loglama.

---

## Güncellemeler (yeniden eskiye)

- **2026-05-17** — v6.1: hasta ismi maskeleme seçeneği; kanonik giriş dosyası `PsyScreen.py`. Yerel sürüm dosyaları (`PsyScreen_6_0.py`, `PsyScreen_6_1.py`) repoya dahil değil.
- **2026-04-04** — GitHub için repo düzeni: `.gitignore`, `profiller.example.json`, `requirements.txt`, `hastaneler/.gitkeep`, yerel `.cursor/RULES.md`, bu README ve güvenlik odaklı ignore kuralları eklendi.

---

## Gereksinimler

- Python **3.9+** (önerilir; `zoneinfo` ve Windows'ta `tzdata` için)
- İkinci monitör kullanımı için `screeninfo` uyumlu ortam

## Kurulum

```bash
cd PsyScreen
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### Doktor profili (zorunlu, yerel)

`profiller.json` **depoya dahil değildir** (içinde API anahtarları olur).

1. `profiller.example.json` dosyasını kopyalayın: `profiller.json`
2. Uygulama içinden **Profil Yöneticisi** ile veya düzenleyerek Telegram, Notion, API alanlarını doldurun.

### Hastane logoları

`hastaneler/` klasörü repoda boş tutulur (yalnızca `.gitkeep`). Logo dosyalarınızı buraya ekleyin; [hastaneler.json](hastaneler.json) içinde yolları `hastaneler/dosya.png` gibi gösterin.

## Çalıştırma

```bash
python PsyScreen.py
```

Sistem tepsisinden kontrol panelini açıp kapatabilirsiniz.

## Önemli dosyalar

| Dosya | Açıklama |
| --- | --- |
| `PsyScreen.py` | Ana uygulama (v6.1) |
| `profiller.example.json` | Profil şablonu (anahtarlar boş) |
| `mesajlar.json` | Hazır anons metinleri |
| `hastaneler.json` | Hastane adı ve logo yolu |

Yerel üretilen dosyalar (`.gitignore`): `mola_kayitlari.xlsx`, `hasta_cagirma_kayitlari.xlsx`, `muayene_sureleri.xlsx`, `hata_gunlugu.txt`.

## Güvenlik

- `profiller.json`, `.env` ve benzeri sırları **asla** commit etmeyin.
- Önceden herkese açık bir depoda sırlar paylaşıldıysa: Telegram bot token'ını yenileyin, Notion entegrasyon anahtarını döndürün, poliklinik API bearer'ını değiştirin.

## Lisans

Belirtilmemiş; kullanım için proje sahibi ile iletişime geçin.
