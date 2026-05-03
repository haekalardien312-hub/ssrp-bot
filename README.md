# 🎭 SSRP Checker Bot v2

Bot Discord otomatis untuk validasi SSRP, sistem poin, export data, dan deteksi member tidak aktif.

---

## ✨ Fitur
- 🤖 **AI Validation** — Gemini AI baca screenshot, validasi otomatis (GRATIS)
- ⚡ **Auto Point** — Kirim foto → langsung dapat poin, tanpa admin approve
- ⏳ **Anti-Spam** — Cooldown 30 detik antar submit
- 📊 **Export Data** — Export ke Excel (.xlsx) dan CSV sekaligus
- ⚠️ **Deteksi Tidak Aktif** — Bot otomatis cek & tag siapa yang belum SSRP seminggu
- 📩 **DM Warning** — Bot DM langsung ke member yang tidak aktif
- 🏆 **Leaderboard** — Ranking poin mingguan & total

---

## 🛠️ Setup

### 1. Install Python 3.10+
Download dari https://python.org

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Dapatkan Gemini API Key (GRATIS)
1. Buka https://aistudio.google.com/app/apikey
2. Login dengan akun Google
3. Klik **"Create API Key"**
4. Copy key-nya

### 4. Dapatkan Discord Bot Token
1. Buka https://discord.com/developers/applications
2. **New Application** → beri nama → Create
3. Tab **Bot** → **Add Bot**
4. Aktifkan:
   - ✅ Message Content Intent
   - ✅ Server Members Intent
5. **Reset Token** → Copy

### 5. Edit Konfigurasi di bot.py
```python
DISCORD_TOKEN         = "TOKEN_DISCORD_KAMU"
GEMINI_API_KEY        = "API_KEY_GEMINI_KAMU"
APPROVAL_CHANNEL_NAME = "ssrp-aproval"     # Channel submit
REPORT_CHANNEL_NAME   = "ssrp-report"      # Channel laporan
INACTIVE_CHANNEL_NAME = "ssrp-inactive"    # Channel notif tidak aktif
SUBMIT_COOLDOWN       = 30                 # Jeda anti-spam (detik)
```

### 6. Buat Channel di Discord
Buat 3 channel:
- `#ssrp-aproval` — tempat player submit screenshot
- `#ssrp-report` — laporan otomatis muncul di sini
- `#ssrp-inactive` — notif mingguan siapa yang tidak aktif

### 7. Invite Bot
1. Developer Portal → **OAuth2** → **URL Generator**
2. Centang: `bot`
3. Bot Permissions:
   - ✅ Read Messages
   - ✅ Send Messages
   - ✅ Embed Links
   - ✅ Attach Files
   - ✅ Read Message History
   - ✅ Add Reactions
   - ✅ Manage Messages
4. Copy URL → invite ke server

### 8. Jalankan
```bash
python bot.py
```

---

## 📌 Commands

### 👤 Member
| Command | Fungsi |
|---------|--------|
| `!point` | Cek poin kamu |
| `!point @user` | Cek poin member lain |
| `!lb` | Leaderboard minggu ini |
| `!ssrphelp` | Bantuan |

### 🔧 Admin Only
| Command | Fungsi |
|---------|--------|
| `!export excel` | Export data ke Excel |
| `!export csv` | Export data ke CSV |
| `!export all` | Export keduanya sekaligus |
| `!inactive` | Cek tidak aktif 7 hari |
| `!inactive 14` | Cek tidak aktif 14 hari |
| `!resetweek` | Reset poin mingguan |

---

## ⚙️ Cara Kerja

1. Player kirim screenshot ke `#ssrp-aproval`
2. Bot cek cooldown (anti-spam 30 detik)
3. Gemini AI validasi apakah gambar benar SSRP
4. Kalau **valid** → langsung +1 poin, muncul laporan di `#ssrp-report`
5. Kalau **tidak valid** → pesan penolakan, tidak dapat poin
6. **Setiap Senin pagi** → bot cek siapa yang tidak submit 7 hari:
   - Tag di `#ssrp-inactive`
   - DM langsung ke member tersebut

---

## 📁 File Structure
```
ssrp-bot-v2/
├── bot.py              ← Kode utama
├── requirements.txt    ← Dependencies
├── README.md           ← Panduan ini
└── ssrp_points.db      ← Database (auto dibuat)
```
