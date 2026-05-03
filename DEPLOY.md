# 📚 EduSave Bot — Render.com ga deploy qilish

## 🚀 Qadama-qadam yo'riqnoma

### 1️⃣ GitHub akkaunt yaratish (agar yo'q bo'lsa)

1. https://github.com ga boring
2. **Sign up** → akkaunt yarating

### 2️⃣ GitHub da yangi repository yarating

1. https://github.com/new ga boring
2. **Repository name**: `edusave-bot`
3. **Public** tanlang
4. **Create repository** bosing

### 3️⃣ Fayllarni GitHub ga yuklang

Eng oson yo'l — **GitHub web interface** orqali:

1. Yaratilgan repository da **"uploading an existing file"** havolasini bosing
2. Quyidagi 4 ta faylni drag & drop qiling:
   - `bot.py`
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`
3. Pastda **"Commit changes"** bosing

### 4️⃣ Render.com da akkaunt yarating

1. https://render.com ga boring
2. **Get Started** → **Sign in with GitHub**
3. Ruxsat bering

### 5️⃣ Yangi Web Service yarating

1. Render dashboard da **"New +"** → **"Web Service"**
2. **"Connect a repository"** → `edusave-bot` ni tanlang
3. Sozlamalar:
   - **Name**: `edusave-bot`
   - **Region**: Frankfurt (Europe)
   - **Branch**: main
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python bot.py`
   - **Plan**: **Free** ✅

### 6️⃣ Environment Variable qo'shing

1. Pastga scroll qiling **"Environment Variables"** bo'limigacha
2. **"Add Environment Variable"** bosing
3. **Key**: `BOT_TOKEN`
4. **Value**: BotFather dan olgan tokeningiz
5. **"Create Web Service"** bosing

### 7️⃣ Kuting (5-10 daqiqa)

Render botni o'rnatadi. Logs da ko'rinishi kerak:
```
🌐 Web server ishga tushdi (port 8080)
🤖 Bot ishga tushdi!
```

### 8️⃣ URL ni nusxalang

Yuqorida URL ko'rinadi: `https://edusave-bot.onrender.com`

Shu URL ga brauzerda kirsangiz: **"🤖 EduSave Bot ishlamoqda!"** chiqishi kerak.

---

## 🔄 UptimeRobot — Bot uxlamasligini ta'minlash

Render bepul rejada **15 daqiqa harakatsizlikdan keyin uxlab qoladi**. Bu botni o'chiradi. Yechim — har 5 daqiqada URL ga so'rov yuborish:

### 1️⃣ UptimeRobot akkaunt yarating

1. https://uptimerobot.com ga boring
2. Bepul akkaunt yarating

### 2️⃣ Monitor qo'shing

1. **+ Add New Monitor** bosing
2. Sozlamalar:
   - **Monitor Type**: HTTP(s)
   - **Friendly Name**: `EduSave Bot`
   - **URL**: `https://edusave-bot.onrender.com/health`
   - **Monitoring Interval**: 5 minutes
3. **Create Monitor** bosing

✅ Tayyor! Endi UptimeRobot har 5 daqiqada botingizga so'rov yuborib turadi va u uxlamaydi.

---

## ⚠️ Muhim eslatmalar

- 🆓 Render bepul rejada oyiga 750 soat beradi (taxminan 31 kun)
- 💾 Render bepul rejada **disk storage cheklangan** — SQLite DB faqat session davomida saqlanadi
- 📦 Doimiy DB uchun keyinroq PostgreSQL qo'shamiz (bepul Render da bor)
- 🔄 Kod o'zgartirsangiz GitHub ga push qiling — Render avtomatik yangilaydi

---

## 🆘 Muammo bo'lsa

Render Logs ni tekshiring (dashboard → Logs tab) — xatolik ko'rinadi.
