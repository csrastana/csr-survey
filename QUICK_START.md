# ⚡ БЫСТРЫЙ СТАРТ - 3 ШАГА

## 🎯 ВАЖНО: ОБНОВЛЕНИЯ

### Новые квоты:
- Астана: 800 (32 ПЕО × 25)
- Алматы: 975 (39 ПЕО × 25)
- Актобе: 525 (21 ПЕО × 25)
- Шымкент: 700 (28 ПЕО × 25)

### Новый Excel отчет:
- 4 листа (Dashboard, Polling Station, Enumerator, Raw Data)

### Чистый URL:
- `csr-survey.vercel.app` (без "fixed")

---

## 🚀 3 ШАГА ДО ЗАПУСКА

### ШАГ 1: GitHub (5 минут)

1. Создайте репозиторий: https://github.com/new
   - Название: `csr-survey`
   - Public или Private
   - Без README/gitignore/license
   
2. Загрузите ВСЕ файлы из этой папки
   - `index.html` должен быть **в корне**!
   - Папка `api/` с двумя файлами внутри
   
---

### ШАГ 2: Vercel (3 минуты)

1. **Удалите старый проект** `csr-survey-fixed` (если есть)
   - https://vercel.com/dashboard
   - Settings → General → Delete Project

2. **Создайте новый**
   - Add New → Project
   - Import `csr-survey` из GitHub
   
3. **Настройте**
   - Project Name: `csr-survey`
   - Framework Preset: **Other** ← ВАЖНО!
   - Всё остальное: оставить пустым
   
4. **Deploy**
   - Подождите 1-2 минуты
   - Готово! ✅

---

### ШАГ 3: Проверка (1 минута)

Откройте: `https://csr-survey.vercel.app`

**Должно быть:**
- ✅ Дэшборд с метриками
- ✅ Квоты: Астана 800, Алматы 975, и т.д.
- ✅ Детали: "32 ПЕО × 25 человек"
- ✅ Кнопка "Скачать Excel" работает
- ✅ Excel с 4 листами

---

## 🔧 ЕСЛИ URL ДРУГОЙ

Если получили `csr-survey-abc123.vercel.app`:

1. Settings → Domains → Add
2. Введите: `csr-survey.vercel.app`
3. Если занято - выберите другое
4. Обновите `API_BASE` в `index.html`

---

## ✅ ГОТОВО!

**URL:** https://csr-survey.vercel.app

Полная инструкция в `README.md`

---

**Проблемы?**
- 404 → Framework Preset = "Other"
- Пустой Excel → Проверьте данные в Kobo
- Failed to fetch → Проверьте API_BASE в index.html
