# CRM Plan Generator

يولّد خطة CRM لأول 100 باص مرتبة حسب الأولوية من الخطة الكاملة.

## تشغيل محلي
```bash
pip install streamlit pandas openpyxl xlrd
streamlit run app.py
```

## نشر على Streamlit Cloud
1. ارفع `app.py` و `requirements.txt` على GitHub repo جديد
2. اذهب على share.streamlit.io
3. اختر الـ repo → branch: main → file: app.py → Deploy

## Priority Logic
- **Passes Left = 1** 🟡 → أرسله فوراً للمرحلة التالية
- **Passes Left = 2** 🟢 → أولوية عالية
- **Passes Left 3+** ⬜ → عادي
- **Not in master** 🔴 → آخر الخطة
- **Warm-up coils** 🟨 → بعد تغيير Work Roll (كويلات سميكة P1/P2/P3)

## هيكل الخطة الناتجة
| الصفوف | المحتوى |
|--------|---------|
| 1–100  | الباصات مرتبة بالأولوية |
| 101    | ⚠️ CHANGE WORK ROLL |
| 102    | 🔥 Header كويلات التسخين |
| 103–105 | 3 كويلات تسخين (أسمك P1) |
| 106    | 📋 UPDATE THE PLAN FROM HERE |
