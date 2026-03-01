# Win Screen Translator (прототип)

Це Windows-версія ідеї з Decky-Translator: зняти екран → OCR → переклад → показати поверх гри/десктопа у вигляді прозорого оверлею.

## Як запустити (Windows)

1) Встанови Python 3.11+ (галочка `Add python.exe to PATH`).

2) У папці проєкту:

```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python run.py
```

3) Керування:
- **Ctrl + Shift + T**: показати/сховати переклад (toggle)
- Також є іконка в треї (правий клік) → Перекласти / Налаштування / Вихід

## Провайдери

- **OCR** за замовчуванням: `ocrspace` (OCR.space). Ключ не обов'язковий (demo `helloworld`), але для нормальної стабільності краще взяти безкоштовний ключ на OCR.space.
- **Переклад** за замовчуванням: `freegoogle` (неофіційний endpoint Google Translate).
- **Google Cloud**: якщо хочеш Vision + Translate, встав свій `google_api_key` і перемкни провайдери на `googlecloud`.

## RapidOCR (опційно, офлайн)

RapidOCR працює локально (без інтернету) і не має лімітів, але важчий по залежностях.

У цьому репо він встановлюється через `requirements.txt`.

За замовчуванням RapidOCR може використовувати свої дефолтні моделі/кеш.
Якщо хочеш жорстко вказати локальні моделі (офлайн, без докачування), поклади їх у:

`models/rapidocr/`

Мінімум (детекція + класифікація):
- `ch_PP-OCRv5_mobile_det.onnx`
- `ch_ppocr_mobile_v2.0_cls_infer.onnx`

Розпізнавання + словники (по сімействам):
- `ch_rec.onnx` + `ch_dict.txt`
- `latin_rec.onnx` + `latin_dict.txt`
- `eslav_rec.onnx` + `eslav_dict.txt`
- `korean_rec.onnx` + `korean_dict.txt`
- `greek_rec.onnx` + `greek_dict.txt`
- `thai_rec.onnx` + `thai_dict.txt`

Потім у налаштуваннях постав `ocr_provider = rapidocr`.

## Постійний режим

У налаштуваннях є чекбокс **Постійний режим** + інтервал (мс).
Коли він увімкнений, **Ctrl+Shift+T** працює як Start/Stop (оновлює переклад з екрану по таймеру).

## Налаштування

Файл налаштувань зберігається тут:
- `%APPDATA%\WinScreenTranslator\settings.json`

