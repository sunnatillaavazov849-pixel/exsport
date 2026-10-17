# Pechat Eksport — CorelDRAW uchun plagin (makros)

Rasmni pechatga (banner, orakal, bekprint) bitta oynadan eksport qiladi:

1. **Rang rejimi** — CMYK yoki RGB
2. **Ranglarni ochroq qilish** — RGB rasm CMYK ga o'tganda rangi to'qlashib ketmasligi uchun (foizini tanlaysiz, odatda 10–20%)
3. **Pechat turi** — Banner / Orakal / Bekprint
4. **Format** — TIF, PNG, JPG, JPEG
5. **Razreshenie (dpi)** — masalan 150
6. **Nimani eksport qilish** — butun sahifa yoki faqat tanlangan obyekt

Fayl nomi o'lcham, tur va rangdan avtomatik tuziladi:

```
150 x 200 cm_banner_cmyk.tif
```

O'lcham sahifa (yoki tanlangan obyekt) o'lchamidan santimetrda olinadi.
Oynaning pastida fayl nomi va piksel o'lchami oldindan ko'rsatib turiladi.
Oxirgi tanlovlaringiz eslab qolinadi.

## O'rnatish (bir marta)

1. CorelDRAW ni oching, **Alt + F11** bosing (Macro Editor / VBA oynasi ochiladi).
2. Chap tomonda **GlobalMacros** (GlobalMacros.gms) loyihasini tanlang.
3. **File → Import File...** orqali `src/PechatExport.bas` faylini import qiling.
4. **Insert → UserForm** bosing. Yangi formaning **(Name)** xossasini (Properties oynasida, F4) `frmExport` deb o'zgartiring.
5. Formaga ikki marta bosib kod oynasini oching, u yerdagi hamma narsani o'chirib,
   `src/frmExport.txt` faylining ichidagi kodni to'liq qo'ying.
   (Formaga hech qanday tugma qo'yish shart emas — hammasi kod orqali o'zi yaratiladi.)
6. **Ctrl + S** bilan saqlang va VBA oynasini yoping.

### Panelga tugma qo'yish

**Tools → Options → Customization → Commands** (CorelDRAW versiyasiga qarab
*Tools → Customization*) bo'limida ro'yxatdan **Macros** ni tanlang,
`GlobalMacros.PechatExport.PechatgaEksport` ni toping va uni sichqoncha bilan
panelga tortib qo'ying. Xohlasangiz klaviatura tugmasini ham biriktirsa bo'ladi.

Yoki: **Tools → Macros → Run Macro** → `PechatgaEksport`.

## Ishlatish

1. Dizaynni sahifada 150 x 200 sm qilib tayyorlang (yoki obyektni tanlang).
2. `PechatgaEksport` tugmasini bosing.
3. Rang, ochroqlik, tur, format va dpi ni tanlang → **Eksport**.
4. Fayl tanlangan papkaga `150 x 200 cm_banner_cmyk.tif` kabi nom bilan saqlanadi.

## Eslatmalar

- **Ochroq qilish** eksport paytida rasm ustiga vaqtinchalik "Brighten" linzasi
  qo'yib bajariladi, sizning asl hujjatingiz o'zgarmaydi (linza eksportdan keyin o'chiriladi).
- **PNG** formati CMYK ni qo'llamaydi — PNG tanlansa avtomatik RGB bo'ladi.
- Katta bannerlar (masalan 150 x 200 sm, 150 dpi, CMYK TIF) bir necha yuz MB bo'lishi mumkin,
  eksport biroz vaqt oladi.
- Fayl allaqachon bo'lsa, ustidan yozish so'raladi.
