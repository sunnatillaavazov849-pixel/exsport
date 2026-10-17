# Pechat Eksport — CorelDRAW uchun plagin (makros)

Rasmni pechatga (banner, orakal, bekprint) bitta oynadan eksport qiladi.
Hamma tanlovlar ochiladigan ro'yxatda:

1. **Rang rejimi** — CMYK yoki RGB
2. **Ranglarni ochroq qilish** — Yo'q yoki 5–50% (RGB rasm CMYK ga o'tganda rangi to'qlashib ketmasligi uchun, odatda 10–20%)
3. **Pechat turi** — Banner / Orakal / Bekprint
4. **Format** — TIF, PNG, JPG, JPEG
5. **Razreshenie (dpi)** — masalan 150 (o'zingiz ham yozsangiz bo'ladi)

**Nima eksport qilinadi:**

- Hech narsa tanlanmagan bo'lsa — butun sahifa.
- Bitta obyekt tanlangan bo'lsa — shu obyekt.
- **Bir nechta obyekt tanlangan bo'lsa — har biri alohida fayl bo'lib, o'z o'lchami bilan saqlanadi.**

Fayl nomi o'lcham, tur va rangdan avtomatik tuziladi:

```
150 x 200 cm_banner_cmyk.tif
```

O'lcham sahifa (yoki har bir obyekt) o'lchamidan santimetrda olinadi.
Oynaning pastida fayl nomlari oldindan ko'rsatib turiladi.
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

1. Dizaynni sahifada 150 x 200 sm qilib tayyorlang, yoki saqlanadigan rasmlarni tanlang (bir nechtasini ham).
2. `PechatgaEksport` tugmasini bosing.
3. Rang, ochroqlik, tur, format va dpi ni tanlang → **Eksport**.
4. Fayl tanlangan papkaga `150 x 200 cm_banner_cmyk.tif` kabi nom bilan saqlanadi.

## Eslatmalar

- **Ochroq qilish** eksport paytida rasm ustiga vaqtinchalik "Brighten" linzasi
  qo'yib bajariladi, sizning asl hujjatingiz o'zgarmaydi (linza eksportdan keyin o'chiriladi).
- **PNG** formati CMYK ni qo'llamaydi — PNG tanlansa avtomatik RGB bo'ladi.
- Katta bannerlar (masalan 150 x 200 sm, 150 dpi, CMYK TIF) bir necha yuz MB bo'lishi mumkin,
  eksport biroz vaqt oladi.
- Shu nomli fayl bo'lsa (masalan, ikkita rasm bir xil o'lchamda), ustidan yozilmaydi —
  nomiga `(2)`, `(3)` qo'shiladi: `150 x 200 cm_banner_cmyk (2).tif`.
