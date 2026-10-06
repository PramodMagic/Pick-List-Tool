# Pick List Generator

Excel upload karo (Sheet 1 = Stock: `sku, bin, qty` | Sheet 2 = Requirement: `sku, qty`) → `pick_list.xlsx` auto download.
Output columns: `sku, bin, final bin, qty, remark`. Sab processing browser me hoti hai (file server par nahi jaati).

## Rules
1. Poori qty ek bin me ho → usi bin se (agar kai bins me ho to sabse chhota bin).
2. Ek bin me na ho → alag-alag bins se (bada bin pehle).
3. Stock kam ho → jitna hai utna, `remark` me "Short by X".
4. Stock zero / SKU stock me nahi → "Out of stock" remark.
5. Output bin ke hisaab se sorted (DF-2 pehle, DF-10 baad me).

## Local run
    npm install
    npm run dev

## Deploy (GitHub → Vercel)
1. Is folder ko GitHub repo me push karo.
2. vercel.com → Add New → Project → repo import → Deploy (koi setting nahi badalni).
