# Panduan Deploy ke Google Apps Script

## Persiapan (Sekali Saja)

### 1. Install Node.js & Clasp
```bash
# Install clasp
npm install -g @google/clasp
```

### 2. Login ke Google
```bash
clasp login
```
- Browser akan terbuka
- Login dengan akun Google yang ingin digunakan
- Berikan izin akses

### 3. Buat Project Baru di Apps Script
1. Buka https://script.google.com
2. Klik **New Project**
3. Copy **Script ID** dari URL:
   ```
   https://script.google.com/d/SCRIPT_ID_INI/edit
   ```

### 4. Update .clasp.json
Edit file `.clasp.json` dan isi Script ID:
```json
{
  "scriptId": "SCRIPT_ID_ANDA_DI_SINI",
  "rootDir": "."
}
```

## Deploy

### Push Semua File
```bash
clasp push
```

### Buka di Browser
```bash
clasp open
```

### Deploy sebagai Web App
```bash
clasp deploy --description "HifdzulApp v6.1"
```

## Update Setelah Perubahan
```bash
# Push perubahan
clasp push

# Deploy ulang
clasp deploy
```

## File yang Di-push
- `*.gs` → Backend (Auth, Utils, Santri, dll)
- `*.html` → Frontend (Index.html, css.html, js.html)

## Troubleshooting

### Error "No script Id found"
- Pastikan `.clasp.json` sudah diisi Script ID

### Error "Permission denied"
- Jalankan `clasp login` ulang

### File tidak muncul
- Pastikan file ada di folder yang sama dengan `.clasp.json`
