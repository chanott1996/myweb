# MIC · ECOFF Analyzer

เว็บแอปวิเคราะห์การกระจาย MIC และกำหนดค่า ECOFF (React + Vite)

## วิธี deploy ขึ้น GitHub Pages (build อัตโนมัติ — ไม่ต้องยุ่งกับ HTML)

1. สร้าง repository ใหม่บน GitHub
2. อัปโหลดไฟล์ในโปรเจกต์นี้ขึ้น repo ทั้งหมด
   - **สำคัญ:** ต้องมีโฟลเดอร์ `.github/workflows/deploy.yml` ติดไปด้วย (เป็นไฟล์ซ่อน)
   - ถ้าอัปผ่านหน้าเว็บ GitHub ให้ลากทั้งโฟลเดอร์เข้าไป หรือใช้ git (ดูด้านล่าง)
3. ไปที่ repo → **Settings → Pages → Build and deployment → Source** เลือก **GitHub Actions**
4. ทุกครั้งที่ push ขึ้น branch `main` → GitHub Actions จะ build แล้ว deploy ให้เอง
   ดูสถานะที่แท็บ **Actions** เมื่อเสร็จจะได้ลิงก์เว็บ (เช่น `https://<username>.github.io/<repo>/`)

### อัปโหลดด้วย git (แนะนำ — ได้ไฟล์ซ่อนครบ)
```bash
git init
git add .
git commit -m "init MIC ECOFF analyzer"
git branch -M main
git remote add origin https://github.com/<username>/<repo>.git
git push -u origin main
```

## รันบนเครื่อง (พัฒนา/แก้โค้ด)
ต้องมี Node.js 18+
```bash
npm install
npm run dev      # เปิด dev server แก้โค้ดเห็นผลทันที
npm run build    # สร้างไฟล์ production ลงโฟลเดอร์ dist/
npm run preview  # ดู production build ในเครื่อง
```

## โครงสร้าง
```
index.html              จุดเริ่ม (โหลด src/main.jsx)
src/main.jsx            mount React
src/App.jsx             โค้ดแอปทั้งหมด (อัลกอริทึม ECOFF + UI ทุกหน้า)
src/index.css          Tailwind + สไตล์พื้นฐาน
vite.config.js         ตั้ง base:"./" สำหรับ GitHub Pages
tailwind.config.js     postcss.config.js   ตั้งค่า Tailwind
.github/workflows/deploy.yml   GitHub Actions: build + deploy อัตโนมัติ
```

## หมายเหตุ
- ข้อมูลที่อัปโหลดถูกบันทึกใน localStorage ของเบราว์เซอร์ผู้ใช้แต่ละคน (ไม่ได้ sync ข้ามเครื่อง)
- ค่า ECOFF ที่คำนวณเป็นการประมาณเชิงสถิติ ควรเทียบกับค่าทางการของ EUCAST ก่อนใช้งานจริง
