# MIC · ECOFF Analyzer — ซอร์สโค้ด

## ไฟล์
- `mic-ecoff-source.jsx` — ซอร์สโค้ดหลักทั้งหมด (React) อ่าน/แก้ที่ไฟล์นี้
- `mic-ecoff-platform.html` — เวอร์ชัน build พร้อมใช้ (เปิดได้เลย ไม่ต้องต่อเน็ต)

## โครงสร้างของ app.jsx (อ่านตามลำดับ)
1. **ค่าสี & helper** — `C` (พาเลตต์), `log2Label`, `parseMic`
2. **สถิติ/อัลกอริทึม** — `erf`/`normCdf`, `ecoffFinder` (วิธี ECOFFinder), `micPercentiles`, `qcChecks`, `buildAnalyses`
3. **อ่านไฟล์** — `detectFormat`, `normalizeRows`
4. **บันทึกข้อมูล** — `saveDataset`/`loadDataset`/... (ใช้ localStorage)
5. **คอมโพเนนต์ UI** — `Chip`, `Stat`, `Section`, `MicChart` (กราฟ SVG), `EmptyState`,
   `ImportModal`, `ComboDashboard`, `RawData`, `GroupedView`, `OverviewPage`, `HelpPage`, `App`

## วิธี build ใหม่เป็นไฟล์ HTML เดียว
ต้องมี Node.js แล้วรันในโฟลเดอร์โปรเจกต์:

```bash
npm install react@18 react-dom@18 papaparse@5 xlsx@0.18.5 tailwindcss@3.4.17

# 1) สร้าง Tailwind CSS (อ่านคลาสจาก src/)
npx tailwindcss -i src/input.css -o dist.css --minify

# 2) bundle JS รวมทุก library เป็นไฟล์เดียว
npx esbuild src/app.jsx --bundle --minify --format=iife --jsx=transform \
  --loader:.jsx=jsx --define:process.env.NODE_ENV='"production"' --outfile=app.bundle.js

# 3) เอา dist.css + app.bundle.js ฝังลงใน <style> และ <script> ของไฟล์ HTML
#    (โครง HTML: <div id="root"></div> + <script>...bundle...</script>)
```

โดย `src/input.css` มีแค่:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;
```
และ `tailwind.config.js`:
```js
module.exports = { content: ["./src/**/*.{js,jsx}"], theme:{extend:{}}, plugins:[] };
```

## หรือใช้แบบ dev server (แก้แล้วเห็นผลทันที)
ใส่ `app.jsx` ลงในโปรเจกต์ Vite + React + Tailwind แล้ว `npm run dev`
(เปลี่ยน `createRoot(...).render(<App/>)` ให้ชี้ element ของคุณ)
