# อัปเดตระบบด้วย GitHub (แทนการคัดลอกไฟล์ทีละไฟล์)

หลังตั้งค่าครั้งเดียวตามด้านล่าง การอัปเดตทุกครั้งคือ **GitHub → Actions → deploy → Run workflow** เท่านั้น
ระบบจะรันชุดทดสอบทั้งหมดก่อน แล้วส่ง **ทุกไฟล์** ขึ้น Apps Script และทำ New version ให้ ถ้าทดสอบไม่ผ่านจะไม่ Deploy

## ไฟล์ที่เกี่ยวข้อง
| ไฟล์ | หน้าที่ |
|---|---|
| `.github/workflows/test.yml` | รันชุดทดสอบทุกครั้งที่ push และเมื่อมี Pull request |
| `.github/workflows/deploy.yml` | กดรันเอง: ทดสอบ → `clasp push` → `clasp deploy` ไปที่ Deployment เดิม (ลิงก์เว็บไม่เปลี่ยน) |
| `.claspignore` | ส่งขึ้น Apps Script เฉพาะ `*.gs`, `index.html`, `OrderApprovalPage.html`, `appsscript.json` |
| `.clasp.json.example` | ตัวอย่างสำหรับใช้ clasp บนเครื่องตัวเอง (ไฟล์จริง `.clasp.json` ไม่ถูก commit) |

## ตั้งค่าครั้งเดียว (เจ้าของ Apps Script ทำ)
1. **เปิด Apps Script API** ของบัญชี Google ที่เป็นเจ้าของสคริปต์: https://script.google.com/home/usersettings → Google Apps Script API = On
2. **ล็อกอิน clasp บนเครื่อง PC** (ต้องมี Node.js):
   ```
   npm install -g @google/clasp@2.4.2
   clasp login
   ```
   จะได้ไฟล์ `.clasprc.json` ในโฟลเดอร์ผู้ใช้ (Windows: `C:\Users\<ชื่อ>\.clasprc.json`)
3. **เก็บรหัส 3 ค่าใน GitHub** → repository → Settings → Secrets and variables → Actions → New repository secret
   - `CLASPRC_JSON` = เนื้อหาทั้งไฟล์ `.clasprc.json` (เป็นสิทธิ์เข้าถึงบัญชี — ห้ามส่งในแชตหรือ commit)
   - `SCRIPT_ID` = Apps Script → ⚙ Project Settings → Script ID
   - `DEPLOYMENT_ID` = Apps Script → Deploy → Manage deployments → Web app → Deployment ID
4. **ตรวจไฟล์ใน Apps Script ก่อน Deploy ครั้งแรก**: `clasp push` จะทำให้โปรเจกต์มีไฟล์ตรงกับ GitHub ทุกไฟล์ — ไฟล์ที่มีใน Apps Script แต่ไม่มีใน GitHub จะถูกลบ (ถ้ามีไฟล์ทดลองของตัวเองให้ย้ายเก็บก่อน)

## อัปเดตระบบ
GitHub → **Actions** → **deploy** → **Run workflow** → ใส่หมายเหตุเวอร์ชัน (ไม่ใส่ก็ได้) → Run
- ทดสอบผ่าน → ส่งไฟล์ → New version บน Deployment เดิม (ผู้ใช้ใช้ลิงก์เดิม)
- ทดสอบไม่ผ่าน → หยุด ไม่มีอะไรเปลี่ยนในระบบจริง · ดูผลที่ Actions → test → Artifacts (`test-output`)

## ถ้ายังไม่พร้อมตั้งค่า
ใช้วิธีเดิมได้ตามปกติ (คัดลอกไฟล์ที่เปลี่ยนไปวางใน Apps Script แล้ว Deploy → New version)
