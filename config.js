/**
 * ตั้งค่าการเชื่อมต่อ Google Sheet
 *
 * API_URL : ลิงก์ Web app ของ Apps Script (ลงท้ายด้วย /exec)
 *           ถ้าเว้นว่าง หน้าเว็บจะเปิดเป็นโหมดตัวอย่าง (รายชื่อสมมติ ไม่บันทึกจริง)
 * API_KEY : ใส่เฉพาะเมื่อตั้ง Script Property ชื่อ API_KEY ไว้ใน Apps Script
 *           (กันคนสุ่มยิง API ได้ระดับหนึ่ง แต่ไม่ใช่รหัสลับ เพราะใครเปิดเว็บก็เห็นไฟล์นี้)
 */
window.CHECKIN_CONFIG = {
  API_URL: '',
  API_KEY: ''
};
