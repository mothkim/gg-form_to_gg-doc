# สร้าง Google Form สำหรับพริ้นเอกสาร (ยังเขียนไม่เสร็จ)#

1. สร้างไฟล์ Google Doc Template เอกสารที่ต้องการ
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/6f7712dd-a294-4444-8ecb-f8c6be7e1fcd">

```
ในเอกสาร จะมี {{TITEL}} , {{FIRSTNAME}} , {{LASTNAME}} คือเป็นค่าที่เราไว้สำหรับทดแทน เมื่อมีการรับค่ามากจาก Google Form
```

2. สร้าง Google Form
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/fde983a3-0677-4fe4-9058-a76224145a82">

---

3. กรอกรายละเอียดหัวข้อที่ต้องการให้กรอก
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/4e64b973-4f5f-4d1e-85bd-e6d044f985b9">

---

4. ลิงค์ไปที่ Google Sheet
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/d6ff6af7-7256-47de-a4d5-c8116c62ee34">

<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/0b24fb54-66db-448d-9924-3ef78c3d215e">

หน้าตาจะประมาณนี้ หลักๆ ดูตรง Field ที่เป็นหัวข้อ
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/b5aa77ca-5bd3-4530-9864-2b0a006e833f">

---

5. ทำการลิงค์กับ App script
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/0b036bdb-de88-435d-a34d-ab48c88ad0fe">


6. วาง script ตามนี้

```
function autoFillDoc(e) {
  // ตัวแปรนี้ จะไปหยิบค่าจากไฟล์ google sheet โดยนับจากคอลัมป์ ซ้ายไปขวา (***Array จะเริ่มต้นที่ 0 ) สามารถเพ่ิ่มต่อท้ายหรือแทรกเองได้โดยอ้างอิงจาก Google Sheet
  // เช่น ถ้ามี Column ใหม่เกิดขึ้นที่ Google sheet ชื่อ "gender" ก็ให้มาเพิ่มตัวแปร เพื่ออ่านค่าเข้าไป ก็จะได้เป็น
  // var gender = e.values[11];
  // แล้วก็อย่าลืมไปเพิ่มที่  body.replaceText ว่าจะ replace แทนคำว่าอะไรด้วย ไม่งั้นก็จะไม่แสดงบนเอกสาร template

  var timeStamp = e.values[0];
  var title = e.values[1];
  var firstName = e.values[2];
  var lastName = e.values[3];
  var tel = e.values[4];
  var studentid = e.values[5];
  var gen = e.values[6];
  var group = e.values[7];
  var year = e.values[8];
  var course = e.values[9];
  var address = e.values[10];

  // ใส่ ID ของไฟล์ที่เป็นไฟล์ Template เอกสาร โดยเอามาจาก URL ตอนเปิดไฟล์จะมี URL ด้านบน ตัวอย่าง "https://docs.google.com/document/d/ตรงนี้คือID/edit" 
  var file = DriveApp.getFileById("ใส่ ID ของชื่อไฟล์ Template");

  // ใส่ ID ของ Folder บน Google Drive ที่ใช้สำหรับสร้างไฟล์ และหยิบไฟล์ template โดยเอามาจาก URL ตอนเปิด Folder บน Google Drive ตัวอย่าง "https://drive.google.com/drive/folders/ตรงนี้คือIDของFolder"
  var folder = DriveApp.getFolderById('ใส่ ID ของ Folder ที่จะหยิบไฟล์ Template ตรงนี้');

  // อันนี้จะเป็น Folder ที่ต้องการเก็บไฟล์ผลลัพธ์ แยกออกไป วิธีการดู ID ทำเหมือนเดิมจาก step ก่อนหน้านี้
  var folderResult = DriveApp.getFolderById('ใส่ ID ของ Folder สำหรับเก็บผลลัพธ์ ตรงนี้');

  // อันนี้ ไว้ตั้งชื่อไฟล์ สามารถแก้ไขได้ ในนี้จะใช้เป็น คำนำหน้าชื่อ_ชื่อ_นามสกุล
  var fileName = title + '_' + firstName + '_' + lastName;

  // ไฟล์จะถูกเก็บไว้ที่อีก Folder เพื่อจะได้ไม่รก เกินไป
  var newFile = file.makeCopy(fileName, folderResult);

  var newFileId = newFile.getId();
  var doc = DocumentApp.openById(newFileId);
  
  //change is in the body, we need to get that
  //see more at https://developers.google.com/apps-script/reference/document/document
  var body = doc.getBody();
  
  //call all of our replaceText methods
  //https://developers.google.com/apps-script/reference/document/body
  // ตรงนี้เป็นจังหวะที่จะแทนค่าลงไปในเอกสาร Template ถ้าตั้งไม่ถูกค่าก็จะไม่เปลี่ยน ต้องไปไล่ดูว่าใส่ค่า หรือ หยิบค่ามาถูกไหม

  body.replaceText('{{TITEL}}', title);
  body.replaceText('{{FIRSTNAME}}', firstName); 
  body.replaceText('{{LASTNAME}}', lastName);  
  body.replaceText('{{TEL}}', tel);
  body.replaceText('{{STUDENTID}}', studentid);
  body.replaceText('{{GEN}}', gen);
  body.replaceText('{{GROUP}}', group);
  body.replaceText('{{YEAR}}', year);
  body.replaceText('{{COURSE}}', course);
  body.replaceText('{{ADDRESS}}', address);

  //save and close the document to persist our changes
  //see more at https://developers.google.com/apps-script/reference/document/document
  doc.saveAndClose();
}
```

---

7. กำหนด Trigger ให้ทำงานเมื่อมีการกด submit

<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/3b29c676-a980-4d2a-b058-37b9fd68248a">

<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/cae28690-77b8-4879-b215-16054f78d8b1">

---

8. ทดสอบกรอก Google Form
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/d5a7b113-41d4-44c7-b776-894f1e5a3513">


9. ดูผลลัพธ์
<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/76f2507b-bbfc-4b34-8d9c-d23c20f578e3">

<img width="500" high="500" alt="image" src="https://github.com/user-attachments/assets/24968eb2-410c-47f2-bf3b-9a92e51a6092">
