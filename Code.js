const EMP_SPREADSHEET_ID = '1fcLEYsJnZpHB-UY-9q7vLlRV4pp626tNV4kl6MhkDyA';
const EMP_SHEET_NAME = 'EmpDetails';

const TOOL_SPREADSHEET_ID = '1VP176486EGKw-Dp4fPQGmhsKmmJj_fCmPsznO5s3LbY';
const TOOL_SHEET_NAME = 'ToolsCount'; 

const MAIN_FOLDER_ID = '1UNMr3rVeyCM83BgSPLUIa9CTjbE-axnB';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบบันทึกเครื่องมือทีมช่าง')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. ฟังก์ชันตรวจสอบและดึงข้อมูลพนักงาน
function checkEmployee(empId) {
  try {
    const ss = SpreadsheetApp.openById(EMP_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(EMP_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === empId.toString().trim()) {
        return {
          success: true,
          empId: data[i][0],
          empName: data[i][1]
        };
      }
    }
    return { success: false, msg: 'ไม่พบรหัสพนักงานนี้ในระบบ' };
  } catch (error) {
    return { success: false, msg: error.toString() };
  }
}

// 2. ฟังก์ชันบันทึกข้อมูล (เพิ่มการบันทึกชื่อและรหัสพนักงานลงบนชีตและ PDF)
function saveData(payload) {
  try {
    const ssTool = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID);
    let targetSheet = ssTool.getSheetByName(TOOL_SHEET_NAME);
    
    // ตั้งค่าหัวคอลัมน์รวมถึง emp_id และ emp_name
    if (!targetSheet) {
      targetSheet = ssTool.insertSheet(TOOL_SHEET_NAME);
      targetSheet.appendRow(['timestamp', 'emppin', 'tagnum', 'full_Image', 'tag_Image', 'pdf_Report', 'emp_id', 'emp_name']);
    }

    const cleanAssetId = payload.assetId.toString().trim().toUpperCase();
    const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
    
    // สร้างชื่อโฟลเดอร์ย่อย: รหัสทรัพย์สิน_YYMMDDHHMM
    const formattedDate = getFormattedDate();
    const subFolderName = `${cleanAssetId}_${formattedDate}`;
    const subFolder = mainFolder.createFolder(subFolderName);
    
    // บันทึกรูปภาพตัวเครื่องมือลง Drive (full_Image)
    const toolImgBlob = Utilities.newBlob(Utilities.base64Decode(payload.toolImg.split(',')[1]), 'image/jpeg', `TOOL_${cleanAssetId}.jpg`);
    const toolFile = subFolder.createFile(toolImgBlob);
    
    // บันทึกรูปภาพรหัสทรัพย์สินลง Drive (tag_Image)
    const assetImgBlob = Utilities.newBlob(Utilities.base64Decode(payload.assetImg.split(',')[1]), 'image/jpeg', `ASSET_${cleanAssetId}.jpg`);
    const assetFile = subFolder.createFile(assetImgBlob);

    const saveDateString = new Date().toLocaleString('th-TH');

    // --- ส่วนการสร้างเอกสาร PDF (แสดงรหัสและชื่อพนักงานผู้บันทึก) ---
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <style>
          @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@400;600&display=swap');
          body { font-family: 'Prompt', sans-serif; margin: 0; padding: 10px; color: #333; font-size: 14px; line-height: 1.4; }
          .header { text-align: center; border-bottom: 2px solid #2563eb; padding-bottom: 8px; margin-bottom: 12px; }
          .title { font-size: 20px; font-weight: bold; color: #1e3a8a; margin: 0; }
          .meta-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; }
          .meta-table td { padding: 6px; border: 1px solid #e5e7eb; }
          .label { font-weight: bold; background-color: #f3f4f6; width: 30%; }
          .img-container { text-align: center; margin-bottom: 10px; }
          .img-title { font-weight: bold; font-size: 13px; color: #4b5563; margin-bottom: 4px; text-align: left; }
          .report-img { width: auto; max-width: 100%; height: 260px; object-fit: contain; border: 1px solid #d1d5db; border-radius: 6px; }
          .footer { text-align: center; font-size: 11px; color: #9ca3af; margin-top: 10px; }
        </style>
      </head>
      <body>
        <div class="header">
          <p class="title">เอกสารรายงานบันทึกรายการเครื่องมือช่าง</p>
        </div>
        
        <table class="meta-table">
          <tr>
            <td class="label">รหัสทรัพย์สิน</td>
            <td style="font-size: 16px; font-weight: bold; color: #2563eb;">${cleanAssetId}</td>
          </tr>
          <tr>
            <td class="label">รหัสพนักงาน</td>
            <td style="font-weight: bold;">${payload.empId}</td>
          </tr>
          <tr>
            <td class="label">ชื่อพนักงานผู้บันทึก</td>
            <td>${payload.empName}</td>
          </tr>
          <tr>
            <td class="label">วันที่-เวลาบันทึก</td>
            <td>${saveDateString}</td>
          </tr>
        </table>

        <div class="img-container">
          <div class="img-title">1. รูปถ่ายหน้าเต็มของเครื่องมือ</div>
          <img class="report-img" src="data:image/jpeg;base64,${payload.toolImg.split(',')[1]}">
        </div>

        <div class="img-container">
          <div class="img-title">2. รูปถ่ายป้ายรหัสทรัพย์สิน</div>
          <img class="report-img" src="data:image/jpeg;base64,${payload.assetImg.split(',')[1]}">
        </div>

        <div class="footer">
          รายงานฉบับนี้สร้างขึ้นโดยระบบอัตโนมัติจากแอปพลิเคชันมือถือทีมช่าง
        </div>
      </body>
      </html>
    `;

    // แปลงรหัส HTML เป็น Blob PDF ขนาด A4 
    const htmlBlob = HtmlService.createHtmlOutput(htmlContent).getBlob();
    const pdfBlob = htmlBlob.getAs('application/pdf').setName(`REPORT_${cleanAssetId}.pdf`);
    
    // บันทึกไฟล์ PDF ลงในโฟลเดอร์ย่อย
    const pdfFile = subFolder.createFile(pdfBlob);

    // บันทึกข้อมูลประวัติลงในไฟล์ Spreadsheet (เพิ่ม emp_id และ emp_name ลงแถวบันทึกด้วย)
    targetSheet.appendRow([
      saveDateString,      // timestamp
      payload.empId,       // emppin
      payload.empName,      // emp_name
      cleanAssetId,        // tagnum
      toolFile.getUrl(),   // full_Image
      assetFile.getUrl(),  // tag_Image
      pdfFile.getUrl()    // pdf_Report

    ]);
    
    return {
      success: true,
      assetId: cleanAssetId,
      saveDate: saveDateString,
      reportUrl: pdfFile.getUrl()
    };
  } catch (error) {
    return { success: false, msg: error.toString() };
  }
}

function getFormattedDate() {
  const d = new Date();
  const yy = d.getFullYear().toString().slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${yy}${mm}${dd}${hh}${min}`;
}

// 3. ฟังก์ชันดึงประวัติรายงานของพนักงานคนนั้นๆ เรียงจากล่าสุด
function getEmployeeReports(empId) {
  try {
    const ssTool = SpreadsheetApp.openById(TOOL_SPREADSHEET_ID);
    const sheet = ssTool.getSheetByName(TOOL_SHEET_NAME);
    if (!sheet) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // หาตำแหน่งคอลัมน์จากชื่อ Header ป้องกันการสลับตำแหน่ง
    const timestampIdx = headers.indexOf('timestamp');
    const emppinIdx = headers.indexOf('emppin');
    const tagnumIdx = headers.indexOf('tagnum');
    const pdfReportIdx = headers.indexOf('pdf_Report');

    let reports = [];

    // วนลูปอ่านข้อมูล (ข้ามแถว Header แถวที่ 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // คัดกรองเอาเฉพาะรหัสพนักงานที่ตรงกัน
      if (row[emppinIdx] && row[emppinIdx].toString().trim() === empId.toString().trim()) {
        reports.push({
          timestamp: row[timestampIdx] ? row[timestampIdx].toString() : '',
          tagnum: row[tagnumIdx] ? row[tagnumIdx].toString() : '',
          pdfUrl: row[pdfReportIdx] ? row[pdfReportIdx].toString() : '#'
        });
      }
    }

    // เรียงลำดับจากวันที่ล่าสุดขึ้นก่อน (Reverse Array)
    reports.reverse();

    return { success: true, data: reports };
  } catch (error) {
    return { success: false, msg: error.toString() };
  }
}
