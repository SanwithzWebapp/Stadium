const LINE_ACCESS_TOKEN = "xxx"; 
const LINE_OA_ID = "@xxx"; 
const PREFILL_TEXT = "กรอกเบอร์โทรเพื่อยืนยัน : "; 
const prefillURL = "https://line.me/R/oaMessage/" + LINE_OA_ID + "/?" + encodeURIComponent(PREFILL_TEXT); 

const Line_USER_ID = "xxx"; // Assuming this is the recipient ID 
const DriveId = "xxx"; 
const CalendarID = "xxx"; 
const ADMIN_APPROVAL_URL = "https://script.google.com/macros/s/xxx/exec"; 

const TIME_SLOTS = { 
  '08:00-12:00': { start: '08:00', end: '12:00', label: '08:00 น. - 12:00 น.' }, 
  '13:00-16:30': { start: '13:00', end: '16:30', label: '13:00 น. - 16:30 น.' }, 
  '08:00-16:30': { start: '08:00', end: '16:30', label: '08:00 น. - 16:30 น.' } 
}; 


function sendLineFlexMessagePush(bookingData) { 
  // Use the provided user ID 
  const toUserId = Line_USER_ID;  
  const url = "https://api.line.me/v2/bot/message/push"; 

  const payload = { 
    to: toUserId, 
    messages: [{ 
      type: "flex", 
      altText: "มีการส่งฟอร์มใหม่", 
      contents: { 
        type: "bubble", 
        header: { 
          type: "box", 
          layout: "vertical", 
          contents: [ 
            { type: "text", text: "แจ้งเตือนการส่งฟอร์ม", weight: "bold", size: "lg" } 
          ] 
        }, 
        body: { 
          type: "box", 
          layout: "vertical", 
          spacing: "md", 
          contents: [ 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "ชื่อ-สกุล:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.fullName, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "เบอร์โทรศัพท์:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.phone, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "วันที่เริ่มต้น:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.bookingStartDate, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "วันที่สิ้นสุด:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.bookingEndDate, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "สถานะ:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.status, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "ประเภทสนามที่จอง:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: bookingData.fieldType, wrap: true, size: "sm", flex: 5 } 
              ] 
            }, 
            { 
              type: "box", 
              layout: "baseline", 
              spacing: "sm", 
              contents: [ 
                { type: "text", text: "ช่วงเวลา:", color: "#aaaaaa", size: "sm", flex: 3 }, 
                { type: "text", text: TIME_SLOTS[bookingData.timeSlot].label, wrap: true, size: "sm", flex: 5 } 
              ] 
            },
            {
              type: "box",
              layout: "baseline",
              spacing: "sm",
              contents: [
                { type: "text", text: "รหัสการจอง:", color: "#aaaaaa", size: "sm", flex: 3 },
                { type: "text", text: bookingData.bookingId, wrap: true, size: "sm", flex: 5 }
              ]
            } 
          ] 
        },
        footer: {
          type: "box",
          layout: "vertical",
          spacing: "sm",
          contents: [
            {
              type: "button",
              style: "primary",
              height: "sm",
              action: {
                type: "uri",
                label: "✅ อนุมัติ",
                uri: `${ADMIN_APPROVAL_URL}?action=approve&bookingId=${bookingData.bookingId}`
              },
              color: "#00B900"
            },
            {
              type: "button",
              style: "secondary",
              height: "sm",
              action: {
                type: "uri",
                label: "❌ ไม่อนุมัติ",
                uri: `${ADMIN_APPROVAL_URL}?action=reject&bookingId=${bookingData.bookingId}`
              },
              color: "#FF5551"
            }
          ]
        }
      } 
    }] 
  }; 

  const options = { 
    method: 'post', 
    headers: { 
      'Content-Type': 'application/json', 
      'Authorization': `Bearer ${LINE_ACCESS_TOKEN}` 
    }, 
    payload: JSON.stringify(payload) 
  }; 

  try { 
    UrlFetchApp.fetch(url, options); 
  } catch (e) { 
    console.error("Error sending LINE Flex Message: " + e.toString()); 
  } 
} 

function bookCalendar(reserveId, startDate, endDate, timeSlotKey, userName, fieldType, note) { 
  const calendarId = CalendarID; 
  const calendar = CalendarApp.getCalendarById(calendarId); 
  const timeSlot = TIME_SLOTS[timeSlotKey]; 

  if (!timeSlot) { 
    throw new Error("ช่วงเวลาไม่ถูกต้อง"); 
  } 

  const currentDate = new Date(startDate); 
  const endDateObj = new Date(endDate); 
    
  // Create a descriptive location name for the calendar event 
  let eventLocation = fieldType; 
  if (fieldType.includes('สนาม')) { 
    eventLocation = 'สนาม' + fieldType.split('สนาม')[1]; 
  } 

  while (currentDate.getTime() <= endDateObj.getTime()) { 
    const startDateTime = new Date(`${currentDate.toISOString().split('T')[0]}T${timeSlot.start}:00`); 
    const endDateTime = new Date(`${currentDate.toISOString().split('T')[0]}T${timeSlot.end}:00`); 

    // Check for conflicts on the same date and time slot 
    const events = calendar.getEvents(startDateTime, endDateTime); 
    for (let event of events) { 
      if (event.getTitle().includes(fieldType)) { 
        throw new Error(`วันที่จองไม่ว่าง\nการจองปฏิทิน: การจองสำหรับ ${fieldType} ในวันที่ ${currentDate.toLocaleDateString('th-TH')} ช่วงเวลา ${timeSlot.label} ถูกจองแล้ว`); 
      } 
    } 

    // Create the event 
    const eventDetails = `ID_booking:${reserveId}|สถานที่:${fieldType}|ผู้จอง:${userName}|หมายเหตุ:${note || "ไม่มี"}`; 
    calendar.createEvent( 
      `จอง${fieldType} โดย ${userName}`, 
      startDateTime, 
      endDateTime, 
      { 
        description: eventDetails, 
        location: eventLocation 
      } 
    ); 
      
    currentDate.setDate(currentDate.getDate() + 1); // Move to the next day 
  } 
} 

function cancelBookingByPhone(phoneNumber) { 
  try { 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
    const data = sheet.getDataRange().getValues(); 
      
    let canceledBookings = []; 
    let rowsToDelete = []; 
      
    for (let i = 1; i < data.length; i++) { 
      let cellPhone = data[i][2].toString().replace(/'/g, ''); 
      if (cellPhone === phoneNumber) { 
        canceledBookings.push({ 
          name: data[i][1], 
          date: data[i][3], 
          endDate: data[i][4], 
          time: data[i][5], 
          field: data[i][6], 
          bookingId: data[i][12] 
        }); 
        rowsToDelete.push(i + 1); 
      } 
    } 
      
    if (canceledBookings.length === 0) { 
      return { success: false, message: "ไม่พบการจองที่ใช้เบอร์โทรนี้" }; 
    } 
      
    rowsToDelete.reverse().forEach(rowIndex => { 
      sheet.deleteRow(rowIndex); 
    }); 
      
    const calendarId = CalendarID; 
    const calendar = CalendarApp.getCalendarById(calendarId); 
      
    canceledBookings.forEach(booking => { 
      if (booking.bookingId) { 
        const events = calendar.getEvents(new Date(booking.date), new Date(new Date(booking.endDate).getTime() + 24*60*60*1000)); 
        events.forEach(event => { 
          if (event.getDescription().includes(`ID_booking:${booking.bookingId}`)) { 
            event.deleteEvent(); 
          } 
        }); 
      } 
    }); 
      
    return {  
      success: true,  
      message: `ยกเลิกการจองเรียบร้อยแล้ว \nจำนวน ${canceledBookings.length} รายการ`, 
      canceledBookings: canceledBookings 
    }; 
      
  } catch (error) { 
    console.error('Error canceling booking:', error); 
    return { success: false, message: "เกิดข้อผิดพลาดในการยกเลิกการจอง" }; 
  } 
} 

function replyToLine(replyToken, message) { 
  const url = 'https://api.line.me/v2/bot/message/reply'; 
  const payload = { 
    replyToken: replyToken, 
    messages: [{ 
      type: 'text', 
      text: message 
    }] 
  }; 
    
  const options = { 
    method: 'POST', 
    headers: { 
      'Content-Type': 'application/json', 
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN 
    }, 
    payload: JSON.stringify(payload) 
  }; 
    
  UrlFetchApp.fetch(url, options); 
} 

function replyFlexMessage(replyToken, flexContent) { 
  const url = 'https://api.line.me/v2/bot/message/reply'; 
  const payload = { 
    replyToken: replyToken, 
    messages: [{ 
      type: 'flex', 
      altText: 'ยกเลิกการจอง', 
      contents: flexContent 
    }] 
  }; 
    
  const options = { 
    method: 'POST', 
    headers: { 
      'Content-Type': 'application/json', 
      'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN 
}, 
    payload: JSON.stringify(payload) 
  }; 
    
  UrlFetchApp.fetch(url, options); 
} 

function createSuccessFlexMessage(result) { 
  let bookingContents = []; 
    
  if (result.canceledBookings && result.canceledBookings.length > 0) { 
    result.canceledBookings.forEach((booking, index) => { 
      if (index > 0) { 
        bookingContents.push({ 
          type: "separator", 
          margin: "md" 
        }); 
      } 
        
      const timeLabel = TIME_SLOTS[booking.time] ? TIME_SLOTS[booking.time].label : booking.time; 

      bookingContents.push({ 
        type: "box", 
        layout: "vertical", 
        contents: [ 
          { 
            type: "text", 
            text: `${index + 1}. ${booking.name}`, 
            weight: "bold", 
            size: "sm", 
            color: "#333333" 
          }, 
          { 
            type: "text", 
            text: `🏟️ ${booking.field}`, 
            size: "xs", 
            color: "#666666", 
            margin: "xs" 
          }, 
          { 
            type: "text", 
            text: `📅 ${booking.date} - ${booking.endDate}`, 
            size: "xs", 
            color: "#666666", 
            margin: "xs" 
          }, 
          { 
            type: "text", 
            text: `⏰ ${timeLabel}`, 
            size: "xs", 
            color: "#666666", 
            margin: "xs" 
          } 
        ], 
        backgroundColor: "#F5F5F5", 
        cornerRadius: "8px", 
        paddingAll: "12px", 
        margin: "sm" 
      }); 
    }); 
  } 

  return { 
    type: "bubble", 
    header: { 
      type: "box", 
      layout: "vertical", 
      contents: [ 
        { 
          type: "text", 
          text: "✅ ยกเลิกเรียบร้อย", 
          weight: "bold", 
          size: "xl", 
          color: "#FFFFFF", 
          align: "center" 
        } 
      ], 
      backgroundColor: "#06C755", 
      paddingAll: "20px" 
    }, 
    body: { 
      type: "box", 
      layout: "vertical", 
      contents: [ 
        { 
          type: "text", 
          text: result.message, 
          size: "md", 
          weight: "bold", 
          color: "#333333", 
          wrap: true 
        }, 
        { 
          type: "separator", 
          margin: "xl" 
        }, 
        { 
          type: "text", 
          text: "📋 รายการที่ยกเลิก", 
          size: "sm", 
          weight: "bold", 
          color: "#666666", 
          margin: "xl" 
        }, 
        ...bookingContents 
      ], 
      paddingAll: "20px" 
    }, 
    footer: { 
      type: "box", 
      layout: "vertical", 
      contents: [ 
        { 
          type: "text", 
          text: "ขอบคุณที่ใช้บริการ 🙏", 
          size: "xs", 
          color: "#999999", 
          align: "center" 
        } 
      ], 
      paddingAll: "15px" 
    } 
  }; 
} 

/**
 * Handles incoming HTTP GET requests, specifically for admin approval.
 * @param {Object} e The event parameter that contains data from the request.
 */
function doGet(e) {
  const params = e.parameter;
  const bookingId = params.bookingId;
  const action = params.action;

  // Check if the request is an admin action (approve or reject)
  if (bookingId && (action === 'approve' || action === 'reject')) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    let rowFound = false;

    // Iterate through the rows to find the booking ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][12] === bookingId) { // Column M (index 12) contains the Booking ID
        
        // If the action is 'reject', also delete the calendar event
        if (action === 'reject') {
          try {
            const calendar = CalendarApp.getCalendarById(CalendarID);
            const bookingStartDate = new Date(data[i][3]); // Column D (index 3)
            const bookingEndDate = new Date(data[i][4]); // Column E (index 4)
            
            const events = calendar.getEvents(bookingStartDate, new Date(bookingEndDate.getTime() + 24 * 60 * 60 * 1000));
            
            for (let event of events) {
              if (event.getDescription().includes(`ID_booking:${bookingId}`)) {
                event.deleteEvent();
                break; // Event found and deleted, no need to check others
              }
            }
          } catch (calendarError) {
            console.error('Error deleting calendar event:', calendarError);
          }
        }
        
        const status = (action === 'approve') ? 'อนุมัติแล้ว' : 'ปฏิเสธแล้ว';
        // Update the value in column N (index 13) with the new status
        sheet.getRange(i + 1, 14).setValue(status); 
        rowFound = true;
        break;
      }
    }

    // Return a simple HTML response to the admin's browser
    if (rowFound) {
      const message = (action === 'approve') ? 'การจองนี้ถูกอนุมัติแล้ว' : 'การจองนี้ถูกปฏิเสธแล้ว';
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html>
        <html lang="th">
        <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>การอัปเดตสถานะ</title>
          <style>
            body { font-family: sans-serif; text-align: center; padding: 50px; }
            .success { color: #06C755; }
            .error { color: #FF5551; }
          </style>
        </head>
        <body>
          <h1 class="${action === 'approve' ? 'success' : 'error'}">${message}</h1>
          <p>ระบบได้อัปเดตสถานะใน Google Sheets เรียบร้อยแล้ว</p>
        </body>
        </html>
      `);
    } else {
      return ContentService.createTextOutput('ไม่พบรหัสการจอง');
    }
  }

  // Original doGet function for fetching data
  const s = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getValues();
  const h = SpreadsheetApp.getActive().getSheets()[0].getDataRange().getValues()[0]?.map(String) || [];
  const d = s.length>1 ? s.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,r[i]]))) : [];
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { 
  try { 
    const contentType = e.postData.type; 
      
    if (contentType === 'application/json') { 
      const lineData = JSON.parse(e.postData.contents); 
        
      if (lineData.events && lineData.events.length > 0) { 
        const event = lineData.events[0]; 
          
        if (event.type === 'message' && event.message.type === 'text') { 
          const messageText = event.message.text.trim(); 
          const replyToken = event.replyToken; 
            
          if (messageText === 'ยกเลิกการจอง') { 
            const flexContent = { 
              type: "bubble", 
              header: { 
                type: "box", 
                layout: "vertical", 
                contents: [ 
                  { 
                    type: "text", 
                    text: "ยกเลิกการจอง", 
                    weight: "bold", 
                    size: "xl", 
                    color: "#FF5551", 
                    align: "center" 
                  } 
                ], 
                backgroundColor: "#FFE5E5", 
                paddingAll: "20px" 
              }, 
              body: { 
                type: "box", 
                layout: "vertical", 
                contents: [ 
                  { 
                    type: "text", 
                    text: "📱 กรุณากรอกเบอร์โทรศัพท์", 
                    size: "lg", 
                    weight: "bold", 
                    color: "#333333", 
                    margin: "md" 
                  }, 
                  { 
                    type: "text", 
                    text: "เพื่อยืนยันการยกเลิกการจองสนาม", 
                    size: "sm", 
                    color: "#666666", 
                    wrap: true, 
                    margin: "sm" 
                  }, 
                  { 
                    type: "separator", 
                    margin: "xl" 
                  }, 
                  { 
                    type: "box", 
                    layout: "vertical", 
                    contents: [ 
                      { 
                        type: "text", 
                        text: "⚠️ หมายเหตุ", 
                        size: "sm", 
                        weight: "bold", 
                        color: "#FF9500", 
                        margin: "md" 
                      }, 
                      { 
                        type: "text", 
                        text: "• ระบบจะยกเลิกการจองทั้งหมดที่ใช้เบอร์นี้\n• การยกเลิกไม่สามารถย้อนกลับได้", 
                        size: "xs", 
                        color: "#999999", 
                        wrap: true, 
                        margin: "sm" 
                      } 
                    ] 
                  } 
                ], 
                paddingAll: "20px" 
              }, 
              footer: { 
                type: "box", 
                layout: "vertical", 
                contents: [ 
                  { 
                    type: "button", 
                    style: "primary", 
                    height: "sm", 
                    action: { 
                      type: "uri", 
                      label: "📞 กรอกเบอร์โทร", 
                      uri: prefillURL 
                    }, 
                    color: "#FF5551" 
                  } 
                ], 
                paddingAll: "20px" 
              } 
            }; 
              
            replyFlexMessage(replyToken, flexContent); 
              
          } else if (messageText.startsWith('กรอกเบอร์โทรเพื่อยืนยัน : ')) { 
            const phoneNumber = messageText.replace('กรอกเบอร์โทรเพื่อยืนยัน : ', '').trim(); 
              
            if (phoneNumber && phoneNumber.length >= 9) { 
              const result = cancelBookingByPhone(phoneNumber); 
                
              if (result.success) { 
                const successFlexContent = createSuccessFlexMessage(result); 
                replyFlexMessage(replyToken, successFlexContent); 
              } else { 
                replyToLine(replyToken, result.message); 
              } 
            } else { 
              replyToLine(replyToken, 'กรุณากรอกเบอร์โทรศัพท์ให้ถูกต้อง'); 
            } 
          } 
        } 
          
        return ContentService.createTextOutput('OK'); 
      } 
    } 
      
    const data = e.parameter; 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
      
    // Set up the header row if it doesn't exist
    if (sheet.getLastRow() === 0) { 
      sheet.getRange(1, 1, 1, 14).setValues([[ 
        'วันที่บันทึก', 'ชื่อ-สกุล', 'เบอร์โทรศัพท์', 'วันที่จอง (เริ่มต้น)', 'วันที่จอง (สิ้นสุด)',  
        'ช่วงเวลา', 'ประเภทสนาม', 'สถานะ', 'อุปกรณ์กีฬา', 'เจ้าหน้าที่', 'หมายเหตุ', 'เอกสารแนบ', 'รหัสการจอง', 'สถานะการจอง' 
      ]]); 
    } 
      
    const reserveId = `BOOK_${new Date().getTime()}`; 
    data.bookingId = reserveId; // Add the bookingId to the data object
      
    let uploadedFileUrl = ''; 
      
    if (data.fileData) { 
      try { 
        const blob = Utilities.newBlob( 
          Utilities.base64Decode(data.fileData), 
          data.fileMimeType, 
          data.fileName 
        ); 
        
        // Define a more descriptive file name based on the file type
        let filePrefix = (data.fileType === 'paymentSlip') ? 'สลิป' : 'เอกสารอนุมัติ';
        let newFileName = `${filePrefix}_${data.fullName}_${new Date().getTime()}.${data.fileName.split('.').pop()}`;

        const file = DriveApp.createFile(blob); 
        file.setName(newFileName); 
          
        const spreadsheetFile = DriveApp.getFileById(DriveId); 
        const parentFolders = spreadsheetFile.getParents(); 
        if (parentFolders.hasNext()) { 
          const folder = parentFolders.next(); 
          folder.addFile(file); 
          DriveApp.getRootFolder().removeFile(file); 
        } 
          
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); 
        uploadedFileUrl = file.getUrl(); 
          
      } catch (error) { 
        console.error('Error uploading file:', error); 
        uploadedFileUrl = 'เกิดข้อผิดพลาดในการอัปโหลด'; 
      } 
    } 
      
    try { 
      bookCalendar( 
        reserveId, 
        data.bookingStartDate, 
        data.bookingEndDate, 
        data.timeSlot, 
        data.fullName, 
        data.fieldType, 
        data.notes 
      ); 
    } catch (calendarError) { 
      throw new Error(`การจองปฏิทิน: ${calendarError.message}`); 
    } 
      
    const newRow = [ 
      new Date(), 
      data.fullName || '', 
      "'"+data.phone || '', 
      data.bookingStartDate || '', 
      data.bookingEndDate || '', 
      TIME_SLOTS[data.timeSlot].label || '', 
      data.fieldType || '', 
      data.status || '', 
      data.equipment || '', 
      data.staff || '', 
      data.notes || '', 
      uploadedFileUrl, 
      reserveId,
      'รอดำเนินการ' // Set the initial status
    ]; 
      
    sheet.appendRow(newRow); 
      
    sendLineFlexMessagePush(data); 

    return ContentService 
      .createTextOutput(JSON.stringify({ 
        success: true, 
        message: 'บันทึกข้อมูลสำเร็จ', 
        bookingId: reserveId 
      })) 
      .setMimeType(ContentService.MimeType.JSON); 
        
  } catch (error) { 
    console.error('Error:', error); 
    return ContentService 
      .createTextOutput(JSON.stringify({ 
        success: false, 
        message: 'เกิดข้อผิดพลาด: ' + error.toString() 
      })) 
      .setMimeType(ContentService.MimeType.JSON); 
  } 
}

