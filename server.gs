const ACCESS_TOKEN = "xxx";
const LINE_OA_ID = "@xxx";
const PREFILL_TEXT = "กรอกเบอร์โทรเพื่อยืนยัน : ";
const prefillURL = "https://line.me/R/oaMessage/" + LINE_OA_ID + "/?" + encodeURIComponent(PREFILL_TEXT);
const LINE_NOTIFY_TOKEN = "xxx";
const DriveId ="xxx"
const CalendarID = "xxx"


const TIME_SLOTS = {
  '08:00-12:00': { start: '08:00', end: '12:00', label: '08:00 น. - 12:00 น.' },
  '13:00-16:30': { start: '13:00', end: '16:30', label: '13:00 น. - 16:30 น.' },
  '08:00-16:30': { start: '08:00', end: '16:30', label: '08:00 น. - 16:30 น.' }
};

/**
 * Sends a message to the Line Notify service.
 * @param {string} message The message to send.
 */
function sendLineNotify(message) {
  const url = "https://notify-api.line.me/api/notify";
  const headers = {
    "Authorization": "Bearer " + LINE_NOTIFY_TOKEN
  };
  const payload = {
    "message": message
  };

  const options = {
    "method": "post",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true // Prevents script from stopping on HTTP errors
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    console.error("Error sending Line Notify message: " + e.toString());
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
        throw new Error(`การจองสำหรับ ${fieldType} ในวันที่ ${currentDate.toLocaleDateString('th-TH')} ช่วงเวลา ${timeSlot.label} ถูกจองแล้ว`);
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
      'Authorization': 'Bearer ' + ACCESS_TOKEN
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
      'Authorization': 'Bearer ' + ACCESS_TOKEN
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
    
    // For booking form
    const data = e.parameter;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 13).setValues([[
        'วันที่บันทึก', 'ชื่อ-สกุล', 'เบอร์โทรศัพท์', 'วันที่จอง (เริ่มต้น)', 'วันที่จอง (สิ้นสุด)', 
        'ช่วงเวลา', 'ประเภทสนาม', 'สถานะ', 'อุปกรณ์กีฬา', 'เจ้าหน้าที่', 'หมายเหตุ', 'สลิปการโอน', 'รหัสการจอง'
      ]]);
    }
    
    const reserveId = `BOOK_${new Date().getTime()}`;
    
    let paymentSlipUrl = '';
    
    if (data.paymentSlipData) {
      try {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(data.paymentSlipData),
          data.paymentSlipMimeType,
          data.paymentSlipName
        );
        
        const file = DriveApp.createFile(blob);
        file.setName(`สลิป_${data.fullName}_${new Date().getTime()}.${data.paymentSlipName.split('.').pop()}`);
        
        const spreadsheetFile = DriveApp.getFileById(DriveId);
        const parentFolders = spreadsheetFile.getParents();
        if (parentFolders.hasNext()) {
          const folder = parentFolders.next();
          folder.addFile(file);
          DriveApp.getRootFolder().removeFile(file);
        }
        
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        paymentSlipUrl = file.getUrl();
        
      } catch (error) {
        console.error('Error uploading file:', error);
        paymentSlipUrl = 'เกิดข้อผิดพลาดในการอัปโหลด';
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
      paymentSlipUrl,
      reserveId
    ];
    
    sheet.appendRow(newRow);
    
    // Send detailed notification via Line Notify
    const notifyMessage = `
🔔 มีการส่งฟอร์มใหม่
--------------------
ชื่อ-สกุล: ${data.fullName}
เบอร์โทรศัพท์: ${data.phone}
วันที่เริ่มต้น: ${data.bookingStartDate}
วันที่สิ้นสุด: ${data.bookingEndDate}
สถานะ: ${data.status}
ประเภทสนามที่จอง: ${data.fieldType}
ช่วงเวลา: ${TIME_SLOTS[data.timeSlot].label}
`;
    sendLineNotify(notifyMessage);
    
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

function doGet() {
  return ContentService
    .createTextOutput('Booking System API is running')
    .setMimeType(ContentService.MimeType.TEXT);
}
