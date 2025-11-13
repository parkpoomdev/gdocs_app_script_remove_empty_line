function deleteEmptyLinesOnly() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  
  var paragraphs = body.getParagraphs();
  var deletedCount = 0;
  
  // ลบจากล่างขึ้นบนเพื่อป้องกันปัญหา index
  for (var i = paragraphs.length - 1; i >= 0; i--) {
    var paragraph = paragraphs[i];
    var text = paragraph.getText().trim();
    
    // ตรวจสอบว่าเป็น empty line จริงๆ
    if (text === '') {
      // ตรวจสอบว่าไม่ใช่ paragraph สุดท้ายในเอกสาร
      if (i < paragraphs.length - 1) {
        paragraph.removeFromParent();
        deletedCount++;
      } else {
        Logger.log('ไม่สามารถลบ paragraph สุดท้ายได้');
      }
    }
  }
  
  Logger.log('ลบ empty lines ไปทั้งหมด: ' + deletedCount + ' บรรทัด');
}
