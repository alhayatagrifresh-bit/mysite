/** Northstar Academy admission management API. Deploy as a Web App. */
const SHEET_STUDENTS = 'Students';
const SHEET_FEES = 'Fee_Settings';
const SHEET_ADMINS = 'Admin_Credentials';
const SHEET_BATCHES = 'Batch_Settings';
const SHEET_DELETED = 'Deleted_Students';
const STUDENT_HEADERS = ['Timestamp','Form_Number','Roll_Number','Full_Name','Father_Name','Mother_Name','DOB','Gender','Mobile','Email','Address','Class','School','Photo_Drive_URL','Fee_Amount','Payment_Status','Transaction_ID','Aadhaar_Number','Batch'];
const FEE_HEADERS = ['Class_Name','Form_Fee_Amount','Last_Assigned_Roll_No','Active','Roll_Prefix'];
const ADMIN_HEADERS = ['Username','Password_Hash','Last_Login'];
const CLASSES = ['Class 6th','Class 7th','Class 8th','Class 9th','Class 10th','Class 11th','Class 12th','JEE','NEET'];
const DEFAULT_BATCHES = ['Morning · 7:00 AM','Afternoon · 1:00 PM','Evening · 5:00 PM','Weekend'];

function doGet(e) {
  // Both routes serve the SPA; the page selector keeps URLs shareable.
  const page = HtmlService.createTemplateFromFile('index').evaluate();
  return page
    .setTitle('Northstar Academy · Admissions')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let result;
    switch (req.action) {
      case 'adminLogin': result = adminLogin_(req); break;
      case 'getFeeData': result = {fees:feeMap_(),classes:Object.keys(feeMap_()),batches:getBatches_()}; break;
      case 'saveCatalog': requireSession_(req.token); result = saveCatalog_(req.classes, req.batches); break;
      case 'getAdminUsers': requireSession_(req.token); result = {users:getAdminUsers_()}; break;
      case 'saveAdminUser': requireSession_(req.token); result = saveAdminUser_(req.username, req.password, req.oldUsername); break;
      case 'saveRazorpayKeys': requireSession_(req.token); result = saveRazorpayKeys_(req.keyId, req.keySecret); break;
      case 'getRazorpayConfig': requireSession_(req.token); result = getRazorpayConfig_(); break;
      case 'getPaymentMethods': result = getPaymentMethods_(); break;
      case 'savePaymentMethods': requireSession_(req.token); result = savePaymentMethods_(req.settings); break;
      case 'getInstituteSettings': result = getInstituteSettings_(); break;
      case 'saveInstituteSettings': requireSession_(req.token); result = saveInstituteSettings_(req.settings); break;
      case 'getRollSeries': requireSession_(req.token); result = {series:getRollSeries_()}; break;
      case 'saveRollSeries': requireSession_(req.token); result = {series:saveRollSeries_(req.series)}; break;
      case 'getAdminData': requireSession_(req.token); result = getAdminData_(); break;
      case 'updateClassFees': requireSession_(req.token); result = updateClassFees_(req.fees); break;
      case 'createPaymentOrder': result = createPaymentOrder_(req); break;
      case 'submitForm': result = {student:submitForm(req.formData, req.payment, req.photoData, req.paymentMode, req.manualReference)}; break;
      case 'confirmManualPayment': requireSession_(req.token); result = confirmManualPayment_(req.formNumber); break;
      case 'updateStudent': requireSession_(req.token); result = updateStudent_(req.student); break;
      case 'deleteStudent': requireSession_(req.token); result = deleteStudent_(req.formNumber); break;
      default: throw new Error('Unknown action.');
    }
    return json_({ok:true, ...result});
  } catch (err) {
    return json_({ok:false, error:String(err.message || err)});
  }
}

/** Run once from the Apps Script editor. Creates tabs, Drive folder and fee rows. */
function initializeApp() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const spreadsheetId = props.getProperty('SPREADSHEET_ID');
  if (spreadsheetId) ss = SpreadsheetApp.openById(spreadsheetId);
  else { ss = SpreadsheetApp.create('Northstar Academy · Admissions'); props.setProperty('SPREADSHEET_ID', ss.getId()); }
  ensureSheet_(ss, SHEET_STUDENTS, STUDENT_HEADERS);
  const feeSheet = ensureSheet_(ss, SHEET_FEES, FEE_HEADERS);
  ensureMissingHeaders_(feeSheet, FEE_HEADERS);
  ensureSheet_(ss, SHEET_ADMINS, ADMIN_HEADERS);
  const batchSheet = ensureSheet_(ss, SHEET_BATCHES, ['Batch_Name']);
  ensureMissingHeaders_(ss.getSheetByName(SHEET_STUDENTS), STUDENT_HEADERS);
  if (feeSheet.getLastRow() < 2) feeSheet.getRange(2,1,CLASSES.length,5).setValues(CLASSES.map((c,i)=>[c,[500,500,600,700,800,1000,1200,1500,1500][i],100,true,classPrefix_(c)]));
  if (batchSheet.getLastRow() < 2) batchSheet.getRange(2,1,DEFAULT_BATCHES.length,1).setValues(DEFAULT_BATCHES.map(b=>[b]));
  const deleted=ensureSheet_(ss,SHEET_DELETED,STUDENT_HEADERS.concat(['Deleted_At']));
  ensureMissingHeaders_(deleted,STUDENT_HEADERS.concat(['Deleted_At']));
  if (!props.getProperty('PHOTO_FOLDER_ID')) {
    const folder = DriveApp.createFolder('Northstar Academy · Student Photos');
    props.setProperty('PHOTO_FOLDER_ID', folder.getId());
  }
  Logger.log('Spreadsheet: ' + ss.getUrl());
  return ss.getUrl();
}

/** Run manually after initializeApp. Use a strong password and remove it from editor history. */
function setInitialAdmin(username, password) {
  if (!username || !password || password.length < 12) throw new Error('Provide an email and a password of at least 12 characters.');
  const sh = sheet_(SHEET_ADMINS);
  const values = sh.getDataRange().getValues();
  const hash = sha256_(password);
  const row = values.findIndex((r,i)=>i>0 && String(r[0]).toLowerCase()===String(username).toLowerCase());
  if (row > 0) sh.getRange(row+1,2).setValue(hash);
  else sh.appendRow([username,hash,'']);
}

function uploadFileToDrive(base64Data, filename) {
  if (!base64Data) return '';
  const match = String(base64Data).match(/^data:(image\/(?:png|jpeg));base64,(.+)$/);
  if (!match) throw new Error('Photo must be a PNG or JPEG image.');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 2*1024*1024) throw new Error('Photo must be 2 MB or smaller.');
  const folderId = PropertiesService.getScriptProperties().getProperty('PHOTO_FOLDER_ID');
  if (!folderId) throw new Error('Run initializeApp() before uploading photos.');
  const safeName = String(filename || 'student-photo').replace(/[^a-zA-Z0-9._-]/g,'_').slice(0,100);
  const blob = Utilities.newBlob(bytes, match[1], safeName);
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function generateRollNumber(className) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { return generateRollNumberLocked_(className); }
  finally { lock.releaseLock(); }
}
function generateRollNumberLocked_(className) {
  const sh = sheet_(SHEET_FEES), rows = sh.getDataRange().getValues();
  const index = rows.findIndex((r,i)=>i>0 && String(r[0])===String(className));
  if (index < 1) throw new Error('Unknown class or course.');
  if (rows[index][3] === false) throw new Error('This class is inactive.');
  const current = Number(rows[index][2]) || 100, next = current + 1;
  sh.getRange(index+1,3).setValue(next);
  return String(rows[index][4]||classPrefix_(className)) + '-' + next;
}

function submitForm(formData, payment, photoData, paymentMode, manualReference) {
  const d=formData||{}, p=payment||{};
  if (!d.fullName || !d.fatherName || !d.mobile || !d.className) throw new Error('Complete the required student details.');
  if (!/^[0-9]{10}$/.test(String(d.mobile))) throw new Error('Enter a valid 10-digit mobile number.');
  if (d.aadhaarNumber && !/^[0-9]{12}$/.test(String(d.aadhaarNumber))) throw new Error('Aadhaar number must contain 12 digits.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const fees=feeMap_(), amount=Number(fees[d.className]);
    if (!amount || amount<1) throw new Error('Fee is not configured for this class.');
    const manual=paymentMode==='manual';
    if(manual){if(!String(manualReference||'').trim()||String(manualReference).trim().length>100)throw new Error('Enter the payment transaction reference.');}
    else verifyRazorpayPayment_(p, amount);
    const photoUrl=photoData?uploadFileToDrive(photoData,d.photoFilename):'';
    const roll=generateRollNumberLocked_(d.className);
    const year=new Date().getFullYear();
    const formNo='ADM-'+year+'-'+Utilities.getUuid().replace(/-/g,'').slice(0,8).toUpperCase();
    const record=[new Date(),formNo,roll,d.fullName,d.fatherName,d.motherName||'',d.dob||'',d.gender||'',String(d.mobile),d.email||'',d.address||'',d.className,d.school||'',photoUrl,amount,manual?'Pending verification':'Paid',manual?String(manualReference).trim():p.razorpay_payment_id,d.aadhaarNumber||'',d.batch||''];
    sheet_(SHEET_STUDENTS).appendRow(record);
    return studentObject_(record);
  } finally { lock.releaseLock(); }
}

function createPaymentOrder_(req) {
  const fee=feeMap_()[req.className]; if (!fee) throw new Error('Fee is not configured for this class.');
  const props=PropertiesService.getScriptProperties(), keyId=props.getProperty('RAZORPAY_KEY_ID'), secret=props.getProperty('RAZORPAY_KEY_SECRET');
  if (!keyId || !secret) throw new Error('Razorpay keys are not configured in Apps Script properties.');
  const auth=Utilities.base64Encode(keyId+':'+secret);
  const response=UrlFetchApp.fetch('https://api.razorpay.com/v1/orders',{method:'post',contentType:'application/json',headers:{Authorization:'Basic '+auth},payload:JSON.stringify({amount:Math.round(Number(fee)*100),currency:'INR',receipt:'adm_'+Utilities.getUuid().slice(0,20),notes:{class_name:req.className}}),muteHttpExceptions:true});
  const order=JSON.parse(response.getContentText());
  if (response.getResponseCode()>=300) throw new Error(order.error&&order.error.description||'Could not create payment order.');
  return {keyId:keyId,orderId:order.id,amount:order.amount,currency:order.currency};
}

function verifyRazorpayPayment_(p, amount) {
  const props=PropertiesService.getScriptProperties(), secret=props.getProperty('RAZORPAY_KEY_SECRET');
  if (!secret || !p.razorpay_order_id || !p.razorpay_payment_id || !p.razorpay_signature) throw new Error('Payment verification details are missing.');
  const expected=Utilities.computeHmacSha256Signature(p.razorpay_order_id+'|'+p.razorpay_payment_id,secret).map(b=>(b+256)%256).map(b=>('0'+b.toString(16)).slice(-2)).join('');
  if (!constantTimeEquals_(expected,p.razorpay_signature)) throw new Error('Payment signature verification failed.');
  const auth=Utilities.base64Encode(props.getProperty('RAZORPAY_KEY_ID')+':'+secret);
  const resp=UrlFetchApp.fetch('https://api.razorpay.com/v1/payments/'+encodeURIComponent(p.razorpay_payment_id),{headers:{Authorization:'Basic '+auth},muteHttpExceptions:true});
  const payment=JSON.parse(resp.getContentText());
  if (resp.getResponseCode()>=300 || payment.order_id!==p.razorpay_order_id || payment.status!=='captured' || Number(payment.amount)!==Math.round(amount*100)) throw new Error('Payment is not captured or does not match the configured fee.');
}

function getAdminData() {
  const students=sheet_(SHEET_STUDENTS).getDataRange().getValues().slice(1).filter(r=>r[1]).map(studentObject_);
  const fees=feeMap_();
  return {students:students,fees:fees,classes:Object.keys(fees),batches:getBatches_(),rollSeries:getRollSeries_(),institute:getInstituteSettings_(),stats:{total:students.length,feesCollected:students.filter(s=>s.paymentStatus==='Paid').reduce((a,s)=>a+Number(s.feeAmount||0),0)}};
}
function getAdminData_(){return getAdminData();}

function updateClassFees(feeData) { return updateClassFees_(feeData); }
function updateClassFees_(feeData) {
  if (!feeData || typeof feeData!=='object') throw new Error('Fee data is required.');
  const sh=sheet_(SHEET_FEES), rows=sh.getDataRange().getValues();
  Object.keys(feeData).forEach(name=>{const i=rows.findIndex((r,n)=>n>0&&r[0]===name);if(i<1)throw new Error('Unknown class: '+name);const value=Number(feeData[name]);if(!Number.isFinite(value)||value<0)throw new Error('Fee must be a non-negative amount.');sh.getRange(i+1,2).setValue(value)});
  return {fees:feeMap_()};
}

function saveCatalog_(classList, batchList) {
  const classes=cleanUniqueNames_(classList), batches=cleanUniqueNames_(batchList);
  if(!classes.length) throw new Error('Keep at least one class or course.');
  if(!batches.length) throw new Error('Keep at least one batch.');
  const sh=sheet_(SHEET_FEES), old=sh.getDataRange().getValues().slice(1), byName={};
  old.forEach(r=>{if(r[0])byName[String(r[0])]=[Number(r[1])||0,Number(r[2])||100,String(r[4]||classPrefix_(r[0]))]});
  const allNames=[...new Set([...old.map(r=>String(r[0]||'')).filter(Boolean),...classes])];
  sh.clearContents(); sh.getRange(1,1,1,FEE_HEADERS.length).setValues([FEE_HEADERS]);
  sh.getRange(1,1,1,FEE_HEADERS.length).setFontWeight('bold').setBackground('#eaf0ff');
  sh.getRange(2,1,allNames.length,5).setValues(allNames.map(c=>[c,...(byName[c]||[0,100,classPrefix_(c)]).slice(0,2),classes.some(n=>n.toLowerCase()===c.toLowerCase()),(byName[c]||[0,100,classPrefix_(c)])[2]]));
  const bs=sheet_(SHEET_BATCHES); bs.clearContents(); bs.getRange(1,1).setValue('Batch_Name').setFontWeight('bold').setBackground('#eaf0ff');
  bs.getRange(2,1,batches.length,1).setValues(batches.map(b=>[b]));
  return {classes:classes,batches:batches,fees:feeMap_()};
}

function getAdminUsers_(){return sheet_(SHEET_ADMINS).getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({username:String(r[0]),lastLogin:r[2]||''}));}
function saveAdminUser_(username,password,oldUsername){
  const name=String(username||'').trim().toLowerCase();
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(name))throw new Error('Enter a valid admin email address.');
  if(password && String(password).length<12)throw new Error('Password must be at least 12 characters.');
  const sh=sheet_(SHEET_ADMINS), rows=sh.getDataRange().getValues(), old=String(oldUsername||'').trim().toLowerCase();
  const row=rows.findIndex((r,i)=>i>0&&String(r[0]).toLowerCase()===old);
  const duplicate=rows.findIndex((r,i)=>i>0&&String(r[0]).toLowerCase()===name&&i!==row);
  if(duplicate>0)throw new Error('An admin with that email already exists.');
  if(row>0){sh.getRange(row+1,1).setValue(name);if(password)sh.getRange(row+1,2).setValue(sha256_(password));}
  else {if(!password)throw new Error('A password is required for a new admin.');sh.appendRow([name,sha256_(password),'']);}
  return {users:getAdminUsers_()};
}

function saveRazorpayKeys_(keyId,keySecret){
  const id=String(keyId||'').trim(), secret=String(keySecret||'').trim();
  if(!/^rzp_(test|live)_[A-Za-z0-9]+$/.test(id))throw new Error('The Key ID format is not recognized.');
  if(secret.length<16)throw new Error('The Key Secret appears incomplete.');
  const props=PropertiesService.getScriptProperties();
  props.setProperties({RAZORPAY_KEY_ID:id,RAZORPAY_KEY_SECRET:secret});
  return getRazorpayConfig_();
}
function getRazorpayConfig_(){
  const p=PropertiesService.getScriptProperties(),id=p.getProperty('RAZORPAY_KEY_ID')||'';
  return {configured:!!id,keyId:id? id.slice(0,9)+'••••'+id.slice(-4):'',mode:id.indexOf('rzp_live_')===0?'live':'test'};
}

function getPaymentMethods_(){
  const p=PropertiesService.getScriptProperties();
  return {upiId:p.getProperty('ADMISSION_UPI_ID')||'',payeeName:p.getProperty('ADMISSION_UPI_PAYEE')||'',bankName:p.getProperty('ADMISSION_BANK_NAME')||'',accountName:p.getProperty('ADMISSION_BANK_ACCOUNT_NAME')||'',accountNumber:p.getProperty('ADMISSION_BANK_ACCOUNT_NUMBER')||'',ifsc:p.getProperty('ADMISSION_BANK_IFSC')||'',branch:p.getProperty('ADMISSION_BANK_BRANCH')||''};
}
function savePaymentMethods_(settings){
  const s=settings||{}, clean={
    ADMISSION_UPI_ID:String(s.upiId||'').trim(),ADMISSION_UPI_PAYEE:String(s.payeeName||'').trim(),
    ADMISSION_BANK_NAME:String(s.bankName||'').trim(),ADMISSION_BANK_ACCOUNT_NAME:String(s.accountName||'').trim(),
    ADMISSION_BANK_ACCOUNT_NUMBER:String(s.accountNumber||'').trim(),ADMISSION_BANK_IFSC:String(s.ifsc||'').trim().toUpperCase(),
    ADMISSION_BANK_BRANCH:String(s.branch||'').trim()
  };
  const hasUpi=!!clean.ADMISSION_UPI_ID,hasBank=!!clean.ADMISSION_BANK_ACCOUNT_NUMBER;
  if(!hasUpi&&!hasBank)throw new Error('Provide a UPI ID or bank account number.');
  if(hasUpi&&!/^[\w.\-]{2,256}@[\w]{2,64}$/.test(clean.ADMISSION_UPI_ID))throw new Error('Enter a valid UPI ID.');
  if(hasBank&&(!clean.ADMISSION_BANK_NAME||!clean.ADMISSION_BANK_ACCOUNT_NAME||!clean.ADMISSION_BANK_IFSC))throw new Error('Bank name, account holder and IFSC are required with an account number.');
  PropertiesService.getScriptProperties().setProperties(clean);
  return {settings:getPaymentMethods_()};
}
function confirmManualPayment_(formNumber){
  const sh=sheet_(SHEET_STUDENTS),rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&String(r[1])===String(formNumber));
  if(i<1)throw new Error('Student registration not found.');
  if(String(rows[i][15])==='Pending verification')sh.getRange(i+1,16).setValue('Paid');
  return {student:studentObject_(sh.getRange(i+1,1,1,STUDENT_HEADERS.length).getValues()[0])};
}

function getInstituteSettings_(){
  const p=PropertiesService.getScriptProperties();
  return {name:p.getProperty('INSTITUTE_NAME')||'Northstar Academy',address:p.getProperty('INSTITUTE_ADDRESS')||'',phone:p.getProperty('INSTITUTE_PHONE')||'+91 1800 123 456'};
}
function saveInstituteSettings_(settings){
  const s=settings||{}, name=String(s.name||'').trim(),address=String(s.address||'').trim(),phone=String(s.phone||'').trim();
  if(!name||name.length>120)throw new Error('Enter an institute name up to 120 characters.');
  if(!address||address.length>300)throw new Error('Enter an institute address up to 300 characters.');
  if(!/^\+?[0-9\s()\-]{7,20}$/.test(phone))throw new Error('Enter a valid institute phone number.');
  PropertiesService.getScriptProperties().setProperties({INSTITUTE_NAME:name,INSTITUTE_ADDRESS:address,INSTITUTE_PHONE:phone});
  return {settings:getInstituteSettings_()};
}
function getRollSeries_(){
  const rows=sheet_(SHEET_FEES).getDataRange().getValues();
  return rows.slice(1).filter(r=>r[0]&&r[3]!==false).map(r=>({className:String(r[0]),prefix:String(r[4]||classPrefix_(r[0])),nextNumber:(Number(r[2])||100)+1}));
}
function saveRollSeries_(series){
  if(!series||typeof series!=='object')throw new Error('Roll series data is required.');
  const sh=sheet_(SHEET_FEES), rows=sh.getDataRange().getValues(), lock=LockService.getScriptLock();lock.waitLock(30000);
  try{Object.keys(series).forEach(name=>{const i=rows.findIndex((r,n)=>n>0&&r[0]===name);if(i<1||rows[i][3]===false)throw new Error('Unknown or inactive class: '+name);const prefix=String(series[name].prefix||'').trim(),next=Number(series[name].nextNumber);if(!/^[A-Za-z0-9-]{1,12}$/.test(prefix))throw new Error('Use 1–12 letters, digits or hyphens for each roll prefix.');if(!Number.isSafeInteger(next)||next<1)throw new Error('Next roll number must be a positive whole number.');sh.getRange(i+1,3).setValue(next-1);sh.getRange(i+1,5).setValue(prefix)});return getRollSeries_()}finally{lock.releaseLock()}
}
function updateStudent_(student){
  const d=student||{},sh=sheet_(SHEET_STUDENTS),lock=LockService.getScriptLock();lock.waitLock(30000);
  try{const rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&String(r[1])===String(d.formNumber));if(i<1)throw new Error('Student registration not found.');if(!d.fullName||!d.fatherName||!d.className)throw new Error('Name, father name and class are required.');if(!/^[0-9]{10}$/.test(String(d.mobile||'')))throw new Error('Enter a valid 10-digit mobile number.');if(d.aadhaarNumber&&!/^[0-9]{12}$/.test(String(d.aadhaarNumber)))throw new Error('Aadhaar number must contain 12 digits.');const oldClass=String(rows[i][11]);if(oldClass!==d.className){if(!Object.prototype.hasOwnProperty.call(feeMap_(),d.className))throw new Error('Select an active class.');sh.getRange(i+1,3).setValue(generateRollNumberLocked_(d.className));}const vals=[d.fullName,d.fatherName,d.motherName||'',d.dob||'',d.gender||'',String(d.mobile),d.email||'',d.address||'',d.className,d.school||''];sh.getRange(i+1,4,1,10).setValues([vals]);sh.getRange(i+1,18).setValue(d.aadhaarNumber||'');sh.getRange(i+1,19).setValue(d.batch||'');return {student:studentObject_(sh.getRange(i+1,1,1,STUDENT_HEADERS.length).getValues()[0])}}finally{lock.releaseLock()}
}
function deleteStudent_(formNumber){
  const sh=sheet_(SHEET_STUDENTS),lock=LockService.getScriptLock();lock.waitLock(30000);
  try{const rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&String(r[1])===String(formNumber));if(i<1)throw new Error('Student registration not found.');const ss=SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));const archive=ensureSheet_(ss,SHEET_DELETED,STUDENT_HEADERS.concat(['Deleted_At']));const record=rows[i].slice(0,STUDENT_HEADERS.length);record.push(new Date());archive.appendRow(record);sh.deleteRow(i+1);return {deleted:true,formNumber:String(formNumber)}}finally{lock.releaseLock()}
}

function adminLogin_(req) {
  const rows=sheet_(SHEET_ADMINS).getDataRange().getValues(), user=String(req.email||'').toLowerCase(), hash=sha256_(String(req.password||''));
  const i=rows.findIndex((r,n)=>n>0&&String(r[0]).toLowerCase()===user&&constantTimeEquals_(String(r[1]),hash));
  if(i<1)throw new Error('Invalid email or password.');
  sheet_(SHEET_ADMINS).getRange(i+1,3).setValue(new Date());
  const token=Utilities.getUuid()+Utilities.getUuid(); CacheService.getScriptCache().put('admin:'+token,user,21600);
  return {token:token};
}
function requireSession_(token){if(!token||!CacheService.getScriptCache().get('admin:'+token))throw new Error('Admin session expired. Please sign in again.');}

function sheet_(name){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw new Error('Run initializeApp() first.');const s=SpreadsheetApp.openById(id).getSheetByName(name);if(!s)throw new Error('Missing sheet: '+name);return s;}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);sh.getRange(1,1,1,Math.max(headers.length,sh.getLastColumn())).setFontWeight('bold').setBackground('#eaf0ff');return sh;}
function ensureMissingHeaders_(sh,headers){const existing=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0];headers.forEach(h=>{if(!existing.includes(h)){sh.getRange(1,sh.getLastColumn()+1).setValue(h);existing.push(h);}});}
function cleanUniqueNames_(xs){if(!Array.isArray(xs))return [];const out=[];xs.forEach(x=>{const v=String(x||'').trim();if(v&&!out.some(y=>y.toLowerCase()===v.toLowerCase()))out.push(v)});return out;}
function getBatches_(){const sh=sheet_(SHEET_BATCHES);return sh.getDataRange().getValues().slice(1).map(r=>String(r[0]||'').trim()).filter(Boolean);}
function feeMap_(){const rows=sheet_(SHEET_FEES).getDataRange().getValues();return Object.fromEntries(rows.slice(1).filter(r=>r[0]&&r[3]!==false).map(r=>[String(r[0]),Number(r[1])||0]));}
function studentObject_(r){return {timestamp:r[0],formNumber:r[1],rollNumber:r[2],fullName:r[3],fatherName:r[4],motherName:r[5],dob:r[6],gender:r[7],mobile:r[8],email:r[9],address:r[10],className:r[11],school:r[12],photoUrl:r[13],feeAmount:r[14],paymentStatus:r[15],transactionId:r[16],aadhaarNumber:r[17]||'',batch:r[18]||''};}
function classPrefix_(name){const m=String(name).match(/Class\s+(\d+)/i);return m?'C'+m[1]:String(name).toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,5);}
function sha256_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function constantTimeEquals_(a,b){if(a.length!==b.length)return false;let n=0;for(let i=0;i<a.length;i++)n|=a.charCodeAt(i)^b.charCodeAt(i);return n===0;}
function json_(data){return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);}
