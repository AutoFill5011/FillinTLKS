javascript:(async function(){

function sleep(ms){return new Promise(r=>setTimeout(r,ms));}

// ========================
//Working dropdown select
// ========================
async function selectDropdown(index,value,exact=true){

let field=visible[index];

if(!field){
alert("Input "+index+" not found");
return;
}

field.focus();
field.click();

// type value to filter dropdown
field.value=value;
field.dispatchEvent(new Event("input",{bubbles:true}));

await sleep(500);

let options=document.querySelectorAll(".vts-select-item-option");

let target=null;

options.forEach(o=>{

let text=o.innerText.trim();

if(exact){

if(text==value){
target=o;
}

}else{

// numeric safe comparison (fixes 2→02 and prevents 3→30)
if(!isNaN(value) && parseInt(text)==parseInt(value)){
target=o;
}

// normal text partial match (Nhật → Nhật Bản)
else if(isNaN(value) && text.toLowerCase().includes(value.toLowerCase())){
target=o;
}

}

});

if(!target){
alert("Dropdown value '"+value+"' not found at input "+index);
return;
}

target.click();
}

    
// ========================
// LOAD GOOGLE SHEET
// ========================
const sheetID="158pZIljpmJNIL_ab06mZampVXjtdCiti_r4AChjmQvE";
const sheetURL=`https://docs.google.com/spreadsheets/d/${sheetID}/export?format=csv`;

let response=await fetch(sheetURL);
let csv=await response.text();

function parseCSV(text){

const rows=[];
let row=[];
let value='';
let insideQuotes=false;

for(let i=0;i<text.length;i++){

const char=text[i];
const next=text[i+1];

if(char=='"' && insideQuotes && next=='"'){
value+='"';
i++;
}

else if(char=='"'){
insideQuotes=!insideQuotes;
}

else if(char==',' && !insideQuotes){
row.push(value);
value='';
}

else if((char=='\n'||char=='\r') && !insideQuotes){
if(value!==''||row.length){
row.push(value);
rows.push(row);
row=[];
value='';
}
}

else{
value+=char;
}

}

if(value!==''){
row.push(value);
rows.push(row);
}

return rows;
}

let rows=parseCSV(csv).slice(1);

let passport=prompt("Enter passport number");
let timestamp=prompt("Enter timestamp (d/m/yyyy hh:mm:ss)");

if(!passport||!timestamp){
alert("Missing input");
return;
}

// convert timestamp
function convertFormat(input){
const [datePart,timePart]=input.split(" ");
const [day,month,year]=datePart.split("/");
return `${month}/${day}/${year} ${timePart}`;
}

let convertedTimestamp=convertFormat(timestamp);

// find matching row
let match=rows.find(r=>
r[2] && r[2].trim().toLowerCase()==passport.toLowerCase() &&
r[0] && r[0].trim()==convertedTimestamp
);

if(!match){
alert("No matching record found");
return;
}


    
// ========================
// BABY NAME SPLIT
// ========================

let fullNameRaw = match[7];   // column containing full name

// clean name
let fullName = fullNameRaw
    .replace(/"/g,"")
    .replace(/\r/g,"")
    .trim();

let nameParts = fullName.split(/\s+/);

if(nameParts.length < 2){
    alert("Name format error");
    return;
}

// assign parts
let surname = nameParts[0];
let firstName = nameParts[nameParts.length - 1];
let middleName = nameParts.slice(1, -1).join(" ");

// ========================
// FILL FORM INPUTS
// ========================

let elements=document.querySelectorAll("input,select,textarea");
let visible=[];

elements.forEach(el=>{
if(el.offsetParent!==null) visible.push(el);
});

// surname
visible[9].value = surname;
visible[9].dispatchEvent(new Event("input",{bubbles:true}));

// middle name(s)
visible[10].value = middleName;
visible[10].dispatchEvent(new Event("input",{bubbles:true}));

// first name
visible[11].value = firstName;
visible[11].dispatchEvent(new Event("input",{bubbles:true}));
    
// ========================
// GET GENDER FROM SHEET
// ========================
let genderRaw = match[11];

// CLEAN the value
let cleanedGender = genderRaw
    .replace(/"/g,"")
    .replace(/\r/g,"")
    .trim()
    .toLowerCase();

let gender = "Nữ";

if(cleanedGender.includes("nam")){
gender = "Nam";
}

// ========================
// WORKING DROPDOWN METHOD
// ========================

await selectDropdown(12,gender,true);

// ========================
// BABY DATE OF BIRTH
// ========================

let dobRaw = match[8];

// clean
let dob = dobRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

// split
let parts = dob.split("/");

let babyMonth = parseInt(parts[0]).toString();
let babyDay = parseInt(parts[1]).toString();
let babyYear = parts[2];

// ========================
//Baby's Day
await selectDropdown(13,babyDay,false);

// ========================
//Baby's Month
await selectDropdown(14,babyMonth,false);

// ========================
//Baby's Year
visible[15].value = babyYear;
visible[15].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
//Baby's Birth place
// ========================
await selectDropdown(17,"nhật",false);

// ========================
//Baby's Hometown
// ========================
let rawPlace = match[10];

// clean
let cleanedPlace = rawPlace
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

// cut at first separator
let firstWord = cleanedPlace.split(/[ \-_,.]/)[0];

// capitalize
firstWord = firstWord.charAt(0).toUpperCase() + firstWord.slice(1).toLowerCase();

visible[20].value = firstWord;
visible[20].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
//Baby's Ethnicity
// ========================
let ethnicityRaw = match[12];

let ethnicity = ethnicityRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

// special conversion
if(ethnicity === "Tày"){
ethnicity = "Tay";
}

// type to filter
await selectDropdown(25,ethnicity,true);

// ========================
// MOTHER NAME
// ========================

let motherRaw = match[21];

let motherName = motherRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

let motherParts = motherName.split(/\s+/);

let motherSurname="";
let motherMiddle="";
let motherFirst="";

// if only 2 words
if(motherParts.length == 2){

motherSurname = motherParts[0];
motherFirst = motherParts[1];
motherMiddle = "";

}else{

motherSurname = motherParts[0];
motherFirst = motherParts[motherParts.length-1];
motherMiddle = motherParts.slice(1,-1).join(" ");

}

// fill inputs
visible[28].value = motherSurname;
visible[28].dispatchEvent(new Event("input",{bubbles:true}));

if(motherMiddle !== ""){
visible[29].value = motherMiddle;
visible[29].dispatchEvent(new Event("input",{bubbles:true}));
}

visible[30].value = motherFirst;
visible[30].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
// MOTHER BIRTH YEAR
// ========================

let motherYear = match[22]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

visible[33].value = motherYear;
visible[33].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
//Mother's Ethnicity
// ========================
let motherEthnicityRaw = match[23];

let motherEthnicity = motherEthnicityRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

// special conversion
if(motherEthnicity === "Tày"){
motherEthnicity = "Tay";
}

// type to filter
await selectDropdown(34,motherEthnicity,true);

// ========================
//Mother's city
// ========================
await selectDropdown(38,"nhật",false);

// ========================
//Mother's place status
// ========================
await selectDropdown(39,"nơi",false);

// ========================
// MOTHER ADDRESS
// ========================

let motherAddr1 = match[25]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

let motherAddr2 = "";

if(match[26]){
motherAddr2 = match[26]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();
}

let motherAddress = motherAddr1;

if(motherAddr2 !== ""){
motherAddress = motherAddr1 + ", " + motherAddr2;
}

visible[42].value = motherAddress;
visible[42].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
// FATHER NAME
// ========================

let fatherRaw = match[15];

let fatherName = fatherRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

let fatherParts = fatherName.split(/\s+/);

let fatherSurname="";
let fatherMiddle="";
let fatherFirst="";

if(fatherParts.length == 2){

fatherSurname = fatherParts[0];
fatherFirst = fatherParts[1];
fatherMiddle = "";

}else{

fatherSurname = fatherParts[0];
fatherFirst = fatherParts[fatherParts.length-1];
fatherMiddle = fatherParts.slice(1,-1).join(" ");

}

// fill inputs
visible[48].value = fatherSurname;
visible[48].dispatchEvent(new Event("input",{bubbles:true}));

if(fatherMiddle !== ""){
visible[49].value = fatherMiddle;
visible[49].dispatchEvent(new Event("input",{bubbles:true}));
}

visible[50].value = fatherFirst;
visible[50].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
// FATHER BIRTH YEAR
// ========================

let fatherYear = match[16]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

visible[53].value = fatherYear;
visible[53].dispatchEvent(new Event("input",{bubbles:true}));

// ========================
//Father's Ethnicity
// ========================
let fatherEthnicityRaw = match[17];

let fatherEthnicity = fatherEthnicityRaw
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

// special conversion
if(fatherEthnicity === "Tày"){
fatherEthnicity = "Tay";
}

// type to filter
await selectDropdown(54,fatherEthnicity,true);

// ========================
//Father's city
// ========================
await selectDropdown(58,"nhật",false);

// ========================
//Father's place status
// ========================
await selectDropdown(59,"nơi",false);

// ========================
// FATHER ADDRESS
// ========================

let fatherAddr1 = match[19]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

let fatherAddr2 = "";

if(match[20]){
fatherAddr2 = match[20]
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();
}

let fatherAddress = fatherAddr1;

if(fatherAddr2 !== ""){
fatherAddress = fatherAddr1 + ", " + fatherAddr2;
}

visible[62].value = fatherAddress;
visible[62].dispatchEvent(new Event("input",{bubbles:true}));


// ========================
// APPLICANT NAME CHECK
// ========================

// helper to clean text
function cleanText(v){
return v
.replace(/"/g,"")
.replace(/\r/g,"")
.trim()
.toLowerCase();
}
    
function convertDateMDYtoInput(dateStr){

let cleaned = dateStr
.replace(/"/g,"")
.replace(/\r/g,"")
.trim();

let parts = cleaned.split("/");

if(parts.length !== 3){
return cleaned;
}

let month = parseInt(parts[0]);
let day = parseInt(parts[1]);
let year = parts[2];

// force 2-digit format
month = String(month).padStart(2,"0");
day = String(day).padStart(2,"0");

return day + month + year;
}
    
let applicantName = cleanText(match[1]);
let mName = cleanText(match[21]);
let fName = cleanText(match[15]);

let valueCol2 = match[2].replace(/"/g,"").replace(/\r/g,"").trim();
let valueCol3 = match[3].replace(/"/g,"").replace(/\r/g,"").trim();
let valueCol4 = convertDateMDYtoInput(match[4]);

let issuedPlace = valueCol3;
// ========================
// IF MATCH MOTHER
// ========================
if(applicantName === mName){

await selectDropdown(43,"hộ",false);

visible[44].value = valueCol2;
visible[44].dispatchEvent(new Event("input",{bubbles:true}));

visible[45].focus();
visible[45].value = valueCol4;
visible[45].dispatchEvent(new Event("input",{bubbles:true}));
visible[45].dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));

visible[46].value = issuedPlace;
visible[46].dispatchEvent(new Event("input",{bubbles:true}));

}


// ========================
// IF MATCH FATHER
// ========================
else if(applicantName === fName){

await selectDropdown(63,"hộ",false);

visible[64].value = valueCol2;
visible[64].dispatchEvent(new Event("input",{bubbles:true}));

visible[65].focus();
visible[65].value = valueCol4;
visible[65].dispatchEvent(new Event("input",{bubbles:true}));
visible[65].dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));

visible[66].value = issuedPlace;
visible[66].dispatchEvent(new Event("input",{bubbles:true}));

}


// ========================
// NO MATCH
// ========================
else{

alert("Kiểm tra lại tên không trùng");

await selectDropdown(63,"hộ",false);

visible[64].value = valueCol2;
visible[64].dispatchEvent(new Event("input",{bubbles:true}));

visible[65].focus();
visible[65].value = valueCol4;
visible[65].dispatchEvent(new Event("input",{bubbles:true}));
visible[65].dispatchEvent(new KeyboardEvent("keydown",{key:"Enter",bubbles:true}));

visible[66].value = issuedPlace;
visible[66].dispatchEvent(new Event("input",{bubbles:true}));

}

alert("Điền form thành công!!!");
})();
