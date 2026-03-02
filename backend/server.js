const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const PptxGenJS = require('pptxgenjs');
const PDFDocument = require('pdfkit');

const app = express();
const PORT = process.env.PORT || 3000;

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Middleware
app.use(cors());
app.use(express.json());

// Track server start time for uptime calculation
const startTime = Date.now();

// ========================================
// DATE HELPER FUNCTIONS
// ========================================

/**
 * Normalizes a Date object to local midnight (date-only, no time component).
 * @param {Date} dateObj - A JavaScript Date object
 * @returns {Date} A new Date at local midnight (year, month, day only)
 */
function normalizeToDateOnly(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) {
    return null;
  }
  const y = dateObj.getFullYear();
  const m = dateObj.getMonth();
  const d = dateObj.getDate();
  return new Date(y, m, d);
}

/**
 * Parses the "مُنشأ في" cell value into a JS Date (date-only).
 * Supports:
 * - "DD-MM-YYYY" string format
 * - "DD/MM/YYYY" string format
 * - Excel serial number (e.g. 45123)
 * - Already a Date object
 * - Returns null for blank/invalid values (no crash)
 * 
 * @param {any} value - The cell value from "مُنشأ في" column
 * @returns {Date|null} A Date at local midnight, or null if invalid
 */
function parseCreatedAtToDateOnly(value) {
  // Handle null, undefined, empty string, or "Blank"
  if (value === null || value === undefined) {
    return null;
  }
  
  // If already a Date object, normalize it
  if (value instanceof Date) {
    return normalizeToDateOnly(value);
  }
  
  // Handle Excel serial number (number type)
  if (typeof value === 'number') {
    try {
      // Use XLSX.SSF.parse_date_code if available
      if (XLSX.SSF && typeof XLSX.SSF.parse_date_code === 'function') {
        const parsed = XLSX.SSF.parse_date_code(value);
        if (parsed && parsed.y && parsed.m && parsed.d) {
          // parsed.m is 1-indexed in XLSX, JS Date month is 0-indexed
          return new Date(parsed.y, parsed.m - 1, parsed.d);
        }
      }
      
      // Fallback: manual Excel serial to date conversion
      // Excel serial date: days since 1899-12-30 (accounting for Excel's leap year bug)
      // Excel incorrectly considers 1900 a leap year, so dates after Feb 28, 1900 are off by 1
      const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
      const jsDate = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
      return normalizeToDateOnly(jsDate);
    } catch (err) {
      return null;
    }
  }
  
  // Handle string values
  if (typeof value === 'string') {
    const strValue = value.trim();
    
    // Check for blank/empty
    if (strValue === '' || strValue.toLowerCase() === 'blank') {
      return null;
    }
    
    // Try parsing DD-MM-YYYY or DD/MM/YYYY format
    if (strValue.includes('/') || strValue.includes('-')) {
      const separator = strValue.includes('/') ? '/' : '-';
      const parts = strValue.split(separator);
      
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10);
        const year = parseInt(parts[2], 10);
        
        // Validate parsed values
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) &&
            day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900) {
          // JS months are 0-indexed
          return new Date(year, month - 1, day);
        }
      }
    }
    
    // Try parsing as a number string (Excel serial as string)
    const numValue = parseFloat(strValue);
    if (!isNaN(numValue) && numValue > 0) {
      return parseCreatedAtToDateOnly(numValue);
    }
  }
  
  // Unrecognized format
  return null;
}

// ========================================
// DATE HELPER TESTS (optional console tests)
// ========================================
console.log('\n🧪 === Date Helper Function Tests ===');

// Test normalizeToDateOnly
const testDate1 = new Date(2024, 5, 15, 14, 30, 45); // June 15, 2024 14:30:45
const normalized1 = normalizeToDateOnly(testDate1);
console.log(`normalizeToDateOnly(${testDate1.toISOString()}) => ${normalized1 ? normalized1.toISOString() : null}`);

// Test parseCreatedAtToDateOnly with various formats
const testCases = [
  { input: '15-06-2024', desc: 'DD-MM-YYYY' },
  { input: '15/06/2024', desc: 'DD/MM/YYYY' },
  { input: 45458, desc: 'Excel serial (45458 = ~Jun 15, 2024)' },
  { input: new Date(2024, 5, 15), desc: 'Date object' },
  { input: '', desc: 'Empty string' },
  { input: 'Blank', desc: 'Blank string' },
  { input: null, desc: 'null' },
  { input: 'invalid', desc: 'Invalid string' },
];

testCases.forEach(({ input, desc }) => {
  const result = parseCreatedAtToDateOnly(input);
  const resultStr = result ? result.toDateString() : 'null';
  console.log(`parseCreatedAtToDateOnly(${JSON.stringify(input)}) [${desc}] => ${resultStr}`);
});

console.log('='.repeat(50) + '\n');

// ========================================
// EXCEL COLUMN FORMATTING HELPER
// Applies fixed display formats to specific columns by header name
// ========================================

/**
 * Apply Excel number formats to specific columns by matching header names.
 * This function runs on every Saher Excel export to enforce consistent formatting.
 * 
 * @param {Object} worksheet - The XLSX worksheet object
 * @param {Array} headers - Array of column header names
 * @param {Object} formatConfig - Configuration object with column names and their formats
 *   Example: { 'Column Name': 'DD/MM/YYYY' }
 */
function applyColumnFormatsByHeader(worksheet, headers, formatConfig) {
  if (!worksheet || !headers || !formatConfig) return;
  
  // Get the worksheet range
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
  const totalRows = range.e.r + 1; // Total number of rows (0-indexed, so +1)
  
  // Process each column that needs formatting
  Object.entries(formatConfig).forEach(([headerName, format]) => {
    // Find the column index for this header
    const colIndex = headers.indexOf(headerName);
    
    if (colIndex === -1) {
      // Header not found, skip safely
      return;
    }
    
    // Get the column letter (A, B, C, ..., AA, AB, etc.)
    const colLetter = XLSX.utils.encode_col(colIndex);
    
    // Apply format to all data cells in this column (skip header row 0)
    for (let row = 1; row < totalRows; row++) {
      const cellAddress = `${colLetter}${row + 1}`; // Excel is 1-indexed
      const cell = worksheet[cellAddress];
      
      if (cell) {
        // Set the number format (z property in XLSX.js)
        cell.z = format;
      }
    }
    
    console.log(`📋 Applied format "${format}" to column "${headerName}" (${colLetter})`);
  });
}

/**
 * Standard column formats for Saher Excel exports.
 * These are applied automatically on every export.
 */
const SAHER_COLUMN_FORMATS = {
  // Date columns - format: DD/MM/YYYY
  'SADAD Enquiry Date': 'DD/MM/YYYY',
  'تاريخ المخالفة': 'DD/MM/YYYY',
  'مُنشأ في': 'DD/MM/YYYY',
  'تم التغيير في': 'DD/MM/YYYY',
  
  // Time columns - format: hh:mm:ss AM/PM
  'SADAD Enquiry Time': 'hh:mm:ss AM/PM',
  'Violation Time': 'hh:mm:ss AM/PM',
  'وقت الإنشاء': 'hh:mm:ss AM/PM',
  'وقت التغيير': 'hh:mm:ss AM/PM'
};

// Health check endpoint
app.get('/health', (req, res) => {
  const uptime = Math.floor((Date.now() - startTime) / 1000);
  
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    uptime: uptime,
    message: 'Server is running smoothly'
  });
});

// ========================================
// EXTRACT DATES ENDPOINT
// Extracts available date range from uploaded file
// ========================================
app.post('/api/extract-dates', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // fileType: 'saher' or 'mva' (accidents)
    const { fileType } = req.body;
    
    if (!fileType || !['saher', 'mva'].includes(fileType)) {
      return res.status(400).json({ error: 'fileType must be "saher" or "mva"' });
    }

    console.log(`\n📅 === Extracting dates from ${fileType.toUpperCase()} file ===`);

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    
    // Target "Sheet1" specifically, or fallback to first sheet
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    console.log(`📄 Processing sheet: "${sheetName}"`);
    
    // Convert to JSON for processing
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    if (data.length === 0) {
      return res.status(200).json({
        success: false,
        error: 'No data found in file',
        minDate: null,
        maxDate: null,
        availableDates: []
      });
    }

    // Determine date column based on file type
    // Saher file: "مُنشأ في" column
    // MVA/Accidents file: To be determined later (placeholder for now)
    let dateColumn;
    if (fileType === 'saher') {
      dateColumn = 'مُنشأ في';
    } else {
      // MVA/Accidents file - placeholder column name
      // This will be updated when accident data logic is shared
      dateColumn = 'Date'; // Fallback - will be updated later
    }

    console.log(`📅 Using date column: "${dateColumn}"`);

    // Check if the date column exists
    const columns = Object.keys(data[0]);
    if (!columns.includes(dateColumn)) {
      console.log(`⚠️ Date column "${dateColumn}" not found. Available columns:`, columns);
      return res.status(200).json({
        success: false,
        error: `Date column "${dateColumn}" not found in file`,
        minDate: null,
        maxDate: null,
        availableDates: []
      });
    }

    // Extract all unique dates from the file
    const today = normalizeToDateOnly(new Date());
    const uniqueDatesSet = new Set();
    let minDate = null;
    let maxDate = null;

    data.forEach(row => {
      const dateValue = row[dateColumn];
      const parsedDate = parseCreatedAtToDateOnly(dateValue);
      
      if (parsedDate !== null) {
        // Only include dates up to today (no future dates)
        if (parsedDate <= today) {
          const dateString = parsedDate.toISOString().split('T')[0]; // YYYY-MM-DD format
          uniqueDatesSet.add(dateString);
          
          // Track min and max dates
          if (minDate === null || parsedDate < minDate) {
            minDate = parsedDate;
          }
          if (maxDate === null || parsedDate > maxDate) {
            maxDate = parsedDate;
          }
        }
      }
    });

    // Convert Set to sorted array
    const availableDates = Array.from(uniqueDatesSet).sort();

    // Ensure maxDate doesn't exceed today
    if (maxDate !== null && maxDate > today) {
      maxDate = today;
    }

    console.log(`📅 Found ${availableDates.length} unique dates`);
    console.log(`📅 Date range: ${minDate ? minDate.toISOString().split('T')[0] : 'null'} to ${maxDate ? maxDate.toISOString().split('T')[0] : 'null'}`);

    return res.status(200).json({
      success: true,
      minDate: minDate ? minDate.toISOString().split('T')[0] : null,
      maxDate: maxDate ? maxDate.toISOString().split('T')[0] : null,
      availableDates: availableDates,
      totalDates: availableDates.length
    });

  } catch (error) {
    console.error('Error extracting dates:', error);
    return res.status(500).json({ 
      error: 'Error extracting dates from file',
      details: error.message 
    });
  }
});

// Process Saher file endpoint
app.post('/api/process-saher', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Validate date range inputs
    const { startDate, endDate } = req.body;
    
    if (!startDate || !endDate) {
      return res.status(400).json({ error: 'startDate and endDate are required' });
    }
    
    // Parse dates from ISO format "YYYY-MM-DD" to JS Date objects (date-only)
    const parsedStartDate = new Date(startDate + 'T00:00:00');
    const parsedEndDate = new Date(endDate + 'T00:00:00');
    
    // Validate that the parsed dates are valid
    if (isNaN(parsedStartDate.getTime())) {
      return res.status(400).json({ error: 'Invalid startDate format. Expected YYYY-MM-DD' });
    }
    
    if (isNaN(parsedEndDate.getTime())) {
      return res.status(400).json({ error: 'Invalid endDate format. Expected YYYY-MM-DD' });
    }
    
    console.log(`📅 Date range: ${startDate} to ${endDate}`);
    console.log(`📅 Parsed dates: ${parsedStartDate.toISOString()} to ${parsedEndDate.toISOString()}`);

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    
    // Target "Sheet1" specifically for Saher file
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    console.log(`\n📄 Processing sheet: "${sheetName}"`);
    
    // Convert to JSON for processing
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    // Log all column headers
    console.log('\n📋 === Column Headers ===');
    if (data.length > 0) {
      console.log(Object.keys(data[0]));
    }

    // ========================================
    // DATE FILTERING - Filter by "مُنشأ في" column
    // ========================================
    console.log(`\n📅 Filtering data by date range: ${startDate} to ${endDate}`);
    console.log(`📅 Total rows before filtering: ${data.length}`);
    
    // Normalize start and end dates to date-only for comparison
    const filterStartDate = normalizeToDateOnly(parsedStartDate);
    const filterEndDate = normalizeToDateOnly(parsedEndDate);
    
    // Filter rows by "مُنشأ في" column
    const filteredData = data.filter(row => {
      const createdAtValue = row['مُنشأ في'];
      const createdAtDate = parseCreatedAtToDateOnly(createdAtValue);
      
      // Exclude rows with null/invalid dates
      if (createdAtDate === null) {
        return false;
      }
      
      // Include rows where createdAtDate >= startDate AND createdAtDate <= endDate
      return createdAtDate >= filterStartDate && createdAtDate <= filterEndDate;
    });
    
    console.log(`📅 Total rows after filtering: ${filteredData.length}`);
    
    // If no data in selected range, return early with noData response
    if (filteredData.length === 0) {
      console.log('⚠️ No data found for selected date range');
      return res.status(200).json({
        noData: true,
        message: 'No data for selected range'
      });
    }
    
    // Use filtered data for all subsequent processing
    const dataToProcess = filteredData;

    // Log all unique Business Line Org Description values
    const uniqueValues = [...new Set(dataToProcess.map(row => row['Business Line Org Description'] || 'Blank'))];
    console.log('\n📋 === Business Line Org Description - Unique Values ===');
    console.log('Total rows:', dataToProcess.length);
    console.log('Unique values found:', uniqueValues.length);
    uniqueValues.forEach((val, index) => {
      console.log(`  ${index + 1}. "${val}"`);
    });
    console.log('='.repeat(50));

    // Normalize Arabic text (handle أ/ا variations)
    const normalizeArabic = (text) => {
      if (!text) return '';
      return text
        .replace(/أ/g, 'ا')
        .replace(/إ/g, 'ا')
        .replace(/آ/g, 'ا')
        .replace(/ى/g, 'ي')
        .replace(/ة/g, 'ه')
        .trim();
    };

    // Process each row
    const processedData = dataToProcess.map((row, index) => {
      // Get Department Org Code for Sector Org Description mapping
      const deptOrgCode = String(row['Department Org Code'] || '').trim();
      
      // Rule 0: Map Department Org Code to Sector Org Description
      const sectorMapping = {
        // وحدة أعمال التوزيع وخدمات المشتركين-وسطى
        '4107001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4108001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4109001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4110001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4122001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4123001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4124001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4125001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4106001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
     /*   '4126001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4111001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',*/
        
        // وحدة أعمال التوزيع وخدمات المشتركين-غربي
        '4201001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4202001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4204001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4205001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4210001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4220001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4230001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        /*'4221001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',*/
        
        
        // وحدة أعمال التوزيع وخدمات المشتركين-شرقي
        '4301001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4302001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4303001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4306001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4307001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4308001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4309001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
    /*    '4316001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',*/
        '4310001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        
        // وحدة أعمال التوزيع وخدمات المشتركين-جنوب
        '4405001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4406001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4411001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4412001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4413001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4414001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        
        // قطاع عمليات انتاج الطاقة الغربي
        '2231001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2232001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2233001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2241001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2242001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2243001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2251001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2252001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2253001': 'قطاع عمليات انتاج الطاقة الغربي',
        
        // قطاع عمليات انتاج الطاقة الشرقي
        '2311001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2312001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2313001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2321001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2322001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2323001': 'قطاع عمليات انتاج الطاقة الشرقي',
        
        // قطاع عمليات انتاج الطاقة الجنوبي
        '2041001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2043001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2046001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2411001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2412001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2413001': 'قطاع عمليات انتاج الطاقة الجنوبي'
      };
      
      // Update Sector Org Description if Department Org Code matches
      if (sectorMapping[deptOrgCode]) {
        row['Sector Org Description'] = sectorMapping[deptOrgCode];
      }
      
      // ========================================
      // FALLBACK RULES for Sector Org Description
      // Apply only if Sector Org Description is still empty/blank/"Blank"
      // ========================================
      const rawSectorOrgDesc = row['Sector Org Description'];
      const currentSectorOrgDesc = String(rawSectorOrgDesc || '').trim();
      const divisionOrgCode = String(row['Division Org Code'] || '').trim();
      const highestDeptOrg = String(row['Highest Department Org'] || '').trim();
      
      // Debug log for first 10 rows to verify input columns
      if (index < 10) {
        console.log(`[Sector Fallback Debug] row=${index} | Sector Org Description (raw)="${rawSectorOrgDesc}" | Division Org Code="${divisionOrgCode}" | Highest Department Org="${highestDeptOrg}"`);
      }
      
      // Check if blank: empty string, whitespace-only, or "Blank" (case-insensitive)
      const isBlankSector = !currentSectorOrgDesc || currentSectorOrgDesc.toLowerCase() === 'blank';
      
      if (isBlankSector) {
        // Rule 1: Southern Distribution - based on Division Org Code
        if (divisionOrgCode === '4400401') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب';
          console.log(`[Sector Fallback Applied] row=${index} rule=SouthernDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 2: Eastern Distribution - based on Highest Department Org
        else if (highestDeptOrg === '4310001') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي';
          console.log(`[Sector Fallback Applied] row=${index} rule=EasternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 3: Western Distribution - based on Highest Department Org
        else if (['4210001', '4220001', '4230001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-غربي';
          console.log(`[Sector Fallback Applied] row=${index} rule=WesternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 4: Central Distribution - based on Highest Department Org
        else if (['4110001', '4126001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى';
          console.log(`[Sector Fallback Applied] row=${index} rule=CentralDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 5: Central Distribution - based on Highest Department Org
        else if (['4502001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال الشبكات الذكية';
          console.log(`[Sector Fallback Applied] row=${index} rule=CentralDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 6: Western Generation Operations - based on Highest Department Org
        else if (highestDeptOrg === '2027001') {
          row['Sector Org Description'] = 'قطاع عمليات انتاج الطاقة الغربي';
          console.log(`[Sector Fallback Applied] row=${index} rule=WesternGenerationOps division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
      }
      
      const businessLineDesc = row['Business Line Org Description'] || 'Blank';
      const normalizedDesc = normalizeArabic(businessLineDesc);
      const sectorOrgDesc = row['Sector Org Description'] || 'Blank';
      const highestDeptDesc = row['Highest Department Org Description'] || 'Blank';
      const deptDesc = row['Department Org Description'] || 'Blank';
      const divisionDesc = row['Division Org Description'] || 'Blank';
      const costCenterDesc = row['Costt Center Org Description'] || row['Cost Center Org Description'] || 'Blank';

      let newBusinessLineDesc = businessLineDesc;
   
      // Rule 1: CEO Org Code mapping (highest priority)
const ceoOrgCode = String(row['CEO Org Code'] || '').trim();

if (ceoOrgCode === '30000001') {
  newBusinessLineDesc = 'الشركة الوطنية لنقل الكهرباء';
} else if (ceoOrgCode === '91000001') {
  newBusinessLineDesc = 'Others';
}


      // Rule 2: Activity (النشاط) column mapping - only if CEO Org Code is NOT 91000001
      if (ceoOrgCode !== '30000001' && ceoOrgCode !== '91000001') {
        const activityCode = String(row['النشاط'] || '').trim();

        // نشاط التوزيع وخدمات المشتركين
        if (['4000001', '4400001', '4200001', '4100001', '4500001', '4300001', '4600001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوزيع وخدمات المشتركين';
        }
        // نشاط التوليد
        else if (['2000001', '2100001', '2200001', '2400001', '2600001', '2300001', '2500001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوليد';
        }
        // نشاط الخدمات الفنية
        else if (activityCode === '16000001') {
          newBusinessLineDesc = 'نشاط الخدمات الفنية';
        }
        // نشاط الصحة المهنية والسلامة والامن والبيئة
        else if (activityCode === '1100001') {
          newBusinessLineDesc = 'نشاط الصحة المهنية والسلامة والامن والبيئة';
        }
        // نشاط الموارد البشرية والخدمات المساندة
        else if (['7000001', '7100001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط الموارد البشرية والخدمات المساندة';
        }
        // NOTE: Activity codes ['11000001', '5000001', '5100001', '10000001'] are classified as "Others"
        // ONLY in statistics aggregation, NOT in the cleaned data. Original value is preserved here.
      }

      // Rule 3: If still Blank, check columns and set to N/A
      // Chain: Business Line → Sector Org → Highest Dept → Dept → Division (Blank only)
      const isBlank = (val) => !val || val === 'Blank' || val.trim() === '';

      if (isBlank(businessLineDesc) && newBusinessLineDesc === businessLineDesc) {
        if (
          isBlank(sectorOrgDesc) &&
          isBlank(highestDeptDesc) &&
          isBlank(deptDesc) &&
          isBlank(divisionDesc)
        ) {
          newBusinessLineDesc = 'N/A';
        }
      }

      // Rule 7: If CEO Org Description contains "شركة وادي الحلول" → N/A
      const ceoOrgDesc = row['CEO Org Description'] || 'Blank';
      if (ceoOrgDesc.includes('شركة وادي الحلول')) {
        newBusinessLineDesc = 'N/A';
      }

      // Log transformation if value changed
      if (businessLineDesc !== newBusinessLineDesc) {
        console.log(`✓ Changed: "${businessLineDesc}" → "${newBusinessLineDesc}"`);
      }

      // Update the row with new value (same column we read from)
      row['Business Line Org Description'] = newBusinessLineDesc;
      
      return row;
    });

    // Get all existing column headers from the original data
    const existingHeaders = processedData.length > 0 ? Object.keys(processedData[0]) : [];
    
    // Add new column "اسم المنطقة" at the very end
    const newColumnName = 'اسم المنطقة';
    const allHeaders = [...existingHeaders, newColumnName];
    
    console.log(`\n📋 Adding new column "${newColumnName}" after ${existingHeaders.length} existing columns`);
    console.log(`📋 Total columns in output: ${allHeaders.length}`);
    
    // Add the new column value to each row based on "المنطقة" column
    processedData.forEach(row => {
      const regionValue = String(row['المنطقة'] || '').trim();
      
      // Map region codes to names
      if (!regionValue || regionValue === 'Blank' || regionValue === '') {
        row[newColumnName] = 'N/A';
      } else if (regionValue.includes('6410')) {
        row[newColumnName] = 'COA';
      } else if (regionValue.includes('6420')) {
        row[newColumnName] = 'WOA';
      } else if (regionValue.includes('6430')) {
        row[newColumnName] = 'EOA';
      } else if (regionValue.includes('6440')) {
        row[newColumnName] = 'SOA';
      } else {
        row[newColumnName] = 'N/A';  // No match, set to N/A
      }
    });

    // Columns to remove from output (delete entirely)
    const columnsToRemove = [
      'وصف المعدات',
      'رقم الصنف',
      'وصف الصنف',
      'CEO Org Description',
      'Department Org Code',
      'Cost Center Org Description',
      'Department Org Description'
    ];

    // Remove these columns from each row
    processedData.forEach(row => {
      columnsToRemove.forEach(col => {
        if (row.hasOwnProperty(col)) {
          delete row[col];
        }
      });
    });

    // Filter out removed columns from headers
    const filteredHeaders = allHeaders.filter(h => !columnsToRemove.includes(h));
    
    console.log(`🗑️ Removed ${columnsToRemove.length} columns from output`);

    // Reorder columns: Move these 5 columns to the END of the output
    const columnsToMoveToEnd = [
      'اسم المنطقة',
      'Business Line Org Description',
      'Sector Org Description',
      'Highest Department Org Description',
      'Division Org Description'
    ];

    // Remove the 5 columns from their current positions
    let reorderedHeaders = filteredHeaders.filter(h => !columnsToMoveToEnd.includes(h));
    
    // Get only the columns that actually exist in filteredHeaders
    const columnsToAppend = columnsToMoveToEnd.filter(col => filteredHeaders.includes(col));
    
    // Append the 5 columns at the end in the specified order
    reorderedHeaders = [...reorderedHeaders, ...columnsToAppend];

    console.log(`📋 Final column count: ${reorderedHeaders.length}`);
    console.log(`📋 Last 5 columns: ${columnsToAppend.join(', ')}`);

    // ========================================
    // STATISTICS CALCULATION (Pure JavaScript)
    // ========================================
    const calculateSaherStatistics = (data) => {
      // Define the fixed regions for the table
      const regions = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'];
      
      // Count cancelled rows
      const cancelledRows = data.filter(row => {
        const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
        return paymentStatus === 'CANCEL';
      });
      const cancelCount = cancelledRows.length;
      
      // Filter out cancelled rows for statistics
      const activeRows = data.filter(row => {
        const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
        return paymentStatus !== 'CANCEL';
      });
      
      // Total violations (excluding cancels)
      const totalViolations = activeRows.length;
      
      // ----------------------------------------
      // Table 1: Business Line × Region Matrix
      // FIXED ROW AND COLUMN ORDER
      // ----------------------------------------
      
      // Fixed row order (Business Lines) - ALWAYS in this exact order
      const fixedBusinessLines = [
        'Generation',
        'National grid',
        'Distribution',
        'PDC',
        'Technical Services',
        'HR',
        'DT',
        'HSSE',
        'Others'
      ];
      
      // Activity codes that should be classified as "Others" in statistics only
      const othersActivityCodes = ['11000001', '5000001', '5100001', '10000001'];
      
      // Mapping from Arabic Business Line names to fixed English labels
      const businessLineMapping = {
        'نشاط التوليد': 'Generation',
        'الشركة الوطنية لنقل الكهرباء': 'National grid',
        'نشاط التوزيع وخدمات المشتركين': 'Distribution',
        'نشاط الخدمات الفنية': 'Technical Services',
        'نشاط الموارد البشرية والخدمات المساندة': 'HR',
        'نشاط الصحة المهنية والسلامة والامن والبيئة': 'HSSE',
        'Others': 'Others',
        'N/A': 'Others'  // Map N/A to Others
      };
      
      // Helper function to get the statistical classification for Business Line
      // Maps to fixed English labels for statistics only
      const getStatisticalBusinessLine = (row) => {
        const activityCode = String(row['النشاط'] || '').trim();
        
        // If activity code is in the "Others" list, classify as "Others" for stats
        if (activityCode && othersActivityCodes.includes(activityCode)) {
          return 'Others';
        }
        
        // Get the actual Business Line Org Description
        const bl = String(row['Business Line Org Description'] || 'N/A').trim();
        
        // Map to fixed English label, default to "Others" if not found
        return businessLineMapping[bl] || 'Others';
      };
      
      // Initialize the matrix with ALL fixed rows and columns (defaulting to 0)
      const businessLineMatrix = {};
      fixedBusinessLines.forEach(bl => {
        businessLineMatrix[bl] = {};
        regions.forEach(region => {
          businessLineMatrix[bl][region] = 0;
        });
        businessLineMatrix[bl]['Total'] = 0;
      });
      
      // Initialize totals row
      const regionTotals = {};
      regions.forEach(region => {
        regionTotals[region] = 0;
      });
      regionTotals['Total'] = 0;
      
      // Count violations per Business Line × Region
      // Uses statistical classification (maps to fixed English labels)
      activeRows.forEach(row => {
        const bl = getStatisticalBusinessLine(row);
        const region = String(row['اسم المنطقة'] || 'N/A').trim();
        
        // Normalize region to match our fixed list (blank → N/A)
        const normalizedRegion = regions.includes(region) ? region : 'N/A';
        
        // Always increment - bl is guaranteed to be in fixedBusinessLines
        businessLineMatrix[bl][normalizedRegion]++;
        businessLineMatrix[bl]['Total']++;
        regionTotals[normalizedRegion]++;
        regionTotals['Total']++;
      });
      
      // Build the table array in FIXED ORDER (always all rows, even if 0)
      const businessLineTable = fixedBusinessLines.map(bl => {
        const row = { 'Business Line Org Description': bl };
        regions.forEach(region => {
          row[region] = businessLineMatrix[bl][region];
        });
        row['Total'] = businessLineMatrix[bl]['Total'];
        return row;
      });
      
      // Add totals row at the end
      const totalsRow = { 'Business Line Org Description': 'Total' };
      regions.forEach(region => {
        totalsRow[region] = regionTotals[region];
      });
      totalsRow['Total'] = regionTotals['Total'];
      businessLineTable.push(totalsRow);
      
      // ----------------------------------------
      // Table 2: HSSE Violations Table (GroupLabel × Region)
      // ----------------------------------------
      
      // Filter for HSSE violations:
      // Business Line = "نشاط الصحة المهنية والسلامة والامن والبيئة"
      const hsseRows = activeRows.filter(row => {
        const bl = String(row['Business Line Org Description'] || '').trim();
        return bl === 'نشاط الصحة المهنية والسلامة والامن والبيئة';
      });
      
      // Helper to check if value is blank
      const isBlankValue = (val) => !val || val === 'Blank' || String(val).trim() === '';
      
      // Get unique GroupLabels using fallback logic
      const hsseGroupLabels = [...new Set(hsseRows.map(row => {
        const sector = String(row['Sector Org Description'] || '').trim();
        const highestDept = String(row['Highest Department Org Description'] || '').trim();
        const division = String(row['Division Org Description'] || '').trim();
        
        if (!isBlankValue(sector)) return sector;
        if (!isBlankValue(highestDept)) return highestDept;
        if (!isBlankValue(division)) return division;
        return 'N/A';
      }))].sort();
      
      // Initialize the HSSE matrix (GroupLabel × Region)
      const hsseMatrix = {};
      hsseGroupLabels.forEach(label => {
        hsseMatrix[label] = {};
        regions.forEach(region => {
          hsseMatrix[label][region] = 0;
        });
        hsseMatrix[label]['Total'] = 0;
      });
      
      // Initialize HSSE region totals
      const hsseRegionTotals = {};
      regions.forEach(region => {
        hsseRegionTotals[region] = 0;
      });
      hsseRegionTotals['Total'] = 0;
      
      // Count HSSE violations per GroupLabel × Region
      hsseRows.forEach(row => {
        const sector = String(row['Sector Org Description'] || '').trim();
        const highestDept = String(row['Highest Department Org Description'] || '').trim();
        const division = String(row['Division Org Description'] || '').trim();
        
        // Apply fallback logic for GroupLabel
        let groupLabel;
        if (!isBlankValue(sector)) {
          groupLabel = sector;
        } else if (!isBlankValue(highestDept)) {
          groupLabel = highestDept;
        } else if (!isBlankValue(division)) {
          groupLabel = division;
        } else {
          groupLabel = 'N/A';
        }
        
        // Get region from "اسم المنطقة", treat blank as N/A
        const regionValue = String(row['اسم المنطقة'] || '').trim();
        const normalizedRegion = (isBlankValue(regionValue) || !regions.includes(regionValue)) ? 'N/A' : regionValue;
        
        if (hsseMatrix[groupLabel]) {
          hsseMatrix[groupLabel][normalizedRegion]++;
          hsseMatrix[groupLabel]['Total']++;
          hsseRegionTotals[normalizedRegion]++;
          hsseRegionTotals['Total']++;
        }
      });
      
      // Build HSSE table array
      const hsseTable = hsseGroupLabels.map(label => {
        const row = { 'Group': label };
        regions.forEach(region => {
          row[region] = hsseMatrix[label][region];
        });
        row['Total'] = hsseMatrix[label]['Total'];
        return row;
      });
      
      // Add totals row
      const hsseTotalsRow = { 'Group': 'Total' };
      regions.forEach(region => {
        hsseTotalsRow[region] = hsseRegionTotals[region];
      });
      hsseTotalsRow['Total'] = hsseRegionTotals['Total'];
      hsseTable.push(hsseTotalsRow);
      
      // Return structured JSON
      return {
        summary: {
          cancelCount: cancelCount,
          totalViolations: totalViolations,
          totalRowsProcessed: data.length
        },
        businessLineByRegion: {
          columns: ['Business Line Org Description', ...regions, 'Total'],
          data: businessLineTable
        },
        hsseViolations: {
          columns: ['Group', ...regions, 'Total'],
          data: hsseTable
        }
      };
    };
    
    // Calculate statistics from processed data
    const statistics = calculateSaherStatistics(processedData);
    
    // Log statistics to console
    console.log('\n📊 === SAHER STATISTICS ===');
    console.log(`Cancel Count: ${statistics.summary.cancelCount}`);
    console.log(`Total Violations (excluding cancels): ${statistics.summary.totalViolations}`);
    console.log(`Total Rows Processed: ${statistics.summary.totalRowsProcessed}`);
    console.log('\n📊 Business Line × Region Table:');
    console.table(statistics.businessLineByRegion.data);
    console.log('\n📊 HSSE Violations (Group × Region):');
    console.table(statistics.hsseViolations.data);
    console.log('='.repeat(50));
    
    // ========================================
    // WEEKLY TOTALS CALCULATION
    // ========================================
    
    // Helper: Format date as YYYY-MM-DD
    const formatDateYMD = (date) => {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    };
    
    // Helper: Get the most recent Wednesday on or before a given date
    const getLastWednesday = (date) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const dayOfWeek = d.getDay(); // 0=Sun, 1=Mon, ..., 3=Wed, ..., 6=Sat
      const daysToSubtract = (dayOfWeek + 7 - 3) % 7; // Days since last Wednesday
      d.setDate(d.getDate() - daysToSubtract);
      return d;
    };
    
    // Helper: Add days to a date
    const addDays = (date, days) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      d.setDate(d.getDate() + days);
      return d;
    };
    
    // Compute weekly range based on selectedEndDate
    const selectedEndDateOnly = normalizeToDateOnly(parsedEndDate);
    const lastWednesday = getLastWednesday(selectedEndDateOnly);
    
    let weeklyStart, weeklyEnd;
    
    // Check if selectedEndDate is a Wednesday (day 3)
    if (selectedEndDateOnly.getDay() === 3) {
      // selectedEndDate IS Wednesday: use previous full week (Wed to Tue)
      weeklyStart = addDays(lastWednesday, -7);
      weeklyEnd = addDays(weeklyStart, 6); // Tuesday
    } else {
      // selectedEndDate is NOT Wednesday: weekly-to-date from last Wednesday
      weeklyStart = lastWednesday;
      weeklyEnd = selectedEndDateOnly;
    }
    
    console.log(`\n📅 Weekly range: ${formatDateYMD(weeklyStart)} to ${formatDateYMD(weeklyEnd)}`);
    
    // Count weekly violations (NON-canceled rows within weeklyStart..weeklyEnd)
    let weeklyTotalViolations = 0;
    let weeklyHsseViolations = 0;
    
    processedData.forEach(row => {
      // Check cancellation status - MUST exclude CANCEL rows first
      const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (paymentStatus === 'CANCEL') {
        return; // Skip cancelled rows
      }
      
      // Parse the date from the row
      const createdAtValue = row['مُنشأ في'];
      const createdAtDate = parseCreatedAtToDateOnly(createdAtValue);
      
      if (createdAtDate === null) {
        return; // Skip invalid dates
      }
      
      // Check if within weekly range (inclusive)
      if (createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd) {
        weeklyTotalViolations++;
        
        // Check if this is an HSSE violation
        const businessLine = String(row['Business Line Org Description'] || '').trim();
        if (businessLine === 'نشاط الصحة المهنية والسلامة والامن والبيئة') {
          weeklyHsseViolations++;
        }
      }
    });
    
    console.log(`📅 Weekly total violations: ${weeklyTotalViolations}`);
    console.log(`📅 Weekly HSSE violations: ${weeklyHsseViolations}`);

    // ========================================
    // RETURN JSON RESPONSE
    // ========================================
    res.json({
      noData: false,
      selectedTotals: {
        totalViolations: statistics.summary.totalViolations,
        cancels: statistics.summary.cancelCount
      },
      weeklyTotals: {
        weeklyStartDate: formatDateYMD(weeklyStart),
        weeklyEndDate: formatDateYMD(weeklyEnd),
        totalViolations: weeklyTotalViolations,
        hsseViolations: weeklyHsseViolations
      },
      statisticsTables: {
        businessLineByRegion: statistics.businessLineByRegion,
        hsseViolations: statistics.hsseViolations
      }
    });

  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).json({ error: 'Error processing file', details: error.message });
  }
});

// Get Saher statistics as JSON endpoint
app.post('/api/saher-stats', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    
    // Target "Sheet1" specifically for Saher file
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON for processing
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    // Normalize Arabic text (handle أ/ا variations)
    const normalizeArabic = (text) => {
      if (!text) return '';
      return text
        .replace(/أ/g, 'ا')
        .replace(/إ/g, 'ا')
        .replace(/آ/g, 'ا')
        .replace(/ى/g, 'ي')
        .replace(/ة/g, 'ه')
        .trim();
    };

    // Process each row (apply same cleaning logic)
    const processedData = data.map((row, index) => {
      const deptOrgCode = String(row['Department Org Code'] || '').trim();
      
      // Sector mapping
      const sectorMapping = {
        '4107001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4108001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4109001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4110001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4122001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4123001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4124001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4125001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4106001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
        '4201001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4202001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4204001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4205001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4210001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4220001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4230001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
        '4301001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4302001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4303001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4306001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4307001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4308001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4309001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4310001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
        '4405001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4406001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4411001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4412001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4413001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '4414001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
        '2231001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2232001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2233001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2241001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2242001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2243001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2251001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2252001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2253001': 'قطاع عمليات انتاج الطاقة الغربي',
        '2311001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2312001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2313001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2321001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2322001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2323001': 'قطاع عمليات انتاج الطاقة الشرقي',
        '2041001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2043001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2046001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2411001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2412001': 'قطاع عمليات انتاج الطاقة الجنوبي',
        '2413001': 'قطاع عمليات انتاج الطاقة الجنوبي'
      };
      
      if (sectorMapping[deptOrgCode]) {
        row['Sector Org Description'] = sectorMapping[deptOrgCode];
      }
      
      // ========================================
      // FALLBACK RULES for Sector Org Description
      // Apply only if Sector Org Description is still empty/blank/"Blank"
      // ========================================
      const rawSectorOrgDesc = row['Sector Org Description'];
      const currentSectorOrgDesc = String(rawSectorOrgDesc || '').trim();
      const divisionOrgCode = String(row['Division Org Code'] || '').trim();
      const highestDeptOrg = String(row['Highest Department Org'] || '').trim();
      
      // Debug log for first 10 rows to verify input columns
      if (index < 10) {
        console.log(`[Sector Fallback Debug - Stats] row=${index} | Sector Org Description (raw)="${rawSectorOrgDesc}" | Division Org Code="${divisionOrgCode}" | Highest Department Org="${highestDeptOrg}"`);
      }
      
      // Check if blank: empty string, whitespace-only, or "Blank" (case-insensitive)
      const isBlankSector = !currentSectorOrgDesc || currentSectorOrgDesc.toLowerCase() === 'blank';
      
      if (isBlankSector) {
        // Rule 1: Southern Distribution - based on Division Org Code
        if (divisionOrgCode === '4400401') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب';
          console.log(`[Sector Fallback Applied - Stats] row=${index} rule=SouthernDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 2: Eastern Distribution - based on Highest Department Org
        else if (highestDeptOrg === '4310001') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي';
          console.log(`[Sector Fallback Applied - Stats] row=${index} rule=EasternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 3: Western Distribution - based on Highest Department Org
        else if (['4210001', '4220001', '4230001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-غربي';
          console.log(`[Sector Fallback Applied - Stats] row=${index} rule=WesternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 4: Central Distribution - based on Highest Department Org
        else if (['4110001', '4126001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى';
          console.log(`[Sector Fallback Applied - Stats] row=${index} rule=CentralDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 5: Western Generation Operations - based on Highest Department Org
        else if (highestDeptOrg === '2027001') {
          row['Sector Org Description'] = 'قطاع عمليات انتاج الطاقة الغربي';
          console.log(`[Sector Fallback Applied - Stats] row=${index} rule=WesternGenerationOps division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
      }
      
      const businessLineDesc = row['Business Line Org Description'] || 'Blank';
      const sectorOrgDesc = row['Sector Org Description'] || 'Blank';
      const highestDeptDesc = row['Highest Department Org Description'] || 'Blank';
      const deptDesc = row['Department Org Description'] || 'Blank';
      const divisionDesc = row['Division Org Description'] || 'Blank';

      let newBusinessLineDesc = businessLineDesc;
   
      const ceoOrgCode = String(row['CEO Org Code'] || '').trim();

      if (ceoOrgCode === '30000001') {
        newBusinessLineDesc = 'الشركة الوطنية لنقل الكهرباء';
      } else if (ceoOrgCode === '91000001') {
        newBusinessLineDesc = 'Others';
      }

      if (ceoOrgCode !== '30000001' && ceoOrgCode !== '91000001') {
        const activityCode = String(row['النشاط'] || '').trim();

        if (['4000001', '4400001', '4200001', '4100001', '4500001', '4300001', '4600001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوزيع وخدمات المشتركين';
        } else if (['2000001', '2100001', '2200001', '2400001', '2600001', '2300001', '2500001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوليد';
        } else if (activityCode === '16000001') {
          newBusinessLineDesc = 'نشاط الخدمات الفنية';
        } else if (activityCode === '1100001') {
          newBusinessLineDesc = 'نشاط الصحة المهنية والسلامة والامن والبيئة';
        } else if (['7000001', '7100001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط الموارد البشرية والخدمات المساندة';
        } else if (['12030001','12010001','12020001','12054001'].includes(activityCode)) {
          newBusinessLineDesc = 'شركة كھرباء السعودیة لتطویر المشاریع';
        
        }
        // NOTE: Activity codes ['11000001', '5000001', '5100001', '10000001'] are classified as "Others"
        // ONLY in statistics aggregation, NOT in the cleaned data. Original value is preserved here.
      }

      const isBlank = (val) => !val || val === 'Blank' || val.trim() === '';

      if (isBlank(businessLineDesc) && newBusinessLineDesc === businessLineDesc) {
        if (isBlank(sectorOrgDesc) && isBlank(highestDeptDesc) && isBlank(deptDesc) && isBlank(divisionDesc)) {
          newBusinessLineDesc = 'N/A';
        }
      }

      const ceoOrgDesc = row['CEO Org Description'] || 'Blank';
      if (ceoOrgDesc.includes('شركة وادي الحلول')) {
        newBusinessLineDesc = 'N/A';
      }

      row['Business Line Org Description'] = newBusinessLineDesc;
      
      // Add اسم المنطقة column
      const regionValue = String(row['المنطقة'] || '').trim();
      if (!regionValue || regionValue === 'Blank' || regionValue === '') {
        row['اسم المنطقة'] = 'N/A';
      } else if (regionValue.includes('6410')) {
        row['اسم المنطقة'] = 'COA';
      } else if (regionValue.includes('6420')) {
        row['اسم المنطقة'] = 'WOA';
      } else if (regionValue.includes('6430')) {
        row['اسم المنطقة'] = 'EOA';
      } else if (regionValue.includes('6440')) {
        row['اسم المنطقة'] = 'SOA';
      } else {
        row['اسم المنطقة'] = 'N/A';
      }
      
      return row;
    });

    // Calculate statistics
    const regions = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'];
    
    const cancelledRows = processedData.filter(row => {
      const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return paymentStatus === 'CANCEL';
    });
    const cancelCount = cancelledRows.length;
    
    const activeRows = processedData.filter(row => {
      const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return paymentStatus !== 'CANCEL';
    });
    
    const totalViolations = activeRows.length;
    
    // ----------------------------------------
    // Business Line × Region Matrix
    // FIXED ROW AND COLUMN ORDER
    // ----------------------------------------
    
    // Fixed row order (Business Lines) - ALWAYS in this exact order
    const fixedBusinessLines = [
      'Generation',
      'National grid',
      'Distribution',
      'PDC',
      'Technical Services',
      'HR',
      'DT',
      'HSSE',
      'Others',
      'N/A'
    ];
    
    // Activity codes that should be classified as "Others" in statistics only
    const othersActivityCodes = ['11000001', '5000001', '5100001', '10000001'];
    
    // Mapping from Arabic Business Line names to fixed English labels
    const businessLineMapping = {
      'نشاط التوليد': 'Generation',
      'الشركة الوطنية لنقل الكهرباء': 'National grid',
      'نشاط التوزيع وخدمات المشتركين': 'Distribution',
      'نشاط الخدمات الفنية': 'Technical Services',
      'نشاط الموارد البشرية والخدمات المساندة': 'HR',
      'نشاط الصحة المهنية والسلامة والامن والبيئة': 'HSSE',
      'Others': 'Others',
      'N/A': 'Others'
    };
    
    // Helper function to get the statistical classification for Business Line
    const getStatisticalBusinessLine = (row) => {
      const activityCode = String(row['النشاط'] || '').trim();
      
      // If activity code is in the "Others" list, classify as "Others" for stats
      if (activityCode && othersActivityCodes.includes(activityCode)) {
        return 'Others';
      }
      
      // Get the actual Business Line Org Description
      const bl = String(row['Business Line Org Description'] || 'N/A').trim();
      
      // Map to fixed English label, default to "Others" if not found
      return businessLineMapping[bl] || 'Others';
    };
    
    // Initialize the matrix with ALL fixed rows and columns (defaulting to 0)
    const businessLineMatrix = {};
    fixedBusinessLines.forEach(bl => {
      businessLineMatrix[bl] = {};
      regions.forEach(region => {
        businessLineMatrix[bl][region] = 0;
      });
      businessLineMatrix[bl]['Total'] = 0;
    });
    
    const regionTotals = {};
    regions.forEach(region => {
      regionTotals[region] = 0;
    });
    regionTotals['Total'] = 0;
    
    // Count violations (maps to fixed English labels)
    activeRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      const region = String(row['اسم المنطقة'] || 'N/A').trim();
      const normalizedRegion = regions.includes(region) ? region : 'N/A';
      
      businessLineMatrix[bl][normalizedRegion]++;
      businessLineMatrix[bl]['Total']++;
      regionTotals[normalizedRegion]++;
      regionTotals['Total']++;
    });
    
    // Build table in FIXED ORDER (always all rows, even if 0)
    const businessLineTable = fixedBusinessLines.map(bl => {
      const row = { 'Business Line Org Description': bl };
      regions.forEach(region => {
        row[region] = businessLineMatrix[bl][region];
      });
      row['Total'] = businessLineMatrix[bl]['Total'];
      return row;
    });
    
    // Add totals row at the end
    const totalsRow = { 'Business Line Org Description': 'Total' };
    regions.forEach(region => {
      totalsRow[region] = regionTotals[region];
    });
    totalsRow['Total'] = regionTotals['Total'];
    businessLineTable.push(totalsRow);
    
    // HSSE Table (GroupLabel × Region)
    const hsseRows = activeRows.filter(row => {
      const bl = String(row['Business Line Org Description'] || '').trim();
      return bl === 'نشاط الصحة المهنية والسلامة والامن والبيئة';
    });
    
    // Helper to check if value is blank
    const isBlankValue = (val) => !val || val === 'Blank' || String(val).trim() === '';
    
    // Get unique GroupLabels using fallback logic
    const hsseGroupLabels = [...new Set(hsseRows.map(row => {
      const sector = String(row['Sector Org Description'] || '').trim();
      const highestDept = String(row['Highest Department Org Description'] || '').trim();
      const division = String(row['Division Org Description'] || '').trim();
      
      if (!isBlankValue(sector)) return sector;
      if (!isBlankValue(highestDept)) return highestDept;
      if (!isBlankValue(division)) return division;
      return 'N/A';
    }))].sort();
    
    // Initialize the HSSE matrix (GroupLabel × Region)
    const hsseMatrix = {};
    hsseGroupLabels.forEach(label => {
      hsseMatrix[label] = {};
      regions.forEach(region => {
        hsseMatrix[label][region] = 0;
      });
      hsseMatrix[label]['Total'] = 0;
    });
    
    // Initialize HSSE region totals
    const hsseRegionTotals = {};
    regions.forEach(region => {
      hsseRegionTotals[region] = 0;
    });
    hsseRegionTotals['Total'] = 0;
    
    // Count HSSE violations per GroupLabel × Region
    hsseRows.forEach(row => {
      const sector = String(row['Sector Org Description'] || '').trim();
      const highestDept = String(row['Highest Department Org Description'] || '').trim();
      const division = String(row['Division Org Description'] || '').trim();
      
      // Apply fallback logic for GroupLabel
      let groupLabel;
      if (!isBlankValue(sector)) {
        groupLabel = sector;
      } else if (!isBlankValue(highestDept)) {
        groupLabel = highestDept;
      } else if (!isBlankValue(division)) {
        groupLabel = division;
      } else {
        groupLabel = 'N/A';
      }
      
      // Get region from "اسم المنطقة", treat blank as N/A
      const regionValue = String(row['اسم المنطقة'] || '').trim();
      const normalizedRegion = (isBlankValue(regionValue) || !regions.includes(regionValue)) ? 'N/A' : regionValue;
      
      if (hsseMatrix[groupLabel]) {
        hsseMatrix[groupLabel][normalizedRegion]++;
        hsseMatrix[groupLabel]['Total']++;
        hsseRegionTotals[normalizedRegion]++;
        hsseRegionTotals['Total']++;
      }
    });
    
    // Build HSSE table array
    const hsseTable = hsseGroupLabels.map(label => {
      const row = { 'Group': label };
      regions.forEach(region => {
        row[region] = hsseMatrix[label][region];
      });
      row['Total'] = hsseMatrix[label]['Total'];
      return row;
    });
    
    // Add totals row
    const hsseTotalsRow = { 'Group': 'Total' };
    regions.forEach(region => {
      hsseTotalsRow[region] = hsseRegionTotals[region];
    });
    hsseTotalsRow['Total'] = hsseRegionTotals['Total'];
    hsseTable.push(hsseTotalsRow);

    // Return structured JSON
    res.json({
      summary: {
        cancelCount: cancelCount,
        totalViolations: totalViolations,
        totalRowsProcessed: processedData.length
      },
      businessLineByRegion: {
        columns: ['Business Line Org Description', ...regions, 'Total'],
        data: businessLineTable
      },
      hsseViolations: {
        columns: ['Group', ...regions, 'Total'],
        data: hsseTable
      }
    });

  } catch (error) {
    console.error('Error calculating statistics:', error);
    res.status(500).json({ error: 'Error calculating statistics', details: error.message });
  }
});

// ========================================
// EXPORT SAHER EXCEL ENDPOINT
// Generates Excel with cleaned data + statistics
// ========================================
app.post('/api/export-saher', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Validate date range inputs
    const { startDate, endDate } = req.body;
    
    if (!startDate || !endDate) {
      return res.status(400).json({ error: 'startDate and endDate are required' });
    }
    
    // Parse dates from ISO format "YYYY-MM-DD" to JS Date objects (date-only)
    const parsedStartDate = new Date(startDate + 'T00:00:00');
    const parsedEndDate = new Date(endDate + 'T00:00:00');
    
    if (isNaN(parsedStartDate.getTime()) || isNaN(parsedEndDate.getTime())) {
      return res.status(400).json({ error: 'Invalid date format. Expected YYYY-MM-DD' });
    }
    
    console.log(`\n📥 Export request: ${startDate} to ${endDate}`);

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    // Date filtering
    const filterStartDate = normalizeToDateOnly(parsedStartDate);
    const filterEndDate = normalizeToDateOnly(parsedEndDate);
    
    const filteredData = data.filter(row => {
      const createdAtValue = row['مُنشأ في'];
      const createdAtDate = parseCreatedAtToDateOnly(createdAtValue);
      if (createdAtDate === null) return false;
      return createdAtDate >= filterStartDate && createdAtDate <= filterEndDate;
    });
    
    console.log(`📅 Filtered rows: ${filteredData.length} of ${data.length}`);

    // Normalize Arabic text helper
    const normalizeArabic = (text) => {
      if (!text) return '';
      return text.replace(/أ/g, 'ا').replace(/إ/g, 'ا').replace(/آ/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').trim();
    };

    // Sector mapping
    const sectorMapping = {
      '4107001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4108001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4109001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4110001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4122001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4123001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4124001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4125001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4106001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4201001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4202001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4204001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4205001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4210001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4220001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4230001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4301001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4302001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4303001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4306001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4307001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4308001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4309001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4310001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4405001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4406001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4411001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4412001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4413001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4414001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '2231001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2232001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2233001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2241001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2242001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2243001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2251001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2252001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2253001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2311001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2312001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2313001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2321001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2322001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2323001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2041001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2043001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2046001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2411001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2412001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2413001': 'قطاع عمليات انتاج الطاقة الجنوبي'
    };

    // Process/clean each row (preserving original Business Line values)
    const processedData = filteredData.map((row, index) => {
      const deptOrgCode = String(row['Department Org Code'] || '').trim();
      
      if (sectorMapping[deptOrgCode]) {
        row['Sector Org Description'] = sectorMapping[deptOrgCode];
      }
      
      // ========================================
      // FALLBACK RULES for Sector Org Description
      // Apply only if Sector Org Description is still empty/blank/"Blank"
      // ========================================
      const rawSectorOrgDesc = row['Sector Org Description'];
      const currentSectorOrgDesc = String(rawSectorOrgDesc || '').trim();
      const divisionOrgCode = String(row['Division Org Code'] || '').trim();
      const highestDeptOrg = String(row['Highest Department Org'] || '').trim();
      
      // Debug log for first 10 rows to verify input columns
      if (index < 10) {
        console.log(`[Sector Fallback Debug - Excel] row=${index} | Sector Org Description (raw)="${rawSectorOrgDesc}" | Division Org Code="${divisionOrgCode}" | Highest Department Org="${highestDeptOrg}"`);
      }
      
      // Check if blank: empty string, whitespace-only, or "Blank" (case-insensitive)
      const isBlankSector = !currentSectorOrgDesc || currentSectorOrgDesc.toLowerCase() === 'blank';
      
      if (isBlankSector) {
        // Rule 1: Southern Distribution - based on Division Org Code
        if (divisionOrgCode === '4400401') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب';
          console.log(`[Sector Fallback Applied - Excel] row=${index} rule=SouthernDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 2: Eastern Distribution - based on Highest Department Org
        else if (highestDeptOrg === '4310001') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي';
          console.log(`[Sector Fallback Applied - Excel] row=${index} rule=EasternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 3: Western Distribution - based on Highest Department Org
        else if (['4210001', '4220001', '4230001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-غربي';
          console.log(`[Sector Fallback Applied - Excel] row=${index} rule=WesternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 4: Central Distribution - based on Highest Department Org
        else if (['4110001', '4126001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى';
          console.log(`[Sector Fallback Applied - Excel] row=${index} rule=CentralDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 5: Western Generation Operations - based on Highest Department Org
        else if (highestDeptOrg === '2027001') {
          row['Sector Org Description'] = 'قطاع عمليات انتاج الطاقة الغربي';
          console.log(`[Sector Fallback Applied - Excel] row=${index} rule=WesternGenerationOps division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
      }
      
      const businessLineDesc = row['Business Line Org Description'] || 'Blank';
      const sectorOrgDesc = row['Sector Org Description'] || 'Blank';
      const highestDeptDesc = row['Highest Department Org Description'] || 'Blank';
      const deptDesc = row['Department Org Description'] || 'Blank';
      const divisionDesc = row['Division Org Description'] || 'Blank';

      let newBusinessLineDesc = businessLineDesc;
      const ceoOrgCode = String(row['CEO Org Code'] || '').trim();

      if (ceoOrgCode === '30000001') {
        newBusinessLineDesc = 'الشركة الوطنية لنقل الكهرباء';
      } else if (ceoOrgCode === '91000001') {
        newBusinessLineDesc = 'Others';
      }

      if (ceoOrgCode !== '30000001' && ceoOrgCode !== '91000001') {
        const activityCode = String(row['النشاط'] || '').trim();
        if (['4000001', '4400001', '4200001', '4100001', '4500001', '4300001', '4600001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوزيع وخدمات المشتركين';
        } else if (['2000001', '2100001', '2200001', '2400001', '2600001', '2300001', '2500001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوليد';
        } else if (activityCode === '16000001') {
          newBusinessLineDesc = 'نشاط الخدمات الفنية';
        } else if (activityCode === '1100001') {
          newBusinessLineDesc = 'نشاط الصحة المهنية والسلامة والامن والبيئة';
        } else if (['7000001', '7100001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط الموارد البشرية والخدمات المساندة';
        }
        // NOTE: Activity codes ['11000001', '5000001', '5100001', '10000001'] are classified as "Others"
        // ONLY in statistics aggregation, NOT in the cleaned data.
      }

      const isBlank = (val) => !val || val === 'Blank' || String(val).trim() === '';
      if (isBlank(businessLineDesc) && newBusinessLineDesc === businessLineDesc) {
        if (isBlank(sectorOrgDesc) && isBlank(highestDeptDesc) && isBlank(deptDesc) && isBlank(divisionDesc)) {
          newBusinessLineDesc = 'N/A';
        }
      }

      const ceoOrgDesc = row['CEO Org Description'] || 'Blank';
      if (ceoOrgDesc.includes('شركة وادي الحلول')) {
        newBusinessLineDesc = 'N/A';
      }

      row['Business Line Org Description'] = newBusinessLineDesc;
      
      // Add اسم المنطقة column
      const regionValue = String(row['المنطقة'] || '').trim();
      if (!regionValue || regionValue === 'Blank' || regionValue === '') {
        row['اسم المنطقة'] = 'N/A';
      } else if (regionValue.includes('6410')) {
        row['اسم المنطقة'] = 'COA';
      } else if (regionValue.includes('6420')) {
        row['اسم المنطقة'] = 'WOA';
      } else if (regionValue.includes('6430')) {
        row['اسم المنطقة'] = 'EOA';
      } else if (regionValue.includes('6440')) {
        row['اسم المنطقة'] = 'SOA';
      } else {
        row['اسم المنطقة'] = 'N/A';
      }
      
      return row;
    });

    // Columns to remove from output
    const columnsToRemove = [
      'وصف المعدات', 'رقم الصنف', 'وصف الصنف', 'CEO Org Description',
      'Department Org Code', 'Cost Center Org Description', 'Department Org Description'
    ];
    processedData.forEach(row => {
      columnsToRemove.forEach(col => { if (row.hasOwnProperty(col)) delete row[col]; });
    });

    // ========================================
    // VALIDATION GATE: Sector Org Description Mapping
    // Block export if any row matches code conditions but has empty Sector Org Description
    // ========================================
    const isEmptySector = (val) => {
      if (val === null || val === undefined) return true;
      const trimmed = String(val).trim();
      return trimmed === '' || trimmed.toLowerCase() === 'blank';
    };
    
    const validationErrors = [];
    processedData.forEach((row, idx) => {
      const sectorOrgDesc = row['Sector Org Description'];
      
      // Only validate if Sector Org Description is empty
      if (isEmptySector(sectorOrgDesc)) {
        const divisionOrgCode = String(row['Division Org Code'] || '').trim();
        const highestDeptOrg = String(row['Highest Department Org'] || '').trim();
        
        // Check Rule 1: Southern Distribution
        if (divisionOrgCode === '4400401') {
          validationErrors.push({
            row: idx,
            rule: 'SouthernDistribution',
            divisionOrgCode,
            highestDeptOrg,
            expected: 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب'
          });
        }
        // Check Rule 2: Eastern Distribution
        else if (highestDeptOrg === '4310001') {
          validationErrors.push({
            row: idx,
            rule: 'EasternDistribution',
            divisionOrgCode,
            highestDeptOrg,
            expected: 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي'
          });
        }
        // Check Rule 3: Western Distribution
        else if (['4210001', '4220001', '4230001'].includes(highestDeptOrg)) {
          validationErrors.push({
            row: idx,
            rule: 'WesternDistribution',
            divisionOrgCode,
            highestDeptOrg,
            expected: 'وحدة أعمال التوزيع وخدمات المشتركين-غربي'
          });
        }
        // Check Rule 4: Central Distribution
        else if (['4110001', '4126001'].includes(highestDeptOrg)) {
          validationErrors.push({
            row: idx,
            rule: 'CentralDistribution',
            divisionOrgCode,
            highestDeptOrg,
            expected: 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى'
          });
        }
        // Check Rule 5: Western Generation Operations
        else if (highestDeptOrg === '2027001') {
          validationErrors.push({
            row: idx,
            rule: 'WesternGenerationOps',
            divisionOrgCode,
            highestDeptOrg,
            expected: 'قطاع عمليات انتاج الطاقة الغربي'
          });
        }
      }
    });
    
    // If validation errors found, abort export
    if (validationErrors.length > 0) {
      console.error(`[Sector Validation FAILED] ${validationErrors.length} rows have unmapped Sector Org Description:`);
      validationErrors.slice(0, 10).forEach(err => {
        console.error(`  Row ${err.row}: rule=${err.rule}, Division=${err.divisionOrgCode}, HighestDept=${err.highestDeptOrg}, Expected="${err.expected}"`);
      });
      if (validationErrors.length > 10) {
        console.error(`  ... and ${validationErrors.length - 10} more errors`);
      }
      
      return res.status(400).json({
        error: 'Saher export blocked: Sector Org Description mapping not fully applied.',
        details: `${validationErrors.length} row(s) match code conditions but have empty Sector Org Description.`,
        affectedRows: validationErrors.slice(0, 10).map(e => ({
          row: e.row,
          rule: e.rule,
          expected: e.expected
        }))
      });
    }
    
    console.log('[Sector Validation PASSED] All Sector Org Description mappings validated successfully.');

    // Get headers and reorder
    const existingHeaders = processedData.length > 0 ? Object.keys(processedData[0]) : [];
    const columnsToMoveToEnd = ['اسم المنطقة', 'Business Line Org Description', 'Sector Org Description', 
                                 'Highest Department Org Description', 'Division Org Description'];
    let reorderedHeaders = existingHeaders.filter(h => !columnsToMoveToEnd.includes(h));
    const columnsToAppend = columnsToMoveToEnd.filter(col => existingHeaders.includes(col));
    reorderedHeaders = [...reorderedHeaders, ...columnsToAppend];

    // ========================================
    // CALCULATE STATISTICS FOR EXPORT
    // ========================================
    const regions = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'];
    
    // Fixed row order for statistics
    const fixedBusinessLines = [
      'Generation', 'National grid', 'Distribution', 'PDC', 'Technical Services',
      'HR', 'DT', 'HSSE', 'Others', 'N/A'
    ];
    
    const othersActivityCodes = ['11000001', '5000001', '5100001', '10000001'];
    const businessLineMapping = {
      'نشاط التوليد': 'Generation',
      'الشركة الوطنية لنقل الكهرباء': 'National grid',
      'نشاط التوزيع وخدمات المشتركين': 'Distribution',
      'نشاط الخدمات الفنية': 'Technical Services',
      'نشاط الموارد البشرية والخدمات المساندة': 'HR',
      'نشاط الصحة المهنية والسلامة والامن والبيئة': 'HSSE',
      'Others': 'Others',
      'N/A': 'N/A'
    };
    
    const getStatisticalBusinessLine = (row) => {
      const activityCode = String(row['النشاط'] || '').trim();
      if (activityCode && othersActivityCodes.includes(activityCode)) return 'Others';
      const bl = String(row['Business Line Org Description'] || 'N/A').trim();
      return businessLineMapping[bl] || 'Others';
    };

    // Count cancels and active rows
    const cancelCount = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return status === 'CANCEL';
    }).length;
    
    const activeRows = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return status !== 'CANCEL';
    });
    const totalViolations = activeRows.length;

    // Calculate weekly violations
    const getLastWednesday = (date) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const dayOfWeek = d.getDay();
      const daysToSubtract = (dayOfWeek + 7 - 3) % 7;
      d.setDate(d.getDate() - daysToSubtract);
      return d;
    };
    const addDays = (date, days) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      d.setDate(d.getDate() + days);
      return d;
    };
    
    const selectedEndDateOnly = normalizeToDateOnly(parsedEndDate);
    const lastWednesday = getLastWednesday(selectedEndDateOnly);
    let weeklyStart, weeklyEnd;
    if (selectedEndDateOnly.getDay() === 3) {
      weeklyStart = addDays(lastWednesday, -7);
      weeklyEnd = addDays(weeklyStart, 6);
    } else {
      weeklyStart = lastWednesday;
      weeklyEnd = selectedEndDateOnly;
    }
    
    let weeklyTotalViolations = 0;
    let weeklyHsseViolations = 0;
    processedData.forEach(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (status === 'CANCEL') return;
      const createdAtDate = parseCreatedAtToDateOnly(row['مُنشأ في']);
      if (createdAtDate && createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd) {
        weeklyTotalViolations++;
        // Check if this is an HSSE violation
        const businessLine = String(row['Business Line Org Description'] || '').trim();
        if (businessLine === 'نشاط الصحة المهنية والسلامة والامن والبيئة') {
          weeklyHsseViolations++;
        }
      }
    });

    // Build SAHER violations matrix with fixed order
    const businessLineMatrix = {};
    fixedBusinessLines.forEach(bl => {
      businessLineMatrix[bl] = {};
      regions.forEach(region => { businessLineMatrix[bl][region] = 0; });
      businessLineMatrix[bl]['Total'] = 0;
    });
    const regionTotals = {};
    regions.forEach(region => { regionTotals[region] = 0; });
    regionTotals['Total'] = 0;
    
    activeRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      const region = String(row['اسم المنطقة'] || 'N/A').trim();
      const normalizedRegion = regions.includes(region) ? region : 'N/A';
      businessLineMatrix[bl][normalizedRegion]++;
      businessLineMatrix[bl]['Total']++;
      regionTotals[normalizedRegion]++;
      regionTotals['Total']++;
    });

    // ========================================
    // CREATE EXCEL WORKBOOK
    // ========================================
    const newWorkbook = XLSX.utils.book_new();
    
    // Sheet1: Cleaned data
    const cleanedSheet = XLSX.utils.json_to_sheet(processedData, { header: reorderedHeaders });
    
    // Apply consistent column formatting (dates and times)
    // This runs on EVERY export to enforce fixed display formats
    applyColumnFormatsByHeader(cleanedSheet, reorderedHeaders, SAHER_COLUMN_FORMATS);
    
    XLSX.utils.book_append_sheet(newWorkbook, cleanedSheet, 'Sheet1');
    
    // Saher Statistic sheet
    const statsSheetData = [];
    statsSheetData.push(['Cancels', cancelCount]);
    statsSheetData.push(['Total Violations', totalViolations]);
    statsSheetData.push(['Weekly Total Violations', weeklyTotalViolations]);
    statsSheetData.push(['Weekly HSSE Violations', weeklyHsseViolations]);
    statsSheetData.push([]);
    statsSheetData.push(['SAHER Violations']);
    statsSheetData.push(['Business Line Org Description', ...regions, 'Total']);
    
    fixedBusinessLines.forEach(bl => {
      const rowData = [bl];
      regions.forEach(region => { rowData.push(businessLineMatrix[bl][region]); });
      rowData.push(businessLineMatrix[bl]['Total']);
      statsSheetData.push(rowData);
    });
    
    // Add totals row
    const totalsRowData = ['Total'];
    regions.forEach(region => { totalsRowData.push(regionTotals[region]); });
    totalsRowData.push(regionTotals['Total']);
    statsSheetData.push(totalsRowData);
    
    const statsSheet = XLSX.utils.aoa_to_sheet(statsSheetData);
    XLSX.utils.book_append_sheet(newWorkbook, statsSheet, 'Saher Statistic');
    
    // ========================================
    // PPT DATA SHEET (NEW - for PowerPoint generation)
    // Contains 4 tables based on WEEKLY data
    // ========================================
    
    // Filter to weekly active rows only (same logic as PPT export)
    const weeklyActiveRows = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (status === 'CANCEL') return false;
      const createdAtDate = parseCreatedAtToDateOnly(row['مُنشأ في']);
      return createdAtDate && createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd;
    });
    
    // --- Table 1: Saher By BL (Business Line summary) ---
    const pptBlCounts = {};
    fixedBusinessLines.forEach(bl => { pptBlCounts[bl] = 0; });
    
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      if (pptBlCounts[bl] !== undefined) {
        pptBlCounts[bl]++;
      }
    });
    
    // --- Table 2: Saher By Area (BL × Region matrix) ---
    const pptBlRegionMatrix = {};
    fixedBusinessLines.forEach(bl => {
      pptBlRegionMatrix[bl] = {};
      regions.forEach(region => { pptBlRegionMatrix[bl][region] = 0; });
    });
    
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      const region = String(row['اسم المنطقة'] || 'N/A').trim();
      const normalizedRegion = regions.includes(region) ? region : 'N/A';
      if (pptBlRegionMatrix[bl]) {
        pptBlRegionMatrix[bl][normalizedRegion]++;
      }
    });
    
    // --- Table 3: DIS – Departments ---
    const disDeptCounts = {};
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      if (bl !== 'Distribution') return;
      
      const highestDept = String(row['Highest Department Org Description'] || '').trim();
      if (!highestDept || highestDept === 'Blank' || highestDept === '') return;
      disDeptCounts[highestDept] = (disDeptCounts[highestDept] || 0) + 1;
    });
    const sortedDisDepts = Object.entries(disDeptCounts).sort((a, b) => b[1] - a[1]);
    
    // --- Table 4: NG – Departments ---
    const ngDeptCounts = {};
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      if (bl !== 'National grid') return;
      
      const highestDept = String(row['Highest Department Org Description'] || '').trim();
      if (!highestDept || highestDept === 'Blank' || highestDept === '') return;
      ngDeptCounts[highestDept] = (ngDeptCounts[highestDept] || 0) + 1;
    });
    const sortedNgDepts = Object.entries(ngDeptCounts).sort((a, b) => b[1] - a[1]);
    
    // Build PPT Data sheet content
    const pptSheetData = [];
    
    // Table 1: Saher By BL
    pptSheetData.push(['Table 1: Saher By BL']);
    pptSheetData.push(['Business Line', 'Violations']);
    fixedBusinessLines.forEach(bl => {
      pptSheetData.push([bl, pptBlCounts[bl]]);
    });
    pptSheetData.push([]); // Empty row separator
    
    // Table 2: Saher By Area
    pptSheetData.push(['Table 2: Saher By Area']);
    pptSheetData.push(['Business Line', 'WOA', 'EOA', 'COA', 'SOA', 'N/A']);
    fixedBusinessLines.forEach(bl => {
      pptSheetData.push([
        bl,
        pptBlRegionMatrix[bl]['WOA'],
        pptBlRegionMatrix[bl]['EOA'],
        pptBlRegionMatrix[bl]['COA'],
        pptBlRegionMatrix[bl]['SOA'],
        pptBlRegionMatrix[bl]['N/A']
      ]);
    });
    pptSheetData.push([]); // Empty row separator
    
    // Table 3: DIS – Departments
    pptSheetData.push(['Table 3: DIS – Departments']);
    pptSheetData.push(['Department Name', 'Count']);
    sortedDisDepts.forEach(([deptName, count]) => {
      pptSheetData.push([deptName, count]);
    });
    if (sortedDisDepts.length === 0) {
      pptSheetData.push(['(No data)', 0]);
    }
    pptSheetData.push([]); // Empty row separator
    
    // Table 4: NG – Departments
    pptSheetData.push(['Table 4: NG – Departments']);
    pptSheetData.push(['Department Name', 'Count']);
    sortedNgDepts.forEach(([deptName, count]) => {
      pptSheetData.push([deptName, count]);
    });
    if (sortedNgDepts.length === 0) {
      pptSheetData.push(['(No data)', 0]);
    }
    
    const pptDataSheet = XLSX.utils.aoa_to_sheet(pptSheetData);
    XLSX.utils.book_append_sheet(newWorkbook, pptDataSheet, 'PPT Data');
    
    console.log(`📊 PPT Data sheet: ${weeklyActiveRows.length} weekly rows, DIS depts: ${sortedDisDepts.length}, NG depts: ${sortedNgDepts.length}`);
    
    // Generate and send Excel buffer
    const buffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=saher_cleaned.xlsx');
    res.send(buffer);
    
    console.log(`✅ Excel exported: ${processedData.length} rows, ${cancelCount} cancels, ${totalViolations} violations`);

  } catch (error) {
    console.error('Error exporting file:', error);
    res.status(500).json({ error: 'Error exporting file', details: error.message });
  }
});

// ========================================
// EXPORT SAHER POWERPOINT ENDPOINT
// Generates a PPTX file for Saher Weekly Report
// ========================================
app.post('/api/export-saher-ppt', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Validate date range inputs
    const { startDate, endDate } = req.body;
    
    if (!startDate || !endDate) {
      return res.status(400).json({ error: 'startDate and endDate are required' });
    }
    
    // Parse dates from ISO format "YYYY-MM-DD" to JS Date objects (date-only)
    const parsedStartDate = new Date(startDate + 'T00:00:00');
    const parsedEndDate = new Date(endDate + 'T00:00:00');
    
    if (isNaN(parsedStartDate.getTime()) || isNaN(parsedEndDate.getTime())) {
      return res.status(400).json({ error: 'Invalid date format. Expected YYYY-MM-DD' });
    }
    
    console.log(`\n📊 Generating Saher PowerPoint report: ${startDate} to ${endDate}`);

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    // Date filtering
    const filterStartDate = normalizeToDateOnly(parsedStartDate);
    const filterEndDate = normalizeToDateOnly(parsedEndDate);
    
    const filteredData = data.filter(row => {
      const createdAtValue = row['مُنشأ في'];
      const createdAtDate = parseCreatedAtToDateOnly(createdAtValue);
      if (createdAtDate === null) return false;
      return createdAtDate >= filterStartDate && createdAtDate <= filterEndDate;
    });
    
    console.log(`📅 Filtered rows: ${filteredData.length} of ${data.length}`);

    // Sector mapping (same as Excel export)
    const sectorMapping = {
      '4107001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4108001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4109001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4110001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4122001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4123001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4124001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4125001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4106001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4201001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4202001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4204001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4205001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4210001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4220001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4230001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4301001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4302001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4303001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4306001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4307001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4308001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4309001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4310001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4405001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4406001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4411001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4412001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4413001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4414001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '2231001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2232001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2233001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2241001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2242001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2243001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2251001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2252001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2253001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2311001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2312001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2313001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2321001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2322001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2323001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2041001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2043001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2046001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2411001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2412001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2413001': 'قطاع عمليات انتاج الطاقة الجنوبي'
    };

    // Process/clean each row (same logic as Excel export)
    const processedData = filteredData.map((row, index) => {
      const deptOrgCode = String(row['Department Org Code'] || '').trim();
      
      if (sectorMapping[deptOrgCode]) {
        row['Sector Org Description'] = sectorMapping[deptOrgCode];
      }
      
      // ========================================
      // FALLBACK RULES for Sector Org Description
      // Apply only if Sector Org Description is still empty/blank/"Blank"
      // ========================================
      const rawSectorOrgDesc = row['Sector Org Description'];
      const currentSectorOrgDesc = String(rawSectorOrgDesc || '').trim();
      const divisionOrgCode = String(row['Division Org Code'] || '').trim();
      const highestDeptOrg = String(row['Highest Department Org'] || '').trim();
      
      // Debug log for first 10 rows to verify input columns
      if (index < 10) {
        console.log(`[Sector Fallback Debug - PPT] row=${index} | Sector Org Description (raw)="${rawSectorOrgDesc}" | Division Org Code="${divisionOrgCode}" | Highest Department Org="${highestDeptOrg}"`);
      }
      
      // Check if blank: empty string, whitespace-only, or "Blank" (case-insensitive)
      const isBlankSector = !currentSectorOrgDesc || currentSectorOrgDesc.toLowerCase() === 'blank';
      
      if (isBlankSector) {
        // Rule 1: Southern Distribution - based on Division Org Code
        if (divisionOrgCode === '4400401') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب';
          console.log(`[Sector Fallback Applied - PPT] row=${index} rule=SouthernDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 2: Eastern Distribution - based on Highest Department Org
        else if (highestDeptOrg === '4310001') {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي';
          console.log(`[Sector Fallback Applied - PPT] row=${index} rule=EasternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 3: Western Distribution - based on Highest Department Org
        else if (['4210001', '4220001', '4230001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-غربي';
          console.log(`[Sector Fallback Applied - PPT] row=${index} rule=WesternDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 4: Central Distribution - based on Highest Department Org
        else if (['4110001', '4126001'].includes(highestDeptOrg)) {
          row['Sector Org Description'] = 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى';
          console.log(`[Sector Fallback Applied - PPT] row=${index} rule=CentralDistribution division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
        // Rule 5: Western Generation Operations - based on Highest Department Org
        else if (highestDeptOrg === '2027001') {
          row['Sector Org Description'] = 'قطاع عمليات انتاج الطاقة الغربي';
          console.log(`[Sector Fallback Applied - PPT] row=${index} rule=WesternGenerationOps division=${divisionOrgCode} highestDept=${highestDeptOrg}`);
        }
      }
      
      const businessLineDesc = row['Business Line Org Description'] || 'Blank';
      const sectorOrgDesc = row['Sector Org Description'] || 'Blank';
      const highestDeptDesc = row['Highest Department Org Description'] || 'Blank';
      const deptDesc = row['Department Org Description'] || 'Blank';
      const divisionDesc = row['Division Org Description'] || 'Blank';

      let newBusinessLineDesc = businessLineDesc;
      const ceoOrgCode = String(row['CEO Org Code'] || '').trim();

      if (ceoOrgCode === '30000001') {
        newBusinessLineDesc = 'الشركة الوطنية لنقل الكهرباء';
      } else if (ceoOrgCode === '91000001') {
        newBusinessLineDesc = 'Others';
      }

      if (ceoOrgCode !== '30000001' && ceoOrgCode !== '91000001') {
        const activityCode = String(row['النشاط'] || '').trim();
        if (['4000001', '4400001', '4200001', '4100001', '4500001', '4300001', '4600001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوزيع وخدمات المشتركين';
        } else if (['2000001', '2100001', '2200001', '2400001', '2600001', '2300001', '2500001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط التوليد';
        } else if (activityCode === '16000001') {
          newBusinessLineDesc = 'نشاط الخدمات الفنية';
        } else if (activityCode === '1100001') {
          newBusinessLineDesc = 'نشاط الصحة المهنية والسلامة والامن والبيئة';
        } else if (['7000001', '7100001'].includes(activityCode)) {
          newBusinessLineDesc = 'نشاط الموارد البشرية والخدمات المساندة';
        }
      }

      const isBlank = (val) => !val || val === 'Blank' || String(val).trim() === '';
      if (isBlank(businessLineDesc) && newBusinessLineDesc === businessLineDesc) {
        if (isBlank(sectorOrgDesc) && isBlank(highestDeptDesc) && isBlank(deptDesc) && isBlank(divisionDesc)) {
          newBusinessLineDesc = 'N/A';
        }
      }

      const ceoOrgDesc = row['CEO Org Description'] || 'Blank';
      if (ceoOrgDesc.includes('شركة وادي الحلول')) {
        newBusinessLineDesc = 'N/A';
      }

      row['Business Line Org Description'] = newBusinessLineDesc;
      
      // Add اسم المنطقة column
      const regionValue = String(row['المنطقة'] || '').trim();
      if (!regionValue || regionValue === 'Blank' || regionValue === '') {
        row['اسم المنطقة'] = 'N/A';
      } else if (regionValue.includes('6410')) {
        row['اسم المنطقة'] = 'COA';
      } else if (regionValue.includes('6420')) {
        row['اسم المنطقة'] = 'WOA';
      } else if (regionValue.includes('6430')) {
        row['اسم المنطقة'] = 'EOA';
      } else if (regionValue.includes('6440')) {
        row['اسم المنطقة'] = 'SOA';
      } else {
        row['اسم المنطقة'] = 'N/A';
      }
      
      return row;
    });

    // Count cancels and active rows (exclude CANCEL before any grouping)
    const cancelCount = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return status === 'CANCEL';
    }).length;
    
    const activeRows = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      return status !== 'CANCEL';
    });
    const totalViolations = activeRows.length;

    // Calculate weekly date range
    const getLastWednesday = (date) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const dayOfWeek = d.getDay();
      const daysToSubtract = (dayOfWeek + 7 - 3) % 7;
      d.setDate(d.getDate() - daysToSubtract);
      return d;
    };
    const addDays = (date, days) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      d.setDate(d.getDate() + days);
      return d;
    };
    const formatDateYMD = (date) => {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    };
    
    const selectedEndDateOnly = normalizeToDateOnly(parsedEndDate);
    const lastWednesday = getLastWednesday(selectedEndDateOnly);
    let weeklyStart, weeklyEnd;
    if (selectedEndDateOnly.getDay() === 3) {
      weeklyStart = addDays(lastWednesday, -7);
      weeklyEnd = addDays(weeklyStart, 6);
    } else {
      weeklyStart = lastWednesday;
      weeklyEnd = selectedEndDateOnly;
    }
    
    // Count weekly violations (excluding CANCEL, within weekly range)
    let weeklyTotalViolations = 0;
    let weeklyHsseViolations = 0;
    processedData.forEach(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (status === 'CANCEL') return;
      const createdAtDate = parseCreatedAtToDateOnly(row['مُنشأ في']);
      if (createdAtDate && createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd) {
        weeklyTotalViolations++;
        const businessLine = String(row['Business Line Org Description'] || '').trim();
        if (businessLine === 'نشاط الصحة المهنية والسلامة والامن والبيئة') {
          weeklyHsseViolations++;
        }
      }
    });

    console.log(`📊 Stats - Cancels: ${cancelCount}, Total: ${totalViolations}, Weekly: ${weeklyTotalViolations}`);
    
    // ========================================
    // CALCULATE WEEKLY DATA FOR PPT CHARTS
    // (Uses SAME logic as /api/process-saher)
    // ========================================
    
    // Fixed regions (same as main endpoint)
    const regions = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'];
    
    // Fixed Business Lines in order (same as main endpoint - English labels)
    const fixedBusinessLines = ['Generation', 'National grid', 'Distribution', 'PDC', 'Technical Services', 'HR', 'DT', 'HSSE', 'Others'];
    
    // Activity codes that should be classified as "Others"
    const othersActivityCodes = ['11000001', '5000001', '5100001', '10000001'];
    
    // Mapping from Arabic Business Line names to fixed English labels (same as main endpoint)
    const businessLineMapping = {
      'نشاط التوليد': 'Generation',
      'الشركة الوطنية لنقل الكهرباء': 'National grid',
      'نشاط التوزيع وخدمات المشتركين': 'Distribution',
      'نشاط الخدمات الفنية': 'Technical Services',
      'نشاط الموارد البشرية والخدمات المساندة': 'HR',
      'نشاط الصحة المهنية والسلامة والامن والبيئة': 'HSSE',
      'Others': 'Others',
      'N/A': 'Others'
    };
    
    // Helper function to get the statistical classification for Business Line (same as main endpoint)
    const getStatisticalBusinessLine = (row) => {
      const activityCode = String(row['النشاط'] || '').trim();
      if (activityCode && othersActivityCodes.includes(activityCode)) return 'Others';
      const bl = String(row['Business Line Org Description'] || 'N/A').trim();
      return businessLineMapping[bl] || 'Others';
    };
    
    // Filter to weekly active rows only (for charts)
    const weeklyActiveRows = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (status === 'CANCEL') return false;
      const createdAtDate = parseCreatedAtToDateOnly(row['مُنشأ في']);
      return createdAtDate && createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd;
    });
    
    // Count weekly cancellations
    const weeklyCancelCount = processedData.filter(row => {
      const status = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (status !== 'CANCEL') return false;
      const createdAtDate = parseCreatedAtToDateOnly(row['مُنشأ في']);
      return createdAtDate && createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd;
    }).length;
    
    // Initialize Business Line × Region matrix for weekly data
    const weeklyBlMatrix = {};
    fixedBusinessLines.forEach(bl => {
      weeklyBlMatrix[bl] = { COA: 0, WOA: 0, EOA: 0, SOA: 0, 'N/A': 0, Total: 0 };
    });
    
    // Count weekly violations per Business Line × Region (using same mapping as main endpoint)
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      const region = String(row['اسم المنطقة'] || 'N/A').trim();
      const normalizedRegion = regions.includes(region) ? region : 'N/A';
      
      if (weeklyBlMatrix[bl]) {
        weeklyBlMatrix[bl][normalizedRegion]++;
        weeklyBlMatrix[bl]['Total']++;
      }
    });
    
    // Calculate DIS (Distribution) department violations (weekly)
    const disDeptCounts = {};
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      if (bl !== 'Distribution') return;
      
      const highestDept = String(row['Highest Department Org Description'] || 'N/A').trim();
      const deptName = highestDept === 'Blank' || highestDept === '' ? 'N/A' : highestDept;
      disDeptCounts[deptName] = (disDeptCounts[deptName] || 0) + 1;
    });
    
    // Calculate NG (National grid) department violations (weekly)
    const ngDeptCounts = {};
    weeklyActiveRows.forEach(row => {
      const bl = getStatisticalBusinessLine(row);
      if (bl !== 'National grid') return;
      
      const highestDept = String(row['Highest Department Org Description'] || 'N/A').trim();
      const deptName = highestDept === 'Blank' || highestDept === '' ? 'N/A' : highestDept;
      ngDeptCounts[deptName] = (ngDeptCounts[deptName] || 0) + 1;
    });
    
    // Sort departments by count (descending)
    const sortedDisDepts = Object.entries(disDeptCounts).sort((a, b) => b[1] - a[1]);
    const sortedNgDepts = Object.entries(ngDeptCounts).sort((a, b) => b[1] - a[1]);
    
    // ========================================
    // LOG PPT STATS FOR DEBUGGING
    // ========================================
    console.log('\n📊 === PPT EXPORT STATS (before building slides) ===');
    console.log(`Weekly Active Rows: ${weeklyActiveRows.length}`);
    console.log(`Weekly Cancels: ${weeklyCancelCount}`);
    console.log(`Weekly HSSE: ${weeklyHsseViolations}`);
    console.log('Business Line Matrix:');
    console.table(weeklyBlMatrix);
    console.log(`DIS Departments: ${sortedDisDepts.length} entries`);
    if (sortedDisDepts.length > 0) console.log('Top DIS depts:', sortedDisDepts.slice(0, 5));
    console.log(`NG Departments: ${sortedNgDepts.length} entries`);
    if (sortedNgDepts.length > 0) console.log('Top NG depts:', sortedNgDepts.slice(0, 5));
    console.log('='.repeat(50));
    
    // ========================================
    // CREATE POWERPOINT PRESENTATION
    // ========================================
    
    const pptx = new PptxGenJS();
    
    // Custom slide size: 33.87 cm × 19.05 cm (widescreen 16:9)
    // Convert cm to inches: inches = cm / 2.54
    const slideWidthInches = 33.87 / 2.54;  // = 13.334645669 inches
    const slideHeightInches = 19.05 / 2.54; // = 7.5 inches
    
    pptx.defineLayout({ 
      name: 'CUSTOM_16x9', 
      width: slideWidthInches, 
      height: slideHeightInches 
    });
    pptx.layout = 'CUSTOM_16x9';
    
    pptx.title = 'Saher Weekly Report';
    pptx.author = 'Saher Analytics System';
    
    // ========================================
    // SHARED SLIDE BACKGROUND (All Slides)
    // Diagonal gradient: top-left light → bottom-right dark navy
    // ========================================
    const slideBackground = {
      color: '061C3A',  // Fallback color
      gradientType: 'linear',
      gradientAngle: 45,  // Top-left to bottom-right diagonal
      gradientColors: [
        { color: '9FB3C8', position: 0 },     // Stop 1: Light steel blue/grey-blue
        { color: '0B2A55', position: 38 },    // Stop 2: Dark navy blue (~35-40%)
        { color: '061C3A', position: 100 }    // Stop 3: Very dark navy blue
      ]
    };
    
    // Color palette (matching dark theme)
    const colors = {
      background: '0F172A',
      cardBg: '1E293B',
      primary: '60A5FA',
      secondary: 'A78BFA',
      success: '34D399',
      warning: 'FBBF24',
      danger: 'F87171',
      textPrimary: 'F8FAFC',
      textSecondary: '94A3B8',
      regionColors: {
        COA: '60A5FA',
        WOA: '34D399',
        EOA: 'FBBF24',
        SOA: 'A78BFA',
        'N/A': '94A3B8'
      }
    };
    
    // ========================================
    // SLIDE 1: SAHER Violations – BL by OA
    // Pixel-perfect reference design
    // ========================================
    const slide1 = pptx.addSlide();
    
    // 1) Slide Background - Use shared gradient (all slides consistent)
    slide1.background = slideBackground;
    
    // 2) Header Title - "SAHER Violations – BL by OA" (Top Center)
    slide1.addText('SAHER Violations – BL by OA', {
      x: 0,
      y: 0.25,
      w: '100%',
      h: 0.6,
      fontSize: 24,
      fontFace: 'Calibri',
      color: 'FFFFFF',
      bold: true,
      align: 'center'
    });
    
    // 3) Top-right Home Icon (subtle gray, small)
    slide1.addText('⌂', {
      x: 12.2,
      y: 0.15,
      w: 0.4,
      h: 0.4,
      fontSize: 18,
      fontFace: 'Arial',
      color: '8899AA',
      align: 'center'
    });
    
    // 4) Main Dark Container (Rounded Rectangle behind charts)
    // Outer container - dark navy card with shadow effect
    const containerX = 0.3;
    const containerY = 0.92;
    const containerW = 12.73;
    const containerH = 4.4;
    
    slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: containerX,
      y: containerY,
      w: containerW,
      h: containerH,
      fill: { color: '0A1628' },
      line: { color: '0A1628', width: 0 },
      rectRadius: 0.15,
      shadow: {
        type: 'outer',
        blur: 10,
        offset: 3,
        angle: 45,
        color: '000000',
        opacity: 0.4
      }
    });
    
    // Inner panel (slightly darker recessed area)
    slide1.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: containerX + 0.1,
      y: containerY + 0.1,
      w: containerW - 0.2,
      h: containerH - 0.2,
      fill: { color: '081020' },
      line: { color: '081020', width: 0 },
      rectRadius: 0.12
    });
    
    // 5) Donut Chart (Left side inside container)
    // Prepare donut data - Business Line totals
    const donutBLs = fixedBusinessLines.filter(bl => weeklyBlMatrix[bl] && weeklyBlMatrix[bl].Total > 0);
    
    // Donut chart colors - icy blues with one coral accent
    const donutColors = [
      'A8D8EA',  // Light icy blue
      '7EC8E3',  // Sky blue
      '5DADE2',  // Medium blue
      'D98880',  // Muted coral/red (accent)
      '85C1E9',  // Soft blue
      'AED6F1',  // Pale blue
      '76D7C4',  // Teal
      'BFC9CA',  // Gray-blue
      '95A5A6'   // Slate gray
    ];
    
    if (donutBLs.length > 0) {
      const donutChartData = donutBLs.map(bl => ({
        name: bl,
        labels: [bl],
        values: [weeklyBlMatrix[bl].Total]
      }));
      
      // Position: 0.5cm = 0.197in horizontal, 2.33cm = 0.917in vertical from slide top-left
      slide1.addChart(pptx.charts.DOUGHNUT, donutChartData, {
        x: 0.2,
        y: 0.92,
        w: 5.5,
        h: 4.3,
        chartColors: donutColors.slice(0, donutBLs.length),
        holeSize: 62,
        showLabel: false,
        showValue: true,
        showPercent: false,
        dataBorder: { pt: 0, color: '0A1628' },
        showLegend: true,
        legendPos: 'l',
        legendColor: 'CCCCCC',
        legendFontFace: 'Calibri',
        legendFontSize: 9,
        dataLabelColor: 'FFFFFF',
        dataLabelFontFace: 'Calibri',
        dataLabelFontSize: 10,
        dataLabelPosition: 'outEnd'
      });
    } else {
      slide1.addText('No data available', {
        x: 0.5,
        y: 2.8,
        w: 5,
        h: 0.5,
        fontSize: 14,
        fontFace: 'Calibri',
        color: '8899AA',
        align: 'center'
      });
    }
    
    // 6) Clustered Column Chart (Right side inside container)
    // Only include regions with actual data for cleaner display
    const chartRegions = ['WOA', 'EOA', 'COA', 'SOA'];
    
    // Filter to BLs that have data for cleaner chart
    const chartBLs = fixedBusinessLines.filter(bl => weeklyBlMatrix[bl] && weeklyBlMatrix[bl].Total > 0);
    
    // Column chart series colors - icy blue tones
    const columnColors = [
      '5DADE2',  // WOA - Medium blue
      '2E86AB',  // EOA - Darker blue
      '7FDBFF',  // COA - Bright cyan
      'D6EAF8'   // SOA - Very light icy blue
    ];
    
    if (chartBLs.length > 0) {
      const columnChartData = chartRegions.map((region, idx) => ({
        name: region,
        labels: chartBLs,
        values: chartBLs.map(bl => weeklyBlMatrix[bl][region] || 0)
      }));
      
      // Position: 12.36cm = 4.866in horizontal, 2.45cm = 0.965in vertical
      slide1.addChart(pptx.charts.BAR, columnChartData, {
        x: 5.8,
        y: 0.97,
        w: 7.0,
        h: 4.2,
        barDir: 'col',
        barGrouping: 'clustered',
        barGapWidthPct: 50,
        chartColors: columnColors,
        dataBorder: { pt: 0 },
        showValue: true,
        dataLabelColor: 'FFFFFF',
        dataLabelFontFace: 'Calibri',
        dataLabelFontSize: 8,
        dataLabelPosition: 'outEnd',
        catAxisLabelColor: 'CCCCCC',
        catAxisLabelFontFace: 'Calibri',
        catAxisLabelFontSize: 9,
        catAxisLabelRotate: 0,
        catAxisLineShow: true,
        catAxisLineColor: '445566',
        catAxisLineSize: 0.5,
        valAxisHidden: true,
        valAxisDisplayUnit: 'none',
        catGridLine: { style: 'none' },
        valGridLine: { style: 'none' },
        showLegend: true,
        legendPos: 'b',
        legendColor: 'CCCCCC',
        legendFontFace: 'Calibri',
        legendFontSize: 9
      });
    }
    
    // 8) Bottom Footer - Classification text (green, centered)
    /*slide1.addText('Public Internal - (عام داخلي)', {
      x: 0,
      y: 5.15,
      w: '100%',
      h: 0.3,
      fontSize: 10,
      fontFace: 'Calibri',
      color: '2ECC71',
      align: 'center'
    });*/
    
    // 9) Bottom-right Down Arrow Icon (subtle gray)
    slide1.addText('▼', {
      x: 12.3,
      y: 5.1,
      w: 0.3,
      h: 0.3,
      fontSize: 14,
      fontFace: 'Arial',
      color: '667788',
      align: 'center'
    });
    
    // ========================================
    // SLIDE 2: Business Line Overview
    // ========================================
    const slide2 = pptx.addSlide();
    slide2.background = slideBackground;
    
    // Title
    slide2.addText('SAHER Violations – Business Line by Region', {
      x: 0.5,
      y: 0.3,
      w: 12.33,
      h: 0.6,
      fontSize: 28,
      fontFace: 'Arial',
      color: colors.textPrimary,
      bold: true,
      align: 'center'
    });
    
    // Prepare data for donut chart (Business Line totals)
    // Filter to only business lines with data
    const activeBLs = fixedBusinessLines.filter(bl => weeklyBlMatrix[bl] && weeklyBlMatrix[bl].Total > 0);
    
    const donutData = activeBLs.map(bl => ({
      name: bl,
      labels: [bl],
      values: [weeklyBlMatrix[bl].Total]
    }));
    
    console.log(`📊 Donut chart data: ${donutData.length} business lines with data`);
    
    // Donut chart (left side)
    if (donutData.length > 0) {
      slide2.addChart(pptx.charts.DOUGHNUT, donutData, {
        x: 0.3,
        y: 1.0,
        w: 5.0,
        h: 4.0,
        chartColors: ['3B82F6', '10B981', 'F59E0B', '8B5CF6', 'EF4444', '06B6D4', 'EC4899', '64748B', '475569'],
        holeSize: 50,
        showLabel: true,
        showValue: true,
        showPercent: false,
        dataBorder: { pt: 1, color: colors.background },
        showLegend: true,
        legendPos: 'b',
        legendColor: colors.textSecondary,
        legendFontSize: 10
      });
    } else {
      slide2.addText('No violations data available for the selected weekly window.', {
        x: 0.5,
        y: 2.5,
        w: 5.0,
        h: 0.5,
        fontSize: 14,
        fontFace: 'Arial',
        color: colors.textSecondary,
        align: 'center'
      });
    }
    
    // Prepare data for grouped bar chart (BL × Region)
    const barChartData = regions.map(region => ({
      name: region,
      labels: activeBLs,
      values: activeBLs.map(bl => weeklyBlMatrix[bl][region])
    }));
    
    console.log(`📊 Bar chart data: ${activeBLs.length} business lines × ${regions.length} regions`);
    
    // Grouped bar chart (right side)
    if (barChartData.length > 0 && activeBLs.length > 0) {
      slide2.addChart(pptx.charts.BAR, barChartData, {
        x: 5.5,
        y: 1.0,
        w: 7.5,
        h: 4.2,
        barDir: 'bar',
        barGrouping: 'clustered',
        chartColors: [colors.regionColors.COA, colors.regionColors.WOA, colors.regionColors.EOA, colors.regionColors.SOA, colors.regionColors['N/A']],
        dataBorder: { pt: 0.5, color: colors.background },
        showValue: true,
        dataLabelColor: colors.textPrimary,
        dataLabelFontSize: 9,
        dataLabelPosition: 'outEnd',
        catAxisLabelColor: colors.textSecondary,
        catAxisLabelFontSize: 10,
        valAxisLabelColor: colors.textSecondary,
        valAxisLabelFontSize: 9,
        catGridLine: { style: 'none' },
        valGridLine: { color: '64748B', style: 'dash' },
        showLegend: true,
        legendPos: 't',
        legendColor: colors.textSecondary,
        legendFontSize: 10
      });
    }
    
    // ========================================
    // SLIDE 3: DIS Departments
    // ========================================
    const slide3 = pptx.addSlide();
    slide3.background = slideBackground;
    
    // Title
    slide3.addText('DIS – Department Violations', {
      x: 0.5,
      y: 0.3,
      w: 12.33,
      h: 0.6,
      fontSize: 28,
      fontFace: 'Arial',
      color: colors.textPrimary,
      bold: true,
      align: 'center'
    });
    
    // Subtitle
    slide3.addText('Distribution (نشاط التوزيع وخدمات المشتركين)', {
      x: 0.5,
      y: 0.85,
      w: 12.33,
      h: 0.4,
      fontSize: 14,
      fontFace: 'Arial',
      color: colors.textSecondary,
      align: 'center'
    });
    
    console.log(`📊 DIS chart: ${sortedDisDepts.length} departments`);
    
    if (sortedDisDepts.length > 0) {
      // Take top 15 departments for readability
      const topDisDepts = sortedDisDepts.slice(0, 15);
      
      const disChartData = [{
        name: 'Violations',
        labels: topDisDepts.map(([name]) => name.length > 35 ? name.substring(0, 35) + '...' : name),
        values: topDisDepts.map(([, count]) => count)
      }];
      
      slide3.addChart(pptx.charts.BAR, disChartData, {
        x: 0.5,
        y: 1.3,
        w: 12.33,
        h: 4.0,
        barDir: 'bar',
        chartColors: [colors.primary],
        dataBorder: { pt: 0.5, color: colors.background },
        showValue: true,
        dataLabelColor: colors.textPrimary,
        dataLabelFontSize: 10,
        dataLabelPosition: 'outEnd',
        catAxisLabelColor: colors.textSecondary,
        catAxisLabelFontSize: 9,
        valAxisLabelColor: colors.textSecondary,
        valAxisLabelFontSize: 9,
        catGridLine: { style: 'none' },
        valGridLine: { color: '64748B', style: 'dash' },
        showLegend: false
      });
    } else {
      slide3.addText('No DIS department violations in the selected weekly window.', {
        x: 0.5,
        y: 2.5,
        w: 12.33,
        h: 0.5,
        fontSize: 16,
        fontFace: 'Arial',
        color: colors.textSecondary,
        align: 'center'
      });
    }
    
    // ========================================
    // SLIDE 4: NG Departments
    // ========================================
    const slide4 = pptx.addSlide();
    slide4.background = slideBackground;
    
    // Title
    slide4.addText('NG – Department Violations', {
      x: 0.5,
      y: 0.3,
      w: 12.33,
      h: 0.6,
      fontSize: 28,
      fontFace: 'Arial',
      color: colors.textPrimary,
      bold: true,
      align: 'center'
    });
    
    // Subtitle
    slide4.addText('National Grid (الشركة الوطنية لنقل الكهرباء)', {
      x: 0.5,
      y: 0.85,
      w: 12.33,
      h: 0.4,
      fontSize: 14,
      fontFace: 'Arial',
      color: colors.textSecondary,
      align: 'center'
    });
    
    console.log(`📊 NG chart: ${sortedNgDepts.length} departments`);
    
    if (sortedNgDepts.length > 0) {
      // Take top 15 departments for readability
      const topNgDepts = sortedNgDepts.slice(0, 15);
      
      const ngChartData = [{
        name: 'Violations',
        labels: topNgDepts.map(([name]) => name.length > 35 ? name.substring(0, 35) + '...' : name),
        values: topNgDepts.map(([, count]) => count)
      }];
      
      slide4.addChart(pptx.charts.BAR, ngChartData, {
        x: 0.5,
        y: 1.3,
        w: 12.33,
        h: 4.0,
        barDir: 'bar',
        chartColors: [colors.secondary],
        dataBorder: { pt: 0.5, color: colors.background },
        showValue: true,
        dataLabelColor: colors.textPrimary,
        dataLabelFontSize: 10,
        dataLabelPosition: 'outEnd',
        catAxisLabelColor: colors.textSecondary,
        catAxisLabelFontSize: 9,
        valAxisLabelColor: colors.textSecondary,
        valAxisLabelFontSize: 9,
        catGridLine: { style: 'none' },
        valGridLine: { color: '64748B', style: 'dash' },
        showLegend: false
      });
    } else {
      slide4.addText('No NG department violations in the selected weekly window.', {
        x: 0.5,
        y: 2.5,
        w: 12.33,
        h: 0.5,
        fontSize: 16,
        fontFace: 'Arial',
        color: colors.textSecondary,
        align: 'center'
      });
    }
    
    // Generate the PPTX as a buffer
    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });
    
    // Set response headers for PPTX download
    const pptxFileName = `saher_weekly_report_${startDate}_${endDate}.pptx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename=${pptxFileName}`);
    res.send(pptxBuffer);
    
    console.log(`✅ PowerPoint generated successfully: ${pptxFileName}`);

  } catch (error) {
    console.error('Error generating PowerPoint:', error);
    res.status(500).json({ error: 'Error generating PowerPoint', details: error.message });
  }
});

// ========================================
// EXPORT WEEKLY PDF ENDPOINT
// Generates a PDF report for the weekly window only
// ========================================
app.post('/api/export-weekly-pdf', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    // Validate date range inputs
    const { startDate, endDate } = req.body;
    
    if (!startDate || !endDate) {
      return res.status(400).json({ error: 'startDate and endDate are required' });
    }
    
    // Parse dates from ISO format "YYYY-MM-DD" to JS Date objects
    const parsedStartDate = new Date(startDate + 'T00:00:00');
    const parsedEndDate = new Date(endDate + 'T00:00:00');
    
    if (isNaN(parsedStartDate.getTime())) {
      return res.status(400).json({ error: 'Invalid startDate format. Expected YYYY-MM-DD' });
    }
    
    if (isNaN(parsedEndDate.getTime())) {
      return res.status(400).json({ error: 'Invalid endDate format. Expected YYYY-MM-DD' });
    }
    
    console.log(`\n📄 === Generating Weekly PDF Report ===`);
    console.log(`📅 Selected date range: ${startDate} to ${endDate}`);

    // Read the Excel file from buffer
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    
    // Target "Sheet1" specifically for Saher file
    const sheetName = workbook.SheetNames.includes('Sheet1') ? 'Sheet1' : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON for processing
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: 'Blank' });

    if (data.length === 0) {
      return res.status(400).json({ error: 'No data found in file' });
    }

    // ========================================
    // WEEKLY WINDOW CALCULATION
    // ========================================
    
    // Helper: Get the most recent Wednesday on or before a given date
    const getLastWednesday = (date) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const dayOfWeek = d.getDay(); // 0=Sun, 1=Mon, ..., 3=Wed, ..., 6=Sat
      const daysToSubtract = (dayOfWeek + 7 - 3) % 7; // Days since last Wednesday
      d.setDate(d.getDate() - daysToSubtract);
      return d;
    };
    
    // Helper: Add days to a date
    const addDays = (date, days) => {
      const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      d.setDate(d.getDate() + days);
      return d;
    };
    
    // Compute weekly range based on selectedEndDate
    const selectedEndDateOnly = normalizeToDateOnly(parsedEndDate);
    const lastWednesday = getLastWednesday(selectedEndDateOnly);
    
    let weeklyStart, weeklyEnd;
    
    // Check if selectedEndDate is a Wednesday (day 3)
    if (selectedEndDateOnly.getDay() === 3) {
      // selectedEndDate IS Wednesday: use previous full week (Wed to Tue)
      weeklyStart = addDays(lastWednesday, -7);
      weeklyEnd = addDays(weeklyStart, 6); // Tuesday
    } else {
      // selectedEndDate is NOT Wednesday: weekly-to-date from last Wednesday
      weeklyStart = lastWednesday;
      weeklyEnd = selectedEndDateOnly;
    }
    
    // Format date as DD/MM/YYYY for display
    const formatDateDMY = (date) => {
      const d = String(date.getDate()).padStart(2, '0');
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const y = date.getFullYear();
      return `${d}/${m}/${y}`;
    };
    
    // Calculate ISO week number from the weekly window END date
    const getISOWeekNumber = (date) => {
      const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
      const dayNum = d.getUTCDay() || 7;
      d.setUTCDate(d.getUTCDate() + 4 - dayNum);
      const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
      return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    };
    
    const weekNumber = getISOWeekNumber(weeklyEnd);
    
    console.log(`📅 Weekly window: ${formatDateDMY(weeklyStart)} to ${formatDateDMY(weeklyEnd)}`);
    console.log(`📅 ISO Week number: ${weekNumber}`);

    // ========================================
    // PROCESS AND CLEAN DATA (same logic as existing)
    // ========================================
    
    // Sector mapping for Department Org Code
    const sectorMapping = {
      '4107001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4108001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4109001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4110001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4122001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4123001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4124001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4125001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4106001': 'وحدة أعمال التوزيع وخدمات المشتركين-وسطى',
      '4201001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4202001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4204001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4205001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4210001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4220001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4230001': 'وحدة أعمال التوزيع وخدمات المشتركين-غربي',
      '4301001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4302001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4303001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4306001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4307001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4308001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4309001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4310001': 'وحدة أعمال التوزيع وخدمات المشتركين-شرقي',
      '4405001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4406001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4411001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4412001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4413001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '4414001': 'وحدة أعمال التوزيع وخدمات المشتركين-جنوب',
      '2231001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2232001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2233001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2241001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2242001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2243001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2251001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2252001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2253001': 'قطاع عمليات انتاج الطاقة الغربي',
      '2311001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2312001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2313001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2321001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2322001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2323001': 'قطاع عمليات انتاج الطاقة الشرقي',
      '2041001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2043001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2046001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2411001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2412001': 'قطاع عمليات انتاج الطاقة الجنوبي',
      '2413001': 'قطاع عمليات انتاج الطاقة الجنوبي'
    };
    
    // Activity codes that should be classified as "Others" in statistics only
    const othersActivityCodes = ['11000001', '5000001', '5100001', '10000001'];
    
    // Business line mapping from Arabic to English
    const businessLineMapping = {
      'نشاط التوليد': 'Generation',
      'الشركة الوطنية لنقل الكهرباء': 'National grid',
      'نشاط التوزيع وخدمات المشتركين': 'Distribution',
      'نشاط الخدمات الفنية': 'Technical Services',
      'نشاط الموارد البشرية والخدمات المساندة': 'HR',
      'نشاط الصحة المهنية والسلامة والامن والبيئة': 'HSSE',
      'Others': 'Others',
      'N/A': 'Others'
    };
    
    // Helper to get statistical business line classification
    const getStatisticalBusinessLine = (row) => {
      const activityCode = String(row['النشاط'] || '').trim();
      
      // Check if activity code maps to "Others" for statistics
      if (othersActivityCodes.includes(activityCode)) {
        return 'Others';
      }
      
      // Get the Business Line Org Description
      let businessLineDesc = String(row['Business Line Org Description'] || '').trim();
      
      // Apply CEO Org Code mapping first
      const ceoOrgCode = String(row['CEO Org Code'] || '').trim();
      if (ceoOrgCode === '30000001') {
        businessLineDesc = 'الشركة الوطنية لنقل الكهرباء';
      } else if (ceoOrgCode === '91000001') {
        businessLineDesc = 'Others';
      }
      
      // Apply activity code mapping
      if (ceoOrgCode !== '30000001' && ceoOrgCode !== '91000001') {
        if (['4000001', '4400001', '4200001', '4100001', '4500001', '4300001', '4600001'].includes(activityCode)) {
          businessLineDesc = 'نشاط التوزيع وخدمات المشتركين';
        } else if (['2000001', '2100001', '2200001', '2400001', '2600001', '2300001', '2500001'].includes(activityCode)) {
          businessLineDesc = 'نشاط التوليد';
        } else if (activityCode === '16000001') {
          businessLineDesc = 'نشاط الخدمات الفنية';
        } else if (activityCode === '1100001') {
          businessLineDesc = 'نشاط الصحة المهنية والسلامة والامن والبيئة';
        } else if (['7000001', '7100001'].includes(activityCode)) {
          businessLineDesc = 'نشاط الموارد البشرية والخدمات المساندة';
        }
      }
      
      // Map to English label
      return businessLineMapping[businessLineDesc] || 'Others';
    };
    
    // Helper to get region from row
    const getRegion = (row) => {
      const regionValue = String(row['المنطقة'] || '').trim();
      
      if (!regionValue || regionValue === 'Blank' || regionValue === '') {
        return 'N/A';
      } else if (regionValue.includes('6410')) {
        return 'COA';
      } else if (regionValue.includes('6420')) {
        return 'WOA';
      } else if (regionValue.includes('6430')) {
        return 'EOA';
      } else if (regionValue.includes('6440')) {
        return 'SOA';
      }
      return 'N/A';
    };

    // ========================================
    // FILTER DATA FOR WEEKLY WINDOW
    // ========================================
    
    // Filter rows:
    // 1. Within weekly window (inclusive)
    // 2. Exclude CANCEL rows
    const weeklyData = data.filter(row => {
      // Exclude CANCEL rows first
      const paymentStatus = String(row['حالة سداد المخالفة'] || '').trim().toUpperCase();
      if (paymentStatus === 'CANCEL') {
        return false;
      }
      
      // Parse the date from the row
      const createdAtValue = row['مُنشأ في'];
      const createdAtDate = parseCreatedAtToDateOnly(createdAtValue);
      
      if (createdAtDate === null) {
        return false; // Skip invalid dates
      }
      
      // Check if within weekly range (inclusive)
      return createdAtDate >= weeklyStart && createdAtDate <= weeklyEnd;
    });
    
    console.log(`📊 Weekly data rows (after filtering): ${weeklyData.length}`);

    // ========================================
    // COUNT UNIQUE VIOLATIONS
    // ========================================
    
    // Count by unique violation number "رقم المخالفة في نظام"
    const violationNumberColumn = 'رقم المخالفة في نظام';
    const uniqueViolations = new Set();
    
    weeklyData.forEach(row => {
      const violationNum = row[violationNumberColumn];
      if (violationNum && violationNum !== 'Blank' && violationNum !== '') {
        uniqueViolations.add(String(violationNum).trim());
      }
    });
    
    // Total weekly violations (by unique violation number, or fallback to row count)
    const totalWeeklyViolations = uniqueViolations.size > 0 ? uniqueViolations.size : weeklyData.length;
    
    console.log(`📊 Total weekly SAHER violations (unique): ${totalWeeklyViolations}`);

    // ========================================
    // BUILD BUSINESS LINE × REGION MATRIX
    // ========================================
    
    const fixedBusinessLines = [
      'Generation',
      'National grid',
      'Distribution',
      'PDC',
      'Technical Services',
      'HR',
      'DT',
      'HSSE',
      'Others'
    ];
    
    const regions = ['COA', 'WOA', 'EOA', 'SOA', 'N/A'];
    
    // Initialize matrix with zeros
    const matrix = {};
    fixedBusinessLines.forEach(bl => {
      matrix[bl] = {};
      regions.forEach(r => {
        matrix[bl][r] = 0;
      });
    });
    
    // Count violations by business line and region
    // Use unique violation numbers to avoid duplicates
    const countedViolations = new Set();
    
    weeklyData.forEach(row => {
      const violationNum = row[violationNumberColumn];
      const violationKey = violationNum && violationNum !== 'Blank' && violationNum !== '' 
        ? String(violationNum).trim() 
        : `row_${Math.random()}`; // Fallback for rows without violation number
      
      // Skip if already counted
      if (countedViolations.has(violationKey)) {
        return;
      }
      countedViolations.add(violationKey);
      
      const businessLine = getStatisticalBusinessLine(row);
      const region = getRegion(row);
      
      if (matrix[businessLine] && regions.includes(region)) {
        matrix[businessLine][region]++;
      }
    });
    
    // Calculate row totals and grand total
    const rowTotals = {};
    fixedBusinessLines.forEach(bl => {
      rowTotals[bl] = regions.reduce((sum, r) => sum + matrix[bl][r], 0);
    });
    
    const columnTotals = {};
    regions.forEach(r => {
      columnTotals[r] = fixedBusinessLines.reduce((sum, bl) => sum + matrix[bl][r], 0);
    });
    
    const grandTotal = Object.values(rowTotals).reduce((sum, val) => sum + val, 0);

    // ========================================
    // GENERATE PDF USING PDFKIT
    // ========================================
    
    const doc = new PDFDocument({ 
      size: 'A4', 
      margin: 50,
      bufferPages: true
    });
    
    // Collect PDF data into buffer
    const chunks = [];
    doc.on('data', chunk => chunks.push(chunk));
    
    const pdfPromise = new Promise((resolve, reject) => {
      doc.on('end', () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);
    });

    // PDF CONTENT
    
    // A) Title: Week{N}
    doc.fontSize(32)
       .font('Helvetica-Bold')
       .fillColor('#1a1a2e')
       .text(`Week${weekNumber}`, { align: 'center' });
    
    doc.moveDown(0.5);
    
    // B) Weekly Window line
    doc.fontSize(14)
       .font('Helvetica')
       .fillColor('#4a4a6a')
       .text(`Weekly Window: ${formatDateDMY(weeklyStart)} - ${formatDateDMY(weeklyEnd)}`, { align: 'center' });
    
    doc.moveDown(1.5);
    
    // C) KPI line
    doc.fontSize(16)
       .font('Helvetica-Bold')
       .fillColor('#0a7c42')
       .text(`Total SAHER Violations (Weekly): ${totalWeeklyViolations}`, { align: 'center' });
    
    doc.moveDown(2);
    
    // D) Table: SAHER Violations (Weekly) by Business Line & Region
    doc.fontSize(14)
       .font('Helvetica-Bold')
       .fillColor('#1a1a2e')
       .text('SAHER Violations (Weekly) by Business Line & Region', { align: 'left' });
    
    doc.moveDown(0.5);
    
    // Table layout
    const tableTop = doc.y;
    const tableLeft = 50;
    const colWidths = [130, 50, 50, 50, 50, 50, 60]; // Business Line, COA, WOA, EOA, SOA, N/A, Total
    const rowHeight = 25;
    const headerHeight = 30;
    
    // Draw table header
    const headers = ['Business Line', 'COA', 'WOA', 'EOA', 'SOA', 'N/A', 'Total'];
    let xPos = tableLeft;
    
    // Header background
    doc.rect(tableLeft, tableTop, colWidths.reduce((a, b) => a + b, 0), headerHeight)
       .fill('#1a1a2e');
    
    // Header text
    doc.font('Helvetica-Bold')
       .fontSize(10)
       .fillColor('#ffffff');
    
    headers.forEach((header, i) => {
      doc.text(header, xPos + 5, tableTop + 8, { 
        width: colWidths[i] - 10, 
        align: i === 0 ? 'left' : 'center' 
      });
      xPos += colWidths[i];
    });
    
    // Draw table rows
    let yPos = tableTop + headerHeight;
    const allRows = [...fixedBusinessLines, 'Total'];
    
    allRows.forEach((bl, rowIndex) => {
      const isTotal = bl === 'Total';
      const bgColor = isTotal ? '#e8f5e9' : (rowIndex % 2 === 0 ? '#f8f9fa' : '#ffffff');
      
      // Row background
      doc.rect(tableLeft, yPos, colWidths.reduce((a, b) => a + b, 0), rowHeight)
         .fill(bgColor);
      
      // Row border
      doc.rect(tableLeft, yPos, colWidths.reduce((a, b) => a + b, 0), rowHeight)
         .stroke('#ddd');
      
      // Row text
      xPos = tableLeft;
      const values = isTotal 
        ? [bl, columnTotals['COA'], columnTotals['WOA'], columnTotals['EOA'], columnTotals['SOA'], columnTotals['N/A'], grandTotal]
        : [bl, matrix[bl]['COA'], matrix[bl]['WOA'], matrix[bl]['EOA'], matrix[bl]['SOA'], matrix[bl]['N/A'], rowTotals[bl]];
      
      doc.font(isTotal ? 'Helvetica-Bold' : 'Helvetica')
         .fontSize(9)
         .fillColor(isTotal ? '#0a7c42' : '#333333');
      
      values.forEach((val, i) => {
        doc.text(String(val), xPos + 5, yPos + 7, { 
          width: colWidths[i] - 10, 
          align: i === 0 ? 'left' : 'center' 
        });
        xPos += colWidths[i];
      });
      
      yPos += rowHeight;
    });
    
    // Draw vertical lines for columns
    xPos = tableLeft;
    for (let i = 0; i <= headers.length; i++) {
      doc.moveTo(xPos, tableTop)
         .lineTo(xPos, yPos)
         .stroke('#ddd');
      xPos += colWidths[i] || 0;
    }
    
    doc.moveDown(4);
    
    // E) Placeholder for MVA section
    doc.fontSize(12)
       .font('Helvetica-Oblique')
       .fillColor('#888888')
       .text('MVA Section: Coming soon (will be added later).', { align: 'left' });
    
    // Finalize PDF
    doc.end();
    
    // Wait for PDF to be generated
    const pdfBuffer = await pdfPromise;
    
    // Set response headers
    const pdfFileName = `Week${weekNumber}.pdf`;
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${pdfFileName}"`);
    res.send(pdfBuffer);
    
    console.log(`✅ PDF generated successfully: ${pdfFileName}`);

  } catch (error) {
    console.error('Error generating PDF:', error);
    res.status(500).json({ error: 'Error generating PDF', details: error.message });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`🚀 Backend server is running on http://localhost:${PORT}`);
  console.log(`📊 Health check available at http://localhost:${PORT}/health`);
  console.log(`📁 Saher file processing at POST http://localhost:${PORT}/api/process-saher`);
  console.log(`📈 Saher statistics (JSON) at POST http://localhost:${PORT}/api/saher-stats`);
  console.log(`📥 Saher export (Excel) at POST http://localhost:${PORT}/api/export-saher`);
  console.log(`📑 Saher export (PowerPoint) at POST http://localhost:${PORT}/api/export-saher-ppt`);
  console.log(`📄 Saher export (Weekly PDF) at POST http://localhost:${PORT}/api/export-weekly-pdf`);
});

