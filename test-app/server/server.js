import cors from 'cors';
import express from 'express';
import fs from 'fs';
import morgan from 'morgan';
import multer from 'multer';
import path from 'path';
import xlsx from 'xlsx';
import { fileURLToPath } from 'url';
import PizZip from 'pizzip';
import { DOMParser } from '@xmldom/xmldom';
import ExcelJS from 'exceljs';
import { execSync } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 4000;

// Configure middleware
app.use(cors({ origin: 'http://localhost:5173', credentials: true }));
app.use(express.json());
app.use(morgan('dev'));

const upload = multer({
  storage: multer.memoryStorage(),
  fileFilter: (req, file, cb) => {
    const isXlsx =
      file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.originalname.toLowerCase().endsWith('.xlsx');
    const isDocx =
      file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
      file.originalname.toLowerCase().endsWith('.docx');

    if (!isXlsx && !isDocx) {
      cb(new Error('Only .xlsx and .docx files are allowed'));
      return;
    }
    cb(null, true);
  },
});

const staticDir = path.join(__dirname, 'static');
const staticWorkbookPath = path.join(
  staticDir,
  'Knowledge Base_KMD Repository (S4HANA 2023).xlsx',
);

// Cache for Fiori releases API
let fioriReleasesCache = { data: null, timestamp: 0 };
const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes in milliseconds

function cleanBrackets(value) {
  if (value === undefined || value === null) return value;
  return String(value).replace(/\s*\([^)]*\)/g, '').trim();
}

function normalizeKey(row, l1Idx = 0, l2Idx = 1, l3Idx = 2) {
  return [row[l1Idx], row[l2Idx], row[l3Idx]]
    .map((v) => (v === undefined || v === null ? '' : String(v).trim().toLowerCase()))
    .join('|');
}

// Helper function to find column index by header name(s)
function findColumnIndex(headers, searchTerms) {
  if (!Array.isArray(searchTerms)) {
    searchTerms = [searchTerms];
  }
  return headers.findIndex((h) => {
    if (!h) return false;
    const headerLower = String(h).toLowerCase().trim();
    return searchTerms.some((term) => headerLower.includes(term.toLowerCase()));
  });
}

// Helper function to get KMD knowledge base column indices
function getKmdColumnIndices(headerRow) {
  const indices = {
    level1: findColumnIndex(headerRow, ['level 1', 'l1']),
    level2: findColumnIndex(headerRow, ['level 2', 'l2']),
    level3: findColumnIndex(headerRow, ['level 3 - process', 'level 3', 'l3']),
    simplificationItem: findColumnIndex(headerRow, ['simplification item', 'simpl']),
    workstream: findColumnIndex(headerRow, ['workstream']),
    kmdDescription: findColumnIndex(headerRow, ['kmd description', 'kmd desc']),
    migrationApproach: findColumnIndex(headerRow, ['migration approach', 'migration']),
    link: findColumnIndex(headerRow, ['link', 'url']),
  };

  console.log('KMD Column indices found:', indices);
  return indices;
}

// Extract text from XML node, only bold text if onlyBold=true
function extractTextFromNode(node, onlyBold = false) {
  if (!node) return '';

  let text = '';
  const childNodes = node.childNodes;

  for (let i = 0; i < childNodes.length; i++) {
    const child = childNodes[i];

    if (child.nodeName === 'w:r') {
      // Check if this run is bold
      const rPr = child.getElementsByTagName('w:rPr')[0];
      const isBold = rPr && rPr.getElementsByTagName('w:b').length > 0;

      if (!onlyBold || isBold) {
        const textNodes = child.getElementsByTagName('w:t');
        for (let j = 0; j < textNodes.length; j++) {
          text += textNodes[j].textContent || '';
        }
      }
    } else if (child.nodeName === 'w:p') {
      text += extractTextFromNode(child, onlyBold);
    }
  }

  return text.trim();
}

// Extract SAP Note hyperlinks from a cell (returns URL if hyperlink, or text otherwise)
function extractSapNoteFromCell(cell, relsDoc) {
  if (!cell) return '';

  const sapNoteLinks = [];

  // Look for hyperlinks in the cell (w:hyperlink elements)
  const hyperlinks = cell.getElementsByTagName('w:hyperlink');
  for (let i = 0; i < hyperlinks.length; i++) {
    const hyperlink = hyperlinks[i];
    const rId = hyperlink.getAttribute('r:id');

    // Get the hyperlink text (display text)
    const linkText = extractTextFromNode(hyperlink, false).trim();

    // If we have relsDoc, get the actual URL
    if (relsDoc && rId) {
      const relationships = relsDoc.getElementsByTagName('Relationship');
      for (let j = 0; j < relationships.length; j++) {
        const rel = relationships[j];
        if (rel.getAttribute('Id') === rId) {
          const targetUrl = rel.getAttribute('Target') || '';
          if (targetUrl) {
            // Return the full URL for the hyperlink
            sapNoteLinks.push(targetUrl);
            console.log(`    Found SAP Note hyperlink: text="${linkText}", url="${targetUrl}"`);
          }
        }
      }
    }

    // If no URL found but we have text that looks like a note number
    if (sapNoteLinks.length === 0 && linkText) {
      const noteMatch = linkText.match(/\b(\d{6,7})\b/);
      if (noteMatch) {
        sapNoteLinks.push(noteMatch[1]);
      }
    }
  }

  // If no hyperlinks found, try to get plain text
  if (sapNoteLinks.length === 0) {
    const cellText = extractTextFromNode(cell, false).trim();
    if (cellText) {
      sapNoteLinks.push(cellText);
    }
  }

  return sapNoteLinks.join(', ');
}

// Extract tables from .docx file (sections 3.1.1 and 3.1.2)
function extractTablesFromDocx(buffer) {
  const zip = new PizZip(buffer);
  const xml = zip.file('word/document.xml').asText();
  const doc = new DOMParser().parseFromString(xml, 'text/xml');

  // Load relationships document for hyperlink URLs
  let relsDoc = null;
  try {
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (relsFile) {
      const relsXml = relsFile.asText();
      relsDoc = new DOMParser().parseFromString(relsXml, 'text/xml');
    }
  } catch (e) {
    console.log('Could not load relationships document:', e.message);
  }

  const paragraphs = doc.getElementsByTagName('w:p');
  const tables = doc.getElementsByTagName('w:tbl');

  let currentSection = '';
  let extractedData = [];

  // Find section headings and associated tables
  for (let i = 0; i < paragraphs.length; i++) {
    const para = paragraphs[i];
    const paraText = extractTextFromNode(para);

    // Check if this paragraph is a section heading
    if (paraText.includes('3.1.1') && paraText.toLowerCase().includes('relevant')) {
      currentSection = '3.1.1';
      console.log('Found section 3.1.1:', paraText);
    } else if (paraText.includes('3.1.2') && paraText.toLowerCase().includes('relevance')) {
      currentSection = '3.1.2';
      console.log('Found section 3.1.2:', paraText);
    }
  }

  // Extract tables (simplified - will need to associate tables with sections properly)
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    const rows = table.getElementsByTagName('w:tr');

    if (rows.length === 0) continue;

    // Get header row
    const headerCells = rows[0].getElementsByTagName('w:tc');
    const headers = [];
    for (let j = 0; j < headerCells.length; j++) {
      headers.push(extractTextFromNode(headerCells[j]).toLowerCase());
    }

    // Check if this table has the columns we need
    const titleIdx = headers.findIndex(h => h.includes('title'));
    const effortIdx = headers.findIndex(h => h.includes('effort'));
    const categoryIdx = headers.findIndex(h => h.includes('category'));
    // Look for exact "sap note" column
    const sapNoteIdx = headers.findIndex(h => h === 'sap note' || h.includes('sap note'));

    if (titleIdx === -1) continue; // Skip tables without Title column

    console.log(`Table ${i}: Headers:`, headers);
    console.log(`  Title index: ${titleIdx}, Effort index: ${effortIdx}, Category index: ${categoryIdx}, SAP Note index: ${sapNoteIdx}`);

    // Extract data rows
    for (let r = 1; r < rows.length; r++) {
      const cells = rows[r].getElementsByTagName('w:tc');

      // Extract SAP Note from dedicated column
      let sapNote = '';
      if (sapNoteIdx >= 0 && cells[sapNoteIdx]) {
        // First try to get hyperlink/note number
        sapNote = extractSapNoteFromCell(cells[sapNoteIdx], relsDoc);
        // If no note found via hyperlink extraction, get plain text
        if (!sapNote) {
          sapNote = extractTextFromNode(cells[sapNoteIdx], false);
        }
      }

      const rowData = {
        title: titleIdx >= 0 && cells[titleIdx] ? extractTextFromNode(cells[titleIdx], true) : '', // Only bold text
        effortRanking: effortIdx >= 0 && cells[effortIdx] ? extractTextFromNode(cells[effortIdx], false) : '',
        category: categoryIdx >= 0 && cells[categoryIdx] ? extractTextFromNode(cells[categoryIdx], false) : '',
        sapNote: sapNote,
      };

      // Only add rows with title
      if (rowData.title) {
        extractedData.push(rowData);
        console.log(`  Row ${r}: Title="${rowData.title.substring(0, 40)}...", SAP Note="${sapNote || 'N/A'}"`);
      }
    }
  }

  console.log(`Extracted ${extractedData.length} rows from .docx`);
  return extractedData;
}

// Map extracted .docx data with KMD file
function mapWithKMD(extractedData) {
  if (!fs.existsSync(staticWorkbookPath)) {
    throw new Error('KMD file not found');
  }

  const kmdWb = xlsx.readFile(staticWorkbookPath);
  const kmdSheetName = kmdWb.SheetNames.find(n => n.trim() === 'KMD Repository - 2023') || kmdWb.SheetNames[0];
  const kmdSheet = kmdWb.Sheets[kmdSheetName];
  const kmdRows = xlsx.utils.sheet_to_json(kmdSheet, { header: 1, defval: '' });

  // Get column indices by header names
  const kmdHeader = kmdRows[0] || [];
  const colIdx = getKmdColumnIndices(kmdHeader);

  // Build lookup map: simplification item -> array of { workstream, kmdDescription }
  // Store arrays to handle multiple KMD matches for same simplification item
  const kmdMap = new Map();
  kmdRows.slice(1).forEach((row) => {
    const simplificationItem = colIdx.simplificationItem >= 0
      ? (row[colIdx.simplificationItem] || '').toString().trim().toLowerCase()
      : '';
    const workstream = colIdx.workstream >= 0
      ? (row[colIdx.workstream] || '').toString().trim()
      : '';
    const kmdDescription = colIdx.kmdDescription >= 0
      ? (row[colIdx.kmdDescription] || '').toString().trim()
      : '';

    if (simplificationItem) {
      if (!kmdMap.has(simplificationItem)) {
        kmdMap.set(simplificationItem, []);
      }
      kmdMap.get(simplificationItem).push({ workstream, kmdDescription });
    }
  });

  console.log(`Built KMD lookup map with ${kmdMap.size} entries`);

  // Map extracted data - create multiple output rows for multiple matches
  const mappedData = [];
  extractedData.forEach((item, idx) => {
    const titleKey = item.title.trim().toLowerCase();
    const matches = kmdMap.get(titleKey);

    if (matches && matches.length > 0) {
      // Create output row for each matching KMD row
      matches.forEach((match) => {
        mappedData.push({
          title: item.title,
          effortRanking: item.effortRanking,
          category: item.category,
          sapNote: item.sapNote,
          workstream: match.workstream,
          kmdDescription: match.kmdDescription,
        });
      });
      console.log(`✓ Matched: "${item.title}" -> ${matches.length} KMD row(s), SAP Note: "${item.sapNote || 'N/A'}"`);
    } else {
      console.log(`✗ No match for: "${item.title}"`);
    }
  });

  console.log(`Mapped ${mappedData.length} out of ${extractedData.length} rows`);
  return mappedData;
}

// Filter data for "Simplification Items" sheet
function filterSimplificationItems(allItems) {
  const filtered = allItems.filter(item => {
    // Remove low effort rankings
    const effortLower = (item.effortRanking || '').toLowerCase();
    const isLow = effortLower.includes('low');

    // Remove null/empty workstream and kmd description
    const hasWorkstream = item.workstream && item.workstream.trim() !== '' && item.workstream.toLowerCase() !== 'na';
    const hasKmdDescription = item.kmdDescription && item.kmdDescription.trim() !== '' && item.kmdDescription.toLowerCase() !== 'na';

    return !isLow && hasWorkstream && hasKmdDescription;
  });

  console.log(`Filtered: ${filtered.length} out of ${allItems.length} items (removed ${allItems.length - filtered.length})`);
  return filtered;
}

// Create pivot table from simplification items
function createPivotTable(simplificationItems) {
  // Group by workstream and category/effort ranking
  const pivot = {};
  const categories = new Set();

  simplificationItems.forEach(item => {
    const workstream = item.workstream || 'Unknown';
    const category = item.category || item.effortRanking || 'Unknown';

    categories.add(category);

    if (!pivot[workstream]) {
      pivot[workstream] = {};
    }
    if (!pivot[workstream][category]) {
      pivot[workstream][category] = 0;
    }
    pivot[workstream][category]++;
  });

  // Convert to array format for Excel
  const categoryArray = Array.from(categories).sort();
  const rows = [];

  // Header row
  rows.push(['Workstream', ...categoryArray, 'Total']);

  // Data rows
  let grandTotal = 0;
  const categoryTotals = categoryArray.map(() => 0);

  Object.keys(pivot).sort().forEach(workstream => {
    const row = [workstream];
    let rowTotal = 0;

    categoryArray.forEach((category, idx) => {
      const count = pivot[workstream][category] || 0;
      row.push(count);
      rowTotal += count;
      categoryTotals[idx] += count;
    });

    row.push(rowTotal);
    grandTotal += rowTotal;
    rows.push(row);
  });

  // Total row
  rows.push(['Total', ...categoryTotals, grandTotal]);

  console.log(`Created pivot table: ${rows.length - 1} workstreams x ${categoryArray.length} categories`);
  return rows;
}

// Create KMD Dispositions from simplification items
function createKMDDispositions(simplificationItems) {
  if (!fs.existsSync(staticWorkbookPath)) {
    throw new Error('KMD file not found');
  }

  // Read KMD file
  const kmdWb = xlsx.readFile(staticWorkbookPath);
  const kmdSheetName = kmdWb.SheetNames.find(n => n.trim() === 'KMD Repository - 2023') || kmdWb.SheetNames[0];
  const kmdSheet = kmdWb.Sheets[kmdSheetName];
  const kmdRows = xlsx.utils.sheet_to_json(kmdSheet, { header: 1, defval: '' });

  // Get column indices by header names
  const kmdHeader = kmdRows[0] || [];
  const colIdx = getKmdColumnIndices(kmdHeader);

  // Build lookup map: KMD Description -> array of { migrationApproach, link }
  // Store arrays to handle multiple KMD matches for same description
  const kmdMap = new Map();
  kmdRows.slice(1).forEach((row) => {
    const kmdDescription = colIdx.kmdDescription >= 0
      ? (row[colIdx.kmdDescription] || '').toString().trim().toLowerCase()
      : '';
    const migrationApproach = colIdx.migrationApproach >= 0
      ? (row[colIdx.migrationApproach] || '').toString().trim()
      : '';
    const link = colIdx.link >= 0
      ? (row[colIdx.link] || '').toString().trim()
      : '';

    if (kmdDescription) {
      if (!kmdMap.has(kmdDescription)) {
        kmdMap.set(kmdDescription, []);
      }
      kmdMap.get(kmdDescription).push({ migrationApproach, link });
    }
  });

  console.log(`Built KMD dispositions lookup map with ${kmdMap.size} entries`);

  // Get unique Workstream + KMD Description + Migration Approach + Link combinations
  const seen = new Set();
  const dispositions = [];

  simplificationItems.forEach(item => {
    const kmdKey = item.kmdDescription.trim().toLowerCase();
    const matches = kmdMap.get(kmdKey);

    if (matches && matches.length > 0) {
      // Create disposition row for each matching migration approach/link combination
      matches.forEach((match) => {
        const uniqueKey = `${item.workstream}|${item.kmdDescription}|${match.migrationApproach}|${match.link}`;
        if (!seen.has(uniqueKey)) {
          seen.add(uniqueKey);
          dispositions.push({
            workstream: item.workstream,
            kmdDescription: item.kmdDescription,
            migrationApproach: match.migrationApproach,
            link: match.link
          });
        }
      });
    } else {
      // No match found, add with empty values
      const uniqueKey = `${item.workstream}|${item.kmdDescription}||`;
      if (!seen.has(uniqueKey)) {
        seen.add(uniqueKey);
        dispositions.push({
          workstream: item.workstream,
          kmdDescription: item.kmdDescription,
          migrationApproach: '',
          link: ''
        });
      }
    }
  });

  // Sort by Workstream, then KMD Description
  dispositions.sort((a, b) => {
    if (a.workstream !== b.workstream) {
      return a.workstream.localeCompare(b.workstream);
    }
    return a.kmdDescription.localeCompare(b.kmdDescription);
  });

  console.log(`Created ${dispositions.length} unique KMD dispositions`);
  return dispositions;
}

async function buildMerged(uploadBuffer) {
  const uploadWb = xlsx.read(uploadBuffer, { type: 'buffer' });
  const staticWb = xlsx.readFile(staticWorkbookPath);

  // Use "Recommended BPML" sheet from uploaded file
  const targetUploadSheetName = uploadWb.SheetNames.find(
    (n) => n.trim() === 'Recommended BPML'
  );
  if (!targetUploadSheetName) {
    throw new Error('Sheet "Recommended BPML" not found in uploaded file');
  }
  const uploadSheet = uploadWb.Sheets[targetUploadSheetName];
  const targetStaticSheetName =
    staticWb.SheetNames.find((n) => n.trim() === 'KMD Repository - 2023') ||
    staticWb.SheetNames[0];
  const staticSheet = staticWb.Sheets[targetStaticSheetName];
  if (!uploadSheet || !staticSheet) {
    throw new Error('Missing sheets in workbook');
  }

  const uploadRows = xlsx.utils.sheet_to_json(uploadSheet, {
    header: 1,
    defval: '',
  });
  const staticRows = xlsx.utils.sheet_to_json(staticSheet, {
    header: 1,
    defval: '',
  });

  if (uploadRows.length === 0 || staticRows.length === 0) {
    throw new Error('Workbooks are empty');
  }

  const uploadHeader = uploadRows[0];
  const staticHeader = staticRows[0];

  // Find Level 1, Level 2, Level 3 column indices in uploaded file by header name
  const uploadL1Idx = findColumnIndex(uploadHeader, ['level 1', 'business/enterprise area']);
  const uploadL2Idx = findColumnIndex(uploadHeader, ['level 2', 'scenario/process group']);
  const uploadL3Idx = findColumnIndex(uploadHeader, ['level 3 - process', 'level 3']);

  if (uploadL1Idx === -1 || uploadL2Idx === -1 || uploadL3Idx === -1) {
    throw new Error(
      `Could not find Level 1, Level 2, or Level 3 columns in Recommended BPML sheet. ` +
      `Found indices: L1=${uploadL1Idx}, L2=${uploadL2Idx}, L3=${uploadL3Idx}`
    );
  }

  console.log(`Found BPML columns - L1: ${uploadL1Idx} (${uploadHeader[uploadL1Idx]}), L2: ${uploadL2Idx} (${uploadHeader[uploadL2Idx]}), L3: ${uploadL3Idx} (${uploadHeader[uploadL3Idx]})`);

  // Get all KMD column indices from static/knowledge base file by header names
  const staticColIdx = getKmdColumnIndices(staticHeader);

  // Use found indices for L1, L2, L3 or defaults
  const sL1Idx = staticColIdx.level1 !== -1 ? staticColIdx.level1 : 0;
  const sL2Idx = staticColIdx.level2 !== -1 ? staticColIdx.level2 : 1;
  const sL3Idx = staticColIdx.level3 !== -1 ? staticColIdx.level3 : 2;

  // Get column headers for output using found indices
  const staticSimplificationHeader = staticColIdx.simplificationItem >= 0 ? staticHeader[staticColIdx.simplificationItem] : 'Simplification Item';
  const staticWorkstreamHeader = staticColIdx.workstream >= 0 ? staticHeader[staticColIdx.workstream] : 'Workstream';
  const staticKmdDescHeader = staticColIdx.kmdDescription >= 0 ? staticHeader[staticColIdx.kmdDescription] : 'KMD Description';
  const staticMigrationHeader = staticColIdx.migrationApproach >= 0 ? staticHeader[staticColIdx.migrationApproach] : 'Migration Approach';
  const staticLinkHeader = staticColIdx.link >= 0 ? staticHeader[staticColIdx.link] : 'Link';

  // Store arrays of rows to handle multiple KMD matches for same L1|L2|L3
  const staticMap = new Map();
  staticRows.slice(1).forEach((row) => {
    const key = normalizeKey(row, sL1Idx, sL2Idx, sL3Idx);
    if (!staticMap.has(key)) {
      staticMap.set(key, []);
    }
    staticMap.get(key).push(row);
  });

  const output = [];
  const headerRow = [
    uploadHeader[uploadL1Idx] || 'Level 1',
    uploadHeader[uploadL2Idx] || 'Level 2',
    uploadHeader[uploadL3Idx] || 'Level 3',
    staticSimplificationHeader,
    staticWorkstreamHeader,
    staticKmdDescHeader,
    staticMigrationHeader,
    staticLinkHeader,
  ];
  output.push(headerRow);

  // Track unique output rows to avoid duplicates
  const seenOutputRows = new Set();

  uploadRows.slice(1).forEach((uRow, index) => {
    // Clean brackets from L1, L2, L3 columns before matching
    const l1Val = cleanBrackets(uRow[uploadL1Idx]);
    const l2Val = cleanBrackets(uRow[uploadL2Idx]);
    const l3Val = cleanBrackets(uRow[uploadL3Idx]);

    // Debug logging for first 3 rows
    if (index < 3) {
      console.log(`\n--- Row ${index + 1} ---`);
      console.log('Original:', { l1: uRow[uploadL1Idx], l2: uRow[uploadL2Idx], l3: uRow[uploadL3Idx] });
      console.log('Cleaned:', { l1: l1Val, l2: l2Val, l3: l3Val });
    }

    // Create a temporary row array for normalizeKey
    const tempRow = [];
    tempRow[0] = l1Val;
    tempRow[1] = l2Val;
    tempRow[2] = l3Val;

    const key = normalizeKey(tempRow, 0, 1, 2);

    if (index < 3) {
      console.log('Match key:', key);
    }

    const matchingRows = staticMap.get(key);

    // Only include rows that have KB matches
    // Create separate row for each unique KB match
    if (matchingRows && matchingRows.length > 0) {
      matchingRows.forEach((sRow) => {
        const simplificationItem = staticColIdx.simplificationItem >= 0 ? (sRow[staticColIdx.simplificationItem] || '').toString().trim() : '';
        const workstream = staticColIdx.workstream >= 0 ? (sRow[staticColIdx.workstream] || '').toString().trim() : '';
        const kmdDescription = staticColIdx.kmdDescription >= 0 ? (sRow[staticColIdx.kmdDescription] || '').toString().trim() : '';
        const migrationApproach = staticColIdx.migrationApproach >= 0 ? (sRow[staticColIdx.migrationApproach] || '').toString().trim() : '';
        const link = staticColIdx.link >= 0 ? (sRow[staticColIdx.link] || '').toString().trim() : '';

        // Create unique key for this output row to avoid duplicates
        const outputKey = `${l1Val}|${l2Val}|${l3Val}|${simplificationItem}|${workstream}|${kmdDescription}|${migrationApproach}|${link}`;

        if (!seenOutputRows.has(outputKey)) {
          seenOutputRows.add(outputKey);
          output.push([
            l1Val,
            l2Val,
            l3Val,
            simplificationItem,
            workstream,
            kmdDescription,
            migrationApproach,
            link,
          ]);
        }
      });
    }
    // No else - rows without matches are not included
  });

  // Use ExcelJS for better formatting (bold headers, clickable links, table format)
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Merged');

  // Add header row (headerRow already defined above)
  worksheet.addRow(headerRow);

  // Style header row - bold only (no colors)
  const wsHeaderRow = worksheet.getRow(1);
  wsHeaderRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Add data rows
  const linkColIndex = headerRow.length; // Link is the last column (1-based will be headerRow.length)
  output.slice(1).forEach((row) => {
    const dataRow = worksheet.addRow(row);

    // Make link clickable (last column)
    const linkCell = dataRow.getCell(linkColIndex);
    const linkValue = row[linkColIndex - 1]; // 0-based index
    if (linkValue && linkValue.toString().trim()) {
      const url = linkValue.toString().trim();
      const fullUrl = url.startsWith('http') ? url : `https://${url}`;
      // Set hyperlink using ExcelJS formula approach
      linkCell.value = {
        formula: `HYPERLINK("${fullUrl}","${url.replace(/"/g, '""')}")`,
        result: url
      };
      linkCell.font = { underline: true, color: { argb: 'FF0563C1' } };
    }

    // Add borders to all cells
    dataRow.eachCell((cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  // Auto-fit column widths
  worksheet.columns.forEach((column, index) => {
    let maxLength = headerRow[index] ? headerRow[index].toString().length : 10;
    output.forEach((row) => {
      const cellValue = row[index] ? row[index].toString() : '';
      maxLength = Math.max(maxLength, Math.min(cellValue.length, 50)); // Cap at 50
    });
    column.width = maxLength + 2;
  });

  // Add table formatting
  if (output.length > 1) {
    const lastCol = String.fromCharCode(64 + headerRow.length); // Convert to letter (A, B, C...)
    const tableRef = `A1:${lastCol}${output.length}`;

    worksheet.addTable({
      name: 'MergedTable',
      ref: 'A1',
      headerRow: true,
      totalsRow: false,
      style: {
        theme: 'TableStyleLight1',
        showRowStripes: false,
      },
      columns: headerRow.map(h => ({ name: h, filterButton: true })),
      rows: output.slice(1),
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();

  return { buffer, rows: output };
}

function ensureStaticWorkbook() {
  if (!fs.existsSync(staticDir)) {
    fs.mkdirSync(staticDir, { recursive: true });
  }
  if (!fs.existsSync(staticWorkbookPath)) {
    console.warn(
      `Static workbook not found at ${staticWorkbookPath}. Place your knowledgebase file there.`,
    );
  }
}

ensureStaticWorkbook();

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Fiori Releases API - Proxies SAP Fiori Library API with caching
app.get('/api/fiori-releases', async (req, res) => {
  try {
    const now = Date.now();

    // Check if cached data is still fresh (less than 5 minutes old)
    if (fioriReleasesCache.data && (now - fioriReleasesCache.timestamp) < CACHE_DURATION) {
      console.log('✓ Returning cached Fiori releases (age: ' + Math.round((now - fioriReleasesCache.timestamp) / 1000) + 's)');
      return res.json(fioriReleasesCache.data);
    }

    // Cache is empty or stale, fetch fresh data from SAP
    console.log('→ Fetching fresh Fiori releases from SAP API...');
    const response = await fetch(
      'https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata/Releases?$format=json'
    );

    if (!response.ok) {
      throw new Error(`SAP API returned status ${response.status}`);
    }

    const data = await response.json();

    // Extract the releases array from the OData response
    const allReleases = data.d?.results || [];

    // Filter by releaseType: 'SOP' or 'SC' only
    const filteredReleases = allReleases.filter(release =>
      release.releaseType === 'SOP' || release.releaseType === 'SC'
    );

    if (filteredReleases.length === 0) {
      console.warn('⚠ SAP API returned no releases matching releaseType SOP or SC');
    } else {
      console.log(`✓ Fetched ${allReleases.length} total releases, filtered to ${filteredReleases.length} (SOP/SC only)`);
    }

    // Update cache with filtered data
    fioriReleasesCache = {
      data: filteredReleases,
      timestamp: now
    };

    res.json(filteredReleases);
  } catch (error) {
    console.error('✗ Error fetching Fiori releases:', error.message);

    // If fetch failed but we have stale cached data, return it as fallback
    if (fioriReleasesCache.data) {
      console.log('→ Returning stale cached data as fallback');
      return res.json(fioriReleasesCache.data);
    }

    // No cache available, return error
    res.status(500).json({
      error: 'Failed to fetch Fiori releases',
      details: error.message
    });
  }
});

app.get('/api/static-file', (req, res) => {
  if (!fs.existsSync(staticWorkbookPath)) {
    res.status(404).json({ error: 'Static workbook not found' });
    return;
  }
  res.download(staticWorkbookPath, 'static-knowledgebase.xlsx');
});

app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }

  try {
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheets = workbook.SheetNames.map((name) => {
      const sheet = workbook.Sheets[name];
      const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });
      return {
        sheet: name,
        rows: rows.length,
        columns: rows[0] ? rows[0].length : 0,
      };
    });

    res.json({
      filename: req.file.originalname,
      size: req.file.size,
      sheets,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to read workbook' });
  }
});

app.post('/api/merge', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }
  if (!fs.existsSync(staticWorkbookPath)) {
    res.status(500).json({ error: 'Static workbook missing' });
    return;
  }

  try {
    const { buffer } = await buildMerged(req.file.buffer);

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="KMD Disposition - BPML.xlsx"',
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Merge failed' });
  }
});

app.post('/api/merge-preview', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }
  if (!fs.existsSync(staticWorkbookPath)) {
    res.status(500).json({ error: 'Static workbook missing' });
    return;
  }

  try {
    const { rows } = await buildMerged(req.file.buffer);
    const [header, ...data] = rows;
    res.json({
      header,
      rows: data,
      totalRows: data.length,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Merge preview failed' });
  }
});

app.post('/api/merge-download', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }
  if (!fs.existsSync(staticWorkbookPath)) {
    res.status(500).json({ error: 'Static workbook missing' });
    return;
  }

  try {
    const { buffer } = await buildMerged(req.file.buffer);

    res.setHeader(
      'Content-Disposition',
      'attachment; filename="KMD Disposition - BPML.xlsx"',
    );
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Merge failed' });
  }
});

app.post('/api/extract-docx', upload.single('file'), (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }

  // Check if it's a .docx file
  if (!req.file.originalname.toLowerCase().endsWith('.docx')) {
    res.status(400).json({ error: 'Only .docx files are allowed' });
    return;
  }

  try {
    const extractedData = extractTablesFromDocx(req.file.buffer);
    console.log(`\n=== Extracted ${extractedData.length} rows from .docx ===`);

    const allItems = mapWithKMD(extractedData);
    console.log(`\n=== Mapped ${allItems.length} rows with KMD ===`);

    const simplificationItems = filterSimplificationItems(allItems);
    console.log(`\n=== Filtered ${simplificationItems.length} simplification items ===`);

    const pivotData = createPivotTable(simplificationItems);
    console.log(`\n=== Created pivot table ===`);

    const kmdDispositions = createKMDDispositions(simplificationItems);
    console.log(`\n=== Created KMD dispositions ===\n`);

    res.json({
      filename: req.file.originalname,
      rowsExtracted: extractedData.length,
      rowsMapped: allItems.length,
      rowsFiltered: simplificationItems.length,
      sheets: [
        {
          name: 'All Items',
          data: allItems,
        },
        {
          name: 'Simplification Items',
          data: simplificationItems,
        },
        {
          name: 'Summary',
          data: pivotData, // This is already in array format for preview
        },
        {
          name: 'KMD Dispositions',
          data: kmdDispositions,
        }
      ]
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to extract data from .docx file', details: err.message });
  }
});

app.post('/api/download-docx', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }

  if (!req.file.originalname.toLowerCase().endsWith('.docx')) {
    res.status(400).json({ error: 'Only .docx files are allowed' });
    return;
  }

  try {
    console.log('Starting download-docx processing...');
    const extractedData = extractTablesFromDocx(req.file.buffer);
    const allItems = mapWithKMD(extractedData);
    const simplificationItems = filterSimplificationItems(allItems);
    const pivotData = createPivotTable(simplificationItems);
    const kmdDispositions = createKMDDispositions(simplificationItems);

    // Use ExcelJS for better formatting
    console.log('Creating workbook with ExcelJS...');
    const workbook = new ExcelJS.Workbook();

    // Sheet 1: All Items
    // Helper to make SAP Note cells hyperlinks
    const makeSapNoteHyperlinks = (worksheet, sapNoteColIndex) => {
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header
        const cell = row.getCell(sapNoteColIndex);
        const value = cell.value ? cell.value.toString() : '';
        if (value.startsWith('http://') || value.startsWith('https://')) {
          cell.value = { text: value, hyperlink: value };
          cell.font = { color: { argb: 'FF0066CC' }, underline: true };
        }
      });
    };

    const ws1 = workbook.addWorksheet('All Items');
    ws1.columns = [
      { header: 'Title', key: 'title', width: 50 },
      { header: 'Effort Ranking', key: 'effortRanking', width: 20 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'SAP Note', key: 'sapNote', width: 40 },
      { header: 'Workstream', key: 'workstream', width: 20 },
      { header: 'KMD Description', key: 'kmdDescription', width: 30 },
    ];
    ws1.addRows(allItems);
    ws1.getRow(1).font = { bold: true };
    makeSapNoteHyperlinks(ws1, 4); // SAP Note is column 4

    // Sheet 2: Simplification Items
    const ws2 = workbook.addWorksheet('Simplification Items');
    ws2.columns = [
      { header: 'Title', key: 'title', width: 50 },
      { header: 'Effort Ranking', key: 'effortRanking', width: 20 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'SAP Note', key: 'sapNote', width: 40 },
      { header: 'Workstream', key: 'workstream', width: 20 },
      { header: 'KMD Description', key: 'kmdDescription', width: 30 },
    ];
    ws2.addRows(simplificationItems);
    ws2.getRow(1).font = { bold: true };
    makeSapNoteHyperlinks(ws2, 4); // SAP Note is column 4

    // Sheet 3: Summary with pivot table
    const ws3 = workbook.addWorksheet('Summary');

    // Add pivot table data
    pivotData.forEach((row, idx) => {
      ws3.addRow(row);
      if (idx === 0 || idx === pivotData.length - 1) {
        ws3.getRow(idx + 1).font = { bold: true };
      }
    });

    ws3.columns.forEach(column => { column.width = 15; });
    ws3.getColumn(1).width = 25;

    // Sheet 4: KMD Dispositions
    const ws4 = workbook.addWorksheet('KMD Dispositions');
    ws4.columns = [
      { header: 'Workstream', key: 'workstream', width: 25 },
      { header: 'KMD Description', key: 'kmdDescription', width: 40 },
      { header: 'Migration Approach', key: 'migrationApproach', width: 25 },
      { header: 'Links', key: 'link', width: 50 },
    ];
    ws4.addRows(kmdDispositions);
    ws4.getRow(1).font = { bold: true };

    // Make links clickable
    kmdDispositions.forEach((item, idx) => {
      if (item.link && item.link.startsWith('http')) {
        const cell = ws4.getCell(`D${idx + 2}`); // Column D, skip header
        cell.value = { text: item.link, hyperlink: item.link };
        cell.font = { color: { argb: 'FF0563C1' }, underline: true };
      }
    });

    // Write to temporary file
    const tempPath = path.join(__dirname, `temp-${Date.now()}.xlsx`);
    console.log('Writing to temp file:', tempPath);

    let buffer;
    try {
      await workbook.xlsx.writeFile(tempPath);
      console.log('File written successfully');

      // Add chart using Python script
      console.log('Adding bar chart using Python...');
      const pythonScript = path.join(__dirname, 'add_chart.py');
      try {
        execSync(`python3 "${pythonScript}" "${tempPath}"`, { encoding: 'utf-8' });
        console.log('Chart added successfully!');
      } catch (pythonError) {
        console.error('Python script error:', pythonError.message);
        throw new Error('Failed to add chart: ' + pythonError.message);
      }

      // Read the file with chart
      buffer = fs.readFileSync(tempPath);
    } finally {
      // Clean up temp file - this always runs, even if there's an error
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
        console.log('Temp file cleaned up:', tempPath);
      }
    }

    res.setHeader('Content-Disposition', 'attachment; filename="KMD Disposition - Readiness Check.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
    console.log('Download successful with embedded chart!');
  } catch (err) {
    console.error('Download error:', err);
    console.error('Error stack:', err.stack);
    res.status(500).json({ error: 'Download failed', details: err.message });
  }
});

// Fetch Fiori apps for transaction codes from AI Tool Excel file
app.post('/api/fetch-fiori-apps', upload.single('file'), async (req, res) => {
  if (!req.file) {
    res.status(400).json({ error: 'No file uploaded' });
    return;
  }

  const { releaseId } = req.body;
  if (!releaseId) {
    res.status(400).json({ error: 'releaseId is required' });
    return;
  }

  try {
    console.log(`\n=== Fetching Fiori apps for releaseId: ${releaseId} ===`);

    // Read Excel file with ExcelJS to preserve formatting
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    // Find the "Recommended BPML" sheet
    const worksheet = workbook.worksheets.find(ws => ws.name.trim() === 'Recommended BPML');

    if (!worksheet) {
      throw new Error('Sheet "Recommended BPML" not found in Excel file.');
    }

    console.log(`Using BPML sheet: "${worksheet.name}"`);

    // Get header row (row 1)
    const headerRow = worksheet.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber] = cell.value ? cell.value.toString().trim() : '';
    });

    if (headers.length === 0) {
      throw new Error('BPML sheet is empty');
    }

    // Find column index for "S/4HANA Transaction Code" (1-based in ExcelJS)
    let tcodeColumnIndex = -1;
    for (let i = 1; i < headers.length; i++) {
      const h = headers[i]?.toLowerCase() || '';
      if (h.includes('s/4hana transaction code') || h.includes('s4 tcode') || h.includes('s/4 tcode')) {
        tcodeColumnIndex = i;
        break;
      }
    }

    if (tcodeColumnIndex === -1) {
      throw new Error('Column "S/4HANA Transaction Code" not found in BPML sheet');
    }

    console.log(`Found transaction code column at index: ${tcodeColumnIndex}`);

    // Extract and deduplicate transaction codes
    const tcodesSet = new Set();
    const rowCount = worksheet.rowCount;
    for (let rowNum = 2; rowNum <= rowCount; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const tcode = row.getCell(tcodeColumnIndex).value;
      if (tcode && tcode.toString().trim() !== '') {
        tcodesSet.add(tcode.toString().trim().toUpperCase());
      }
    }

    const uniqueTcodes = Array.from(tcodesSet);
    console.log(`Extracted ${uniqueTcodes.length} unique transaction codes from ${rowCount - 1} rows`);
    console.log('Unique tcodes:', uniqueTcodes);

    // Fetch Fiori apps for each unique tcode in parallel
    const fioriApiUrl = 'https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/services/SingleApp.xsodata';

    const fetchPromises = uniqueTcodes.map(async (tcode) => {
      const url = `${fioriApiUrl}/AppListResult(sWhereClause='(1=1) and ("releaseId" = ''${releaseId}'') and ("TRANSACTION_MATCH" = ''${tcode}'')',INPLANGUAGE='None',sUUID='')/Results?$top=100&$select=fioriId,AppName,ApplicationType,TRANSACTION_MATCH,GTMAppDescription,UITechnology&$format=json`;

      try {
        console.log(`→ Fetching Fiori apps for tcode: ${tcode}`);
        const response = await fetch(url);

        if (!response.ok) {
          console.log(`✗ Failed to fetch for ${tcode}: ${response.status}`);
          return { tcode, apps: [] };
        }

        const data = await response.json();
        const results = data.d?.results || [];

        // Filter out "SAP GUI" application types
        const filteredApps = results.filter(app =>
          app.ApplicationType && app.ApplicationType !== 'SAP GUI'
        );

        console.log(`✓ Found ${filteredApps.length} Fiori apps for ${tcode} (filtered out SAP GUI)`);
        return { tcode, apps: filteredApps };
      } catch (error) {
        console.log(`✗ Error fetching for ${tcode}:`, error.message);
        return { tcode, apps: [] };
      }
    });

    // Wait for all API calls to complete
    const results = await Promise.allSettled(fetchPromises);

    // Build map: tcode -> array of Fiori apps
    const tcodeToAppsMap = new Map();
    results.forEach((result) => {
      if (result.status === 'fulfilled') {
        const { tcode, apps } = result.value;
        tcodeToAppsMap.set(tcode, apps);
      }
    });

    console.log(`\n=== Built map with ${tcodeToAppsMap.size} tcodes ===`);

    // Add 4 new columns at the end
    const lastColIndex = worksheet.columnCount;
    const startCol = lastColIndex + 1; // Start after the last existing column
    const newColumns = ['Fiori ID', 'Fiori App Name', 'Application Type', 'App Details Link'];
    const fioriAppLibraryBaseUrl = 'https://fioriappslibrary.hana.ondemand.com/sap/fix/externalViewer/?appId=';

    // Get reference cell from existing data to copy formatting (use first data column)
    const referenceHeaderCell = headerRow.getCell(1);
    const referenceDataCell = worksheet.getRow(2).getCell(1);

    // Set header values with formatting copied from existing headers
    for (let i = 0; i < newColumns.length; i++) {
      const cell = headerRow.getCell(startCol + i);
      cell.value = newColumns[i];
      // Copy font from reference header but ensure bold
      if (referenceHeaderCell.font) {
        cell.font = { ...referenceHeaderCell.font, bold: true };
      } else {
        cell.font = { bold: true };
      }
      // Copy other formatting from reference header
      if (referenceHeaderCell.fill) cell.fill = referenceHeaderCell.fill;
      if (referenceHeaderCell.border) cell.border = referenceHeaderCell.border;
      if (referenceHeaderCell.alignment) cell.alignment = referenceHeaderCell.alignment;
    }

    // Fill in Fiori data for each row with formatting
    for (let rowNum = 2; rowNum <= rowCount; rowNum++) {
      const row = worksheet.getRow(rowNum);
      const tcode = row.getCell(tcodeColumnIndex).value;
      const tcodeKey = tcode && tcode.toString().trim() !== ''
        ? tcode.toString().trim().toUpperCase()
        : null;

      let fioriIds = '';
      let fioriAppNames = '';
      let applicationTypes = '';
      let appDetailsLinks = '';

      if (tcodeKey && tcodeToAppsMap.has(tcodeKey)) {
        const apps = tcodeToAppsMap.get(tcodeKey);

        if (apps.length > 0) {
          fioriIds = apps.map(app => app.fioriId || '').join(', ');
          fioriAppNames = apps.map(app => app.AppName || '').join(', ');
          // Get unique application types only
          const uniqueAppTypes = [...new Set(apps.map(app => app.ApplicationType || '').filter(t => t))];
          applicationTypes = uniqueAppTypes.join(', ');
          appDetailsLinks = apps.map(app => app.fioriId ? `${fioriAppLibraryBaseUrl}${app.fioriId}` : '').join(', ');
        }
      }

      // Get reference cell from same row to copy formatting
      const refCell = row.getCell(1);

      // Set values and copy formatting for each new cell
      const fioriData = [fioriIds, fioriAppNames, applicationTypes];
      for (let i = 0; i < fioriData.length; i++) {
        const cell = row.getCell(startCol + i);
        cell.value = fioriData[i];
        // Copy formatting from reference cell
        if (refCell.font) cell.font = refCell.font;
        if (refCell.fill) cell.fill = refCell.fill;
        if (refCell.border) cell.border = refCell.border;
        if (refCell.alignment) cell.alignment = refCell.alignment;
      }

      // Handle App Details Link - show all app links
      const linkCell = row.getCell(startCol + 3);
      if (tcodeKey && tcodeToAppsMap.has(tcodeKey)) {
        const apps = tcodeToAppsMap.get(tcodeKey);
        if (apps.length > 0) {
          // Create all app links
          const allLinks = apps
            .filter(app => app.fioriId)
            .map(app => `${fioriAppLibraryBaseUrl}${app.fioriId}`);

          if (allLinks.length === 1) {
            // Single link - make it a clickable hyperlink
            linkCell.value = {
              text: allLinks[0],
              hyperlink: allLinks[0]
            };
            linkCell.font = { color: { argb: 'FF0066CC' }, underline: true };
          } else if (allLinks.length > 1) {
            // Multiple links - show as newline-separated text (Excel limitation: only one hyperlink per cell)
            linkCell.value = allLinks.join('\n');
            linkCell.font = { color: { argb: 'FF0066CC' } };
            linkCell.alignment = { ...refCell.alignment, wrapText: true };
          } else {
            linkCell.value = '';
          }
        } else {
          linkCell.value = '';
        }
      } else {
        linkCell.value = '';
      }
      if (refCell.fill) linkCell.fill = refCell.fill;
      if (refCell.border) linkCell.border = refCell.border;
    }

    console.log(`Updated ${rowCount - 1} rows with Fiori data (added at end, columns ${startCol}-${startCol + 3})`);

    // Write workbook to buffer
    const buffer = await workbook.xlsx.writeBuffer();

    // Send response with updated Excel and processed tcodes
    res.setHeader('Content-Disposition', 'attachment; filename="BPML - FIORI Mapping.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('X-Processed-Tcodes', JSON.stringify(uniqueTcodes));
    res.send(buffer);

    console.log('✓ Successfully sent updated Excel with Fiori fields\n');
  } catch (err) {
    console.error('Fiori fetch error:', err);
    res.status(500).json({ error: 'Failed to fetch Fiori apps', details: err.message });
  }
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(400).json({ error: err.message || 'Upload failed' });
});

app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});

