import { useMemo, useRef, useState, useEffect } from 'react';
import pwcLogo from './assets/pwc-logo.png';
import * as XLSX from 'xlsx-js-style';
import { PieChart, Pie, Cell, Tooltip, Legend, ResponsiveContainer, BarChart, Bar, XAxis, YAxis, CartesianGrid } from 'recharts';

function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.min(sizes.length - 1, Math.floor(Math.log(bytes) / Math.log(k)));
  return `${(bytes / k ** i).toFixed(2)} ${sizes[i]}`;
}

export default function App() {
  const API_BASE = import.meta.env.VITE_API_BASE ?? '';
  const [cards, setCards] = useState([
    { title: 'L1-L5 BPML File', subtitle: 'Supported: Excel (xlsx / xlsm)', fileType: 'xlsx', file: null },
    { title: 'Readiness Check File', subtitle: 'Supported: Word (.docx)', fileType: 'docx', file: null },
  ]);
  const [activeCard, setActiveCard] = useState(0);
  const [outputs, setOutputs] = useState([]); // {cardIdx, filename, size, sheets, header, rows, progress, done}
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);
  const [mergingIdx, setMergingIdx] = useState(null);
  const [view, setView] = useState('main'); // 'main' | 'preview'
  const [previewIdx, setPreviewIdx] = useState(null);
  const [page, setPage] = useState(1);
  const pageSize = 25;
  const fileInputRef = useRef(null);
  const [fioriVersion, setFioriVersion] = useState('');
  const [fioriReleases, setFioriReleases] = useState([]);
  const [loadingReleases, setLoadingReleases] = useState(false);
  const [fioriLoading, setFioriLoading] = useState(false);
  const [fioriProgress, setFioriProgress] = useState(0);
  const [fioriOutput, setFioriOutput] = useState(null); // {header, rows, done, filename, size}
  const [fioriBlob, setFioriBlob] = useState(null);
  const [bpmlChartData, setBpmlChartData] = useState([]); // Pie chart data for L1 vs L4
  const [ocmChartData, setOcmChartData] = useState([]); // Bar chart data for L1 vs OCM Valid
  const [changeCategoryChartData, setChangeCategoryChartData] = useState([]); // Stacked bar chart data for Change Category by L1
  const [stackedCategoryL1Keys, setStackedCategoryL1Keys] = useState([]); // L1 area names used as stack keys
  const [ocmL5ChartData, setOcmL5ChartData] = useState([]); // Bar chart data for L1 vs unique L5 count (filtered by OCM Valid: SCH_R, SCH_N, TRN_N, TRN_C, FRC)
  const [bpmlAnalysisData, setBpmlAnalysisData] = useState(null); // Full data from Recommended BPML sheet for analysis page
  const [showBpmlAnalysis, setShowBpmlAnalysis] = useState(false); // Show BPML Analysis full page
  const [testingScopeData, setTestingScopeData] = useState(null); // Data for Testing Scope Dashboard (L1, L4, L5)
  const [showTestingScope, setShowTestingScope] = useState(false); // Show Testing Scope Dashboard full page

  const hasFile = useMemo(() => cards.some((c) => c.file), [cards]);
  const uploadedCount = useMemo(() => cards.filter((c) => c.file).length, [cards]);
  const consideredCount = useMemo(() => Math.min(uploadedCount, 2), [uploadedCount]);
  const overallProgress = useMemo(() => {
    const hasFileProgress = consideredCount > 0;
    const hasFioriProgress = fioriLoading || fioriOutput;

    if (!hasFileProgress && !hasFioriProgress) return 0;

    const fileSum = outputs.reduce((acc, o) => acc + (o.progress || 0), 0);
    const fileAvg = consideredCount > 0 ? fileSum / consideredCount : 0;

    if (hasFioriProgress && hasFileProgress) {
      // Both file processing and Fiori mapping - average all three
      const totalCount = consideredCount + 1;
      return Math.min(100, Math.round((fileSum + fioriProgress) / totalCount));
    } else if (hasFioriProgress) {
      // Only Fiori mapping
      return Math.min(100, Math.round(fioriProgress));
    } else {
      // Only file processing
      return Math.min(100, Math.round(fileAvg));
    }
  }, [outputs, consideredCount, fioriLoading, fioriOutput, fioriProgress]);
  const analysisDone = useMemo(() => {
    const needed = consideredCount;
    const completed = outputs.filter((o) => o.done).length;
    return needed > 0 && completed === needed;
  }, [outputs, consideredCount]);
  const canSubmit = useMemo(() => hasFile && !loading, [hasFile, loading]);

  // Fetch Fiori releases on component mount
  useEffect(() => {
    const fetchReleases = async () => {
      setLoadingReleases(true);
      try {
        const response = await fetch(`${API_BASE}/api/fiori-releases`);

        if (!response.ok) {
          throw new Error('Failed to fetch Fiori releases');
        }

        const releases = await response.json();

        // Sort by releaseRank in descending order (latest version first)
        const sorted = releases.sort((a, b) => b.releaseRank - a.releaseRank);
        setFioriReleases(sorted);
      } catch (err) {
        console.error('Failed to fetch Fiori releases:', err);
        setError('Failed to load S/4 versions');
      } finally {
        setLoadingReleases(false);
      }
    };
    fetchReleases();
  }, [API_BASE]);

  const handleFilePick = () => {
    fileInputRef.current?.click();
  };

  // Parse BPML file and extract chart data
  const parseBpmlForCharts = async (file) => {
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(data), { type: 'array' });

      // Find "Recommended BPML" sheet
      const sheetName = workbook.SheetNames.find(name => name.trim() === 'Recommended BPML');
      if (!sheetName) {
        console.log('Recommended BPML sheet not found for chart');
        setBpmlAnalysisData(null);
        return;
      }

      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (rows.length < 2) {
        setBpmlAnalysisData(null);
        return;
      }

      const headers = rows[0];

      // Store full data for BPML Analysis page
      setBpmlAnalysisData({
        header: headers,
        rows: rows.slice(1),
        filename: file.name,
        size: file.size,
      });

      // Find column indices
      const l1Idx = headers.findIndex(h => h && h.toString().toLowerCase().includes('level 1'));
      const l4Idx = headers.findIndex(h => h && h.toString().toLowerCase().includes('level 4'));
      const l5Idx = headers.findIndex(h => h && (h.toString().toLowerCase().includes('level 5') || h.toString().toLowerCase().includes('task (l5)')));
      const ocmIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('ocm valid'));

      // Store data for Testing Scope Dashboard (L1, L4, L5) - same filter logic as bar chart
      const changeCategoryIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('change category'));

      if (l1Idx !== -1 && l4Idx !== -1 && l5Idx !== -1 && ocmIdx !== -1 && changeCategoryIdx !== -1) {
        const l1ColName = headers[l1Idx];
        const l4ColName = headers[l4Idx];
        const l5ColName = headers[l5Idx];

        // Same filter criteria as the bar chart
        const validChangeCategories = ['SCH_R', 'SCH_N', 'TRN_N', 'TRN_C', 'FRC'];

        // Build L1 -> Map of L5 -> L4 (for unique L5 per L1 with their L4)
        const l1ToL5L4Map = new Map();

        rows.slice(1).forEach(row => {
          const l1 = (row[l1Idx] || '').toString().trim();
          const l4 = (row[l4Idx] || '').toString().trim();
          const l5 = (row[l5Idx] || '').toString().trim();
          const ocmValid = (row[ocmIdx] ?? '').toString().trim().toLowerCase();
          const changeCategory = (row[changeCategoryIdx] ?? '').toString().trim().toUpperCase();

          // Apply same filter as bar chart: OCM Valid = 'Change' AND valid Change Category
          if (l1 && l5 && ocmValid === 'change' && validChangeCategories.includes(changeCategory)) {
            if (!l1ToL5L4Map.has(l1)) {
              l1ToL5L4Map.set(l1, new Map());
            }
            // Store L5 -> L4 mapping (first L4 found for each unique L5)
            if (!l1ToL5L4Map.get(l1).has(l5)) {
              l1ToL5L4Map.get(l1).set(l5, l4);
            }
          }
        });

        // Extract unique L1 values that have data
        const uniqueL1Values = Array.from(l1ToL5L4Map.keys()).sort();

        // Convert to flat rows format for the component
        const testingRows = [];
        l1ToL5L4Map.forEach((l5Map, l1) => {
          l5Map.forEach((l4, l5) => {
            testingRows.push({ l1, l4, l5 });
          });
        });

        setTestingScopeData({
          uniqueL1Values,
          rows: testingRows,
          l1ColName,
          l4ColName,
          l5ColName,
          filename: file.name,
          size: file.size,
        });
      } else {
        setTestingScopeData(null);
      }

      // Helper to clean Level 1 names (remove "- Industry Edge")
      const cleanL1Name = (name) => name.replace(/\s*-\s*Industry Edge/gi, '').trim();

      // Chart 1: Level 1 vs unique Level 4 counts
      if (l1Idx !== -1 && l4Idx !== -1) {
        const l1ToL4Map = new Map();
        rows.slice(1).forEach(row => {
          const l1 = cleanL1Name((row[l1Idx] || '').toString().trim());
          const l4 = (row[l4Idx] || '').toString().trim();
          if (l1 && l4) {
            if (!l1ToL4Map.has(l1)) {
              l1ToL4Map.set(l1, new Set());
            }
            l1ToL4Map.get(l1).add(l4);
          }
        });

        const chartData = Array.from(l1ToL4Map.entries())
          .map(([l1, l4Set]) => ({
            name: l1,
            value: l4Set.size,
          }))
          .filter(item => item.name && item.value > 0);
        setBpmlChartData(chartData);
        console.log('Pie chart data:', chartData);
      }

      // Shared abbreviation logic for Chart 2 and Chart 4
      // Extract leading capital letters from name (e.g., "OTC Order to Cash" → "OTC")
      const getAbbreviation = (name) => {
        if (!name) return '';
        const match = name.match(/^([A-Z]+)/);
        if (match && match[1]) {
          return match[1];
        }
        // Fallback: first 3 characters uppercase
        return name.substring(0, 3).toUpperCase();
      };

      // Chart 2: Level 1 vs OCM Valid (New, Change, Existing)
      if (l1Idx !== -1 && ocmIdx !== -1) {
        const l1OcmMap = new Map(); // l1 -> { New, Change, Existing }

        rows.slice(1).forEach((row) => {
          const l1 = cleanL1Name((row[l1Idx] || '').toString().trim());
          const ocmValid = ((row[ocmIdx] ?? '').toString().trim()).toLowerCase();

          if (!l1 || !ocmValid) return;

          let key = null;
          if (ocmValid === 'new') key = 'New';
          else if (ocmValid === 'change') key = 'Change';
          else if (ocmValid === 'existing') key = 'Existing';

          if (!key) return;

          if (!l1OcmMap.has(l1)) l1OcmMap.set(l1, { New: 0, Change: 0, Existing: 0 });
          l1OcmMap.get(l1)[key]++;
        });

        // Track used abbreviations to avoid duplicates
        const usedAbbrs = new Map();
        const ocmData = Array.from(l1OcmMap.entries())
          .filter(([l1, counts]) => l1 && (counts.New + counts.Change + counts.Existing) > 0)
          .map(([l1, counts], index) => {
            let abbr = getAbbreviation(l1, index);
            if (usedAbbrs.has(abbr)) {
              const num = usedAbbrs.get(abbr) + 1;
              usedAbbrs.set(abbr, num);
              abbr = `${abbr}${num}`;
            } else {
              usedAbbrs.set(abbr, 1);
            }
            return {
              fullName: l1,
              shortCode: abbr,
              New: counts.New,
              Change: counts.Change,
              Existing: counts.Existing,
            };
          });
        setOcmChartData(ocmData);
        console.log('Bar chart data (OCM):', ocmData);
      }

      // Chart 3: Change Category counts by Level 1 (stacked bar chart)
      if (changeCategoryIdx !== -1 && l1Idx !== -1) {
        const categoryDescriptions = {
          'NCG': 'No Change',
          'SCH_R': 'Screen Change – Revised',
          'SCH_N': 'Screen Change – New',
          'TRN_N': 'New Transaction – New',
          'TRN_C': 'New Transaction – Changed',
          'PRC': 'Process Change',
          'FRC': 'Fiori Change',
        };

        // nested map: category -> { l1Name -> count }
        const categoryL1Map = new Map();
        const l1NamesSet = new Set();
        rows.slice(1).forEach(row => {
          const category = (row[changeCategoryIdx] || '').toString().trim().toUpperCase();
          const l1 = cleanL1Name((row[l1Idx] || '').toString().trim());
          if (category && l1) {
            l1NamesSet.add(l1);
            if (!categoryL1Map.has(category)) categoryL1Map.set(category, new Map());
            const l1Map = categoryL1Map.get(category);
            l1Map.set(l1, (l1Map.get(l1) || 0) + 1);
          }
        });

        const l1Keys = Array.from(l1NamesSet).sort();
        const changeCategoryData = Array.from(categoryL1Map.entries())
          .filter(([code]) => code)
          .map(([code, l1Map]) => {
            const entry = { name: categoryDescriptions[code] || code, code };
            l1Keys.forEach(l1 => { entry[l1] = l1Map.get(l1) || 0; });
            return entry;
          });
        setChangeCategoryChartData(changeCategoryData);
        setStackedCategoryL1Keys(l1Keys);
        console.log('Stacked bar chart data (Change Category):', changeCategoryData);
      }

      // Chart 4: Level 1 vs unique Level 5 count (filtered by OCM Valid = 'Change' AND specific Change Category values)
      console.log('Chart 4 - Column indices:', { l1Idx, l5Idx, ocmIdx, changeCategoryIdx });
      console.log('Chart 4 - Level 5 column name:', l5Idx !== -1 ? headers[l5Idx] : 'NOT FOUND');
      console.log('Chart 4 - OCM Valid column name:', ocmIdx !== -1 ? headers[ocmIdx] : 'NOT FOUND');
      console.log('Chart 4 - Change Category column name:', changeCategoryIdx !== -1 ? headers[changeCategoryIdx] : 'NOT FOUND');

      // Sample OCM Valid and Change Category values for debugging
      if (ocmIdx !== -1 && changeCategoryIdx !== -1) {
        const sampleOcmValues = new Set();
        const sampleCategoryValues = new Set();
        rows.slice(1, 11).forEach(row => {
          const ocmVal = (row[ocmIdx] ?? '').toString().trim();
          const catVal = (row[changeCategoryIdx] ?? '').toString().trim();
          if (ocmVal) sampleOcmValues.add(ocmVal);
          if (catVal) sampleCategoryValues.add(catVal);
        });
        console.log('Chart 4 - Sample OCM Valid values (first 10 rows):', Array.from(sampleOcmValues));
        console.log('Chart 4 - Sample Change Category values (first 10 rows):', Array.from(sampleCategoryValues));
      }

      if (l1Idx !== -1 && l5Idx !== -1 && ocmIdx !== -1 && changeCategoryIdx !== -1) {
        // Specific Change Category values to include
        const validChangeCategories = ['SCH_R', 'SCH_N', 'TRN_N', 'TRN_C', 'FRC'];
        const l1ToL5Map = new Map();
        let matchedRows = 0;
        let ocmChangeCount = 0;
        let validCategoryCount = 0;
        let bothConditionsCount = 0;

        rows.slice(1).forEach(row => {
          const l1 = cleanL1Name((row[l1Idx] || '').toString().trim());
          const l5 = (row[l5Idx] || '').toString().trim();

          // Check OCM Valid = 'Change'
          const rawOcm = row[ocmIdx];
          const ocmValid = (rawOcm ?? '').toString().trim().toLowerCase();

          // Check Change Category
          const rawCategory = row[changeCategoryIdx];
          const changeCategory = (rawCategory ?? '').toString().trim().toUpperCase();

          // Count rows matching each condition
          if (ocmValid === 'change') ocmChangeCount++;
          if (validChangeCategories.includes(changeCategory)) validCategoryCount++;
          if (ocmValid === 'change' && validChangeCategories.includes(changeCategory)) bothConditionsCount++;

          // Only include rows where OCM Valid = 'Change' AND Change Category matches one of the specified values
          if (l1 && l5 && ocmValid === 'change' && validChangeCategories.includes(changeCategory)) {
            matchedRows++;
            if (!l1ToL5Map.has(l1)) {
              l1ToL5Map.set(l1, new Set());
            }
            l1ToL5Map.get(l1).add(l5);
          }
        });

        console.log('Chart 4 - Rows where OCM Valid = "change":', ocmChangeCount);
        console.log('Chart 4 - Rows with valid Change Category:', validCategoryCount);
        console.log('Chart 4 - Rows matching BOTH conditions:', bothConditionsCount);
        console.log('Chart 4 - Matched rows with L1 and L5:', matchedRows);
        console.log('Chart 4 - L1 to L5 map size:', l1ToL5Map.size);

        // Generate chart data with abbreviations (same logic as Chart 2)
        const usedAbbrs = new Map();
        const ocmL5Data = Array.from(l1ToL5Map.entries())
          .filter(([l1, l5Set]) => l1 && l5Set.size > 0)
          .map(([l1, l5Set], index) => {
            let abbr = getAbbreviation(l1, index);
            // Handle duplicate abbreviations
            if (usedAbbrs.has(abbr)) {
              const num = usedAbbrs.get(abbr) + 1;
              usedAbbrs.set(abbr, num);
              abbr = `${abbr}${num}`;
            } else {
              usedAbbrs.set(abbr, 1);
            }
            return {
              fullName: l1,
              shortCode: abbr,
              count: l5Set.size,
            };
          });
        setOcmL5ChartData(ocmL5Data);
        console.log('Bar chart data (L1 vs unique L5 where OCM Valid = Change):', ocmL5Data);
      } else {
        console.warn('Chart 4 - Cannot generate: Missing required columns', {
          hasL1: l1Idx !== -1,
          hasL5: l5Idx !== -1,
          hasOcmValid: ocmIdx !== -1,
          hasChangeCategory: changeCategoryIdx !== -1
        });
        setOcmL5ChartData([]);
      }
    } catch (err) {
      console.error('Error parsing BPML for charts:', err);
    }
  };

  const handleFileChange = (e) => {
    const f = e.target.files?.[0] || null;
    if (!f) return;

    // Validate file type matches the card
    const card = cards[activeCard];
    const isValidFile =
      (card.fileType === 'docx' && f.name.toLowerCase().endsWith('.docx')) ||
      (card.fileType === 'xlsx' && (f.name.toLowerCase().endsWith('.xlsx') || f.name.toLowerCase().endsWith('.xlsm')));

    if (!isValidFile) {
      setError(`Please select a valid ${card.fileType === 'docx' ? '.docx' : '.xlsx or .xlsm'} file`);
      e.target.value = ''; // Reset input
      return;
    }

    setCards((prev) =>
      prev.map((c, idx) => (idx === activeCard ? { ...c, file: f } : c)),
    );
    setOutputs((prev) => prev.filter((o) => o.cardIdx !== activeCard));
    setError('');
    setPreviewIdx(null);
    setView('main');

    // If L1-L5 BPML file (card index 0), parse for charts
    if (activeCard === 0) {
      parseBpmlForCharts(f);
    }

    // Reset input to allow re-uploading same file
    e.target.value = '';
  };

  const handleDeleteFile = (cardIdx, e) => {
    e.stopPropagation(); // Prevent triggering file picker
    setCards((prev) =>
      prev.map((c, idx) => (idx === cardIdx ? { ...c, file: null } : c)),
    );
    setOutputs((prev) => prev.filter((o) => o.cardIdx !== cardIdx));
    setError('');
    setPreviewIdx(null);
    setView('main');

    // Clear chart data and Fiori state if BPML file is deleted
    if (cardIdx === 0) {
      setBpmlChartData([]);
      setOcmChartData([]);
      setChangeCategoryChartData([]);
      setOcmL5ChartData([]);
      setBpmlAnalysisData(null);
      setShowBpmlAnalysis(false);
      setFioriOutput(null);
      setFioriBlob(null);
      setFioriProgress(0);
    }
  };

  const upsertOutput = (cardIdx, patch) => {
    setOutputs((prev) => {
      const existing = prev.find((o) => o.cardIdx === cardIdx) || {
        cardIdx,
        header: [],
        rows: [],
        progress: 0,
        done: false,
      };
      const next = { ...existing, ...patch };
      const others = prev.filter((o) => o.cardIdx !== cardIdx);
      return [...others, next].sort((a, b) => a.cardIdx - b.cardIdx);
    });
  };

  const processFile = async (file, cardIdx, selectedFioriVersion) => {
    if (!file) return;
    const data = new FormData();
    data.append('file', file);
    if (selectedFioriVersion) {
      data.append('fioriVersion', selectedFioriVersion);
    }

    const isDocx = file.name.toLowerCase().endsWith('.docx');

    upsertOutput(cardIdx, { progress: 10, done: false });

    if (isDocx) {
      // Process .docx file
      const extractRes = await fetch(`${API_BASE}/api/extract-docx`, { method: 'POST', body: data });
      const extractPayload = await extractRes.json();
      if (!extractRes.ok) throw new Error(extractPayload.error || 'Extraction failed');

      upsertOutput(cardIdx, {
        filename: file.name,
        size: file.size,
        sheets: [
          { sheet: 'All Items', rows: extractPayload.rowsMapped, columns: 6 },
          { sheet: 'Simplification Items', rows: extractPayload.rowsFiltered, columns: 6 },
          { sheet: 'Summary', rows: extractPayload.sheets[2].data.length - 1, columns: extractPayload.sheets[2].data[0].length },
          { sheet: 'KMD Dispositions', rows: extractPayload.sheets[3].data.length, columns: 4 }
        ],
        progress: 50,
      });

      // Convert mapped data to table format for preview - prepare all sheets
      const detailHeader = ['Title', 'Effort Ranking', 'Category', 'SAP Note', 'Workstream', 'KMD Description'];

      const allItemsRows = extractPayload.sheets[0].data.map(item => [
        item.title,
        item.effortRanking,
        item.category,
        item.sapNote,
        item.workstream,
        item.kmdDescription
      ]);

      const simplificationRows = extractPayload.sheets[1].data.map(item => [
        item.title,
        item.effortRanking,
        item.category,
        item.sapNote,
        item.workstream,
        item.kmdDescription
      ]);

      // Summary sheet is already in array format (pivot table)
      const summaryData = extractPayload.sheets[2].data;
      const summaryHeader = summaryData[0]; // First row is header
      const summaryRows = summaryData.slice(1); // Rest are data rows

      // KMD Dispositions sheet
      const kmdDispositionsHeader = ['Workstream', 'KMD Description', 'Migration Approach', 'Links'];
      const kmdDispositionsRows = extractPayload.sheets[3].data.map(item => [
        item.workstream,
        item.kmdDescription,
        item.migrationApproach,
        item.link
      ]);

      upsertOutput(cardIdx, {
        multiSheet: true,
        sheets: [
          { name: 'All Items', header: detailHeader, rows: allItemsRows },
          { name: 'Simplification Items', header: detailHeader, rows: simplificationRows },
          { name: 'Summary', header: summaryHeader, rows: summaryRows },
          { name: 'KMD Dispositions', header: kmdDispositionsHeader, rows: kmdDispositionsRows }
        ],
        progress: 100,
        done: true,
      });
    } else {
      // Process .xlsx file (existing logic)
      const uploadRes = await fetch(`${API_BASE}/api/upload`, { method: 'POST', body: data });
      const uploadPayload = await uploadRes.json();
      if (!uploadRes.ok) throw new Error(uploadPayload.error || 'Upload failed');
      upsertOutput(cardIdx, {
        filename: file.name,
        size: file.size,
        sheets: uploadPayload.sheets || [],
        progress: 50,
      });

      const previewRes = await fetch(`${API_BASE}/api/merge-preview`, {
        method: 'POST',
        body: data,
      });
      const previewPayload = await previewRes.json();
      if (!previewRes.ok) throw new Error(previewPayload.error || 'Merge failed');

      upsertOutput(cardIdx, {
        header: previewPayload.header || [],
        rows: previewPayload.rows || [],
        progress: 100,
        done: true,
      });
    }
  };

  const handleSubmit = async () => {
    if (!hasFile) return;
    setLoading(true);
    setError('');
    setOutputs([]);
    setFioriOutput(null);
    setFioriBlob(null);
    setFioriProgress(0);
    setPreviewIdx(null);
    setView('main');
    try {
      const filesToProcess = cards
        .map((c, idx) => ({ file: c.file, cardIdx: idx }))
        .filter((f) => f.file)
        .slice(0, 2);

      // Process all files in parallel
      await Promise.all(
        filesToProcess.map(item => processFile(item.file, item.cardIdx, fioriVersion))
      );
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async (cardIdx) => {
    const targetFile = cards[cardIdx]?.file;
    if (!targetFile) return;
    setMergingIdx(cardIdx);
    setError('');
    try {
      const data = new FormData();
      data.append('file', targetFile);

      // Determine which endpoint to use based on file type
      const isDocx = targetFile.name.toLowerCase().endsWith('.docx');
      const endpoint = isDocx ? '/api/download-docx' : '/api/merge-download';
      const filename = isDocx ? 'KMD Disposition - Readiness Check.xlsx' : 'KMD Disposition - BPML.xlsx';

      const res = await fetch(`${API_BASE}${endpoint}`, {
        method: 'POST',
        body: data,
      });
      if (!res.ok) {
        const payload = await res.json().catch(() => ({}));
        throw new Error(payload.error || 'Download failed');
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError(err.message);
    } finally {
      setMergingIdx(null);
    }
  };

  const handleFetchFioriFields = async () => {
    // Get BPML file (cardIdx 0)
    const bpmlFile = cards[0]?.file;
    if (!bpmlFile) {
      setError('Please upload L1-L5 BPML file first');
      return;
    }

    if (!fioriVersion) {
      setError('Please select S/4 Target Version first');
      return;
    }

    // Find the releaseId for the selected fioriVersion
    const selectedRelease = fioriReleases.find(r => r.releaseName === fioriVersion);
    if (!selectedRelease) {
      setError('Invalid S/4 version selected');
      return;
    }

    setError('');
    setFioriOutput(null);
    setFioriBlob(null);
    setFioriProgress(10);
    setFioriLoading(true);

    try {
      setFioriProgress(30);

      const data = new FormData();
      data.append('file', bpmlFile);
      data.append('releaseId', selectedRelease.releaseId);

      const res = await fetch(`${API_BASE}/api/fetch-fiori-apps`, {
        method: 'POST',
        body: data,
      });

      setFioriProgress(70);

      if (!res.ok) {
        const payload = await res.json().catch(() => ({}));
        throw new Error(payload.error || 'Failed to fetch Fiori apps');
      }

      // Store the blob for download
      const blob = await res.blob();
      setFioriBlob(blob);

      setFioriProgress(85);

      // Parse the blob to extract preview data
      const arrayBuffer = await blob.arrayBuffer();
      const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });

      // Find the "Recommended BPML" sheet
      const bpmlSheetName = workbook.SheetNames.find(name =>
        name.trim() === 'Recommended BPML'
      );

      if (!bpmlSheetName) {
        throw new Error('Sheet "Recommended BPML" not found in result file.');
      }

      const sheet = workbook.Sheets[bpmlSheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (rows.length === 0) {
        throw new Error('Result sheet is empty');
      }

      const headerRow = rows[0];

      // Find column indices for transaction code and Fiori fields
      const tcodeColIdx = headerRow.findIndex(col =>
        col && (col.toString().toLowerCase().includes('s/4hana transaction code') ||
                col.toString().toLowerCase().includes('s4 tcode') ||
                col.toString().toLowerCase().includes('s/4 tcode'))
      );
      const fioriIdColIdx = headerRow.findIndex(col =>
        col && col.toString().toLowerCase() === 'fiori id'
      );
      const fioriAppNameColIdx = headerRow.findIndex(col =>
        col && col.toString().toLowerCase() === 'fiori app name'
      );
      const appTypeColIdx = headerRow.findIndex(col =>
        col && col.toString().toLowerCase() === 'application type'
      );
      const fioriLinkColIdx = headerRow.findIndex(col =>
        col && (col.toString().toLowerCase().includes('app details link') ||
                col.toString().toLowerCase().includes('fiori app link') ||
                col.toString().toLowerCase().includes('app link'))
      );

      // Helper to get cell address (e.g., "A1", "B2")
      const getCellAddress = (rowIdx, colIdx) => {
        const colLetter = XLSX.utils.encode_col(colIdx);
        return `${colLetter}${rowIdx + 1}`;
      };

      // Helper to extract all hyperlink URLs from cell (handles multiple links)
      const extractLinks = (rowIdx, colIdx) => {
        if (colIdx < 0) return '';
        const cellAddr = getCellAddress(rowIdx, colIdx);
        const cell = sheet[cellAddr];
        if (!cell) return '';

        // Get cell value
        const val = (cell.v || '').toString();

        // Check if cell has a hyperlink (XLSX stores in cell.l.Target)
        if (cell.l && cell.l.Target) {
          return cell.l.Target;
        }

        // Check if value contains URLs (could be newline-separated)
        if (val.includes('http://') || val.includes('https://')) {
          return val; // Return as-is, will be split in the render
        }

        return val;
      };

      // Build preview data with unique tcodes and their Fiori mappings
      const previewHeader = ['S/4 Transaction Code', 'Fiori ID', 'Fiori App Name', 'Application Type', 'Fiori App Link'];
      const seenTcodes = new Set();
      const previewRows = [];

      rows.slice(1).forEach((row, idx) => {
        const rowIdx = idx + 1; // Account for header row
        const tcode = tcodeColIdx >= 0 ? (row[tcodeColIdx] || '').toString().trim().toUpperCase() : '';
        if (tcode && !seenTcodes.has(tcode)) {
          seenTcodes.add(tcode);
          const linkValue = extractLinks(rowIdx, fioriLinkColIdx);
          previewRows.push([
            tcode,
            fioriIdColIdx >= 0 ? (row[fioriIdColIdx] || '') : '',
            fioriAppNameColIdx >= 0 ? (row[fioriAppNameColIdx] || '') : '',
            appTypeColIdx >= 0 ? (row[appTypeColIdx] || '') : '',
            linkValue
          ]);
        }
      });

      setFioriOutput({
        filename: 'BPML - FIORI Mapping.xlsx',
        size: blob.size,
        header: previewHeader,
        rows: previewRows,
        progress: 100,
        done: true,
      });
      setFioriProgress(100);
    } catch (err) {
      setError(err.message);
      setFioriOutput(null);
      setFioriBlob(null);
      setFioriProgress(0);
    } finally {
      setFioriLoading(false);
    }
  };

  const handleFioriDownload = () => {
    if (!fioriBlob) return;
    const url = window.URL.createObjectURL(fioriBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'BPML - FIORI Mapping.xlsx';
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);
  };

  return (
    <div style={styles.shell}>
      <div style={styles.hero}>
        <div>
          <h1 style={styles.title}>SmartBPML-KMD Disposition</h1>
        </div>
        <img
          src={pwcLogo}
          alt="PwC"
          style={styles.logo}
        />
      </div>

      <section style={styles.card}>
        <div style={styles.sectionHeader}>
          <h3 style={styles.sectionTitle}>Input Files (2)</h3>
          <button style={styles.collapse}>Collapse</button>
        </div>
        <div style={styles.cardGrid}>
          {cards.map((c, idx) => {
            const ready = !!c.file;
            const isActive = activeCard === idx;
            return (
              <div
                key={c.title}
                style={{
                  ...styles.dropCard,
                  border: isActive ? '2px solid #fb923c' : styles.dropCard.border,
                }}
                onClick={() => {
                  setActiveCard(idx);
                  handleFilePick();
                }}
              >
                <div style={styles.cardTopRow}>
                  <span style={styles.cardTitle}>{c.title}</span>
                  <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                    <span style={ready ? styles.ready : styles.pending}>{ready ? 'Ready' : 'Empty'}</span>
                    {ready && (
                      <button
                        style={styles.deleteFileBtn}
                        onClick={(e) => handleDeleteFile(idx, e)}
                        aria-label="Delete file"
                        title="Delete file"
                      >
                        ✕
                      </button>
                    )}
                  </div>
                </div>
                <p style={styles.cardHelp}>{c.subtitle}</p>
                <div style={styles.dropZone}>
                  <div style={styles.dropIcon}>⧉</div>
                  <div>
                    <div style={styles.dropText}>Click or drop file</div>
                    <div style={styles.dropSub}>
                      {ready ? 'File attached' : c.subtitle}
                    </div>
                    {ready && (
                      <div style={styles.fileName}>
                        {c.file.name} · {formatBytes(c.file.size)}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </section>

      <section style={styles.card}>
        <div style={styles.sectionHeader}>
          <h3 style={styles.sectionTitle}>S/4 Target Version (Fiori)</h3>
        </div>
        <div style={styles.dropdownWrapper}>
          <select
            style={styles.dropdown}
            value={fioriVersion}
            onChange={(e) => setFioriVersion(e.target.value)}
            disabled={loadingReleases}
          >
            <option value="">Select version...</option>
            {fioriReleases.map((release) => (
              <option key={release.releaseId} value={release.releaseName}>
                {release.releaseName}
              </option>
            ))}
          </select>
          {loadingReleases && <span style={styles.loadingText}>Loading versions...</span>}
        </div>
      </section>

      <section style={styles.outputsCard}>
        <div style={styles.outputsHeader}>
          <div>
            <h3 style={styles.sectionTitle}>Outputs</h3>
            <p style={styles.outputsSub}>Track pipeline status and inspect categorized results.</p>
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
            <a
              style={{
                ...styles.actionButton,
                textDecoration: 'none',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: 6,
                minWidth: 150,
              }}
              href={`${API_BASE}/api/static-file`}
              aria-label="Download knowledgebase workbook"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                width="18"
                height="18"
                fill="currentColor"
                viewBox="0 0 16 16"
              >
                <path d="M.5 9.5a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V10a.5.5 0 0 1 1 0v2.5A2 2 0 0 1 14 14.5H2A2 2 0 0 1 0 12.5V10a.5.5 0 0 1 .5-.5" />
                <path d="M7.646 11.854a.5.5 0 0 0 .708 0l3-3a.5.5 0 0 0-.708-.708L8.5 10.293V1.5a.5.5 0 0 0-1 0v8.793L5.354 8.146a.5.5 0 1 0-.708.708z" />
              </svg>
              <span>Knowledgebase</span>
            </a>
            <button
              style={{
                ...styles.runBtn,
                ...(canSubmit ? {} : styles.disabledButton),
              }}
              onClick={handleSubmit}
              disabled={!canSubmit}
            >
              {loading ? 'Uploading…' : 'Run Analysis'}
            </button>
            <button
              style={{
                ...styles.fioriBtn,
                ...(!cards[0]?.file || !fioriVersion || fioriLoading ? styles.disabledButton : {}),
              }}
              onClick={handleFetchFioriFields}
              disabled={!cards[0]?.file || !fioriVersion || fioriLoading}
            >
              {fioriLoading ? 'Processing…' : 'Map Fiori Apps'}
            </button>
          </div>
        </div>

        <div style={styles.progressWrap}>
          <div style={styles.progressMeta}>
            <div style={styles.progressLabel}>Overall Progress</div>
            <div style={styles.progressPct}>{overallProgress}%</div>
          </div>
          <div style={styles.progressBarOuter}>
            <div style={{ ...styles.progressBarInner, width: `${overallProgress}%` }} />
          </div>
        </div>

        {(fioriLoading || fioriOutput) && (
          <div style={styles.progressWrap}>
            <div style={styles.progressMeta}>
              <div style={styles.progressLabel}>Fiori Mapping Progress</div>
              <div style={styles.progressPct}>{fioriProgress}%</div>
            </div>
            <div style={styles.progressBarOuter}>
              <div style={{ ...styles.progressBarInner, width: `${fioriProgress}%` }} />
            </div>
          </div>
        )}

        <div style={styles.summaryRow}>
          {/* BPML Analysis Tile - Only enabled when L1-L5 BPML file is uploaded */}
          <SummaryCard
            key="bpml-analysis"
            title="BPML Analysis"
            output={bpmlAnalysisData ? {
              filename: bpmlAnalysisData.filename,
              size: bpmlAnalysisData.size,
              progress: 100,
              done: true,
              sheets: [{ name: 'Recommended BPML', rows: bpmlAnalysisData.rows.length }],
            } : null}
            onPreview={() => {
              if (bpmlAnalysisData) {
                setShowBpmlAnalysis(true);
              }
            }}
          />
          <SummaryCard
            key="testing-scope"
            title="Testing Scope Dashboard"
            output={testingScopeData ? {
              filename: testingScopeData.filename,
              size: testingScopeData.size,
              progress: 100,
              done: true,
              sheets: [{ name: 'Testing Scope', rows: testingScopeData.rows.length }],
            } : null}
            onPreview={() => {
              if (testingScopeData) {
                setShowTestingScope(true);
              }
            }}
          />
          {[0, 1].map((idx) => {
            const card = cards[idx];
            const output = outputs.find((o) => o.cardIdx === idx);
            return (
              <SummaryCard
                key={card?.title || idx}
                title={idx === 0 ? 'KMD Disposition - BPML' : 'KMD Disposition - Readiness Check'}
                output={output}
                onPreview={() => {
                  if (output?.done) {
                    setPreviewIdx(idx);
                    setPage(1);
                    setView('preview');
                  }
                }}
              />
            );
          })}
          <SummaryCard
            key="fiori-output"
            title="BPML - FIORI Mapping"
            output={fioriOutput}
            onPreview={() => {
              if (fioriOutput?.done) {
                setPreviewIdx('fiori');
                setPage(1);
                setView('preview');
              }
            }}
          />
        </div>

        <div style={styles.actionsRow}>
        </div>

        {error && <div style={styles.error}>{error}</div>}

      </section>

      <input
        ref={fileInputRef}
        type="file"
        accept={
          cards[activeCard]?.fileType === 'docx'
            ? '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            : '.xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
        style={{ display: 'none' }}
        onChange={handleFileChange}
        key={activeCard}
      />
      {view === 'preview' && previewIdx !== null && (
        <PreviewPage
          output={previewIdx === 'fiori' ? fioriOutput : outputs.find((o) => o.cardIdx === previewIdx)}
          onClose={() => setView('main')}
          onDownload={previewIdx === 'fiori' ? handleFioriDownload : () => handleDownload(previewIdx)}
          downloading={previewIdx === 'fiori' ? false : mergingIdx === previewIdx}
          page={page}
          setPage={setPage}
          pageSize={pageSize}
          title={previewIdx === 'fiori' ? 'Fiori Mapping Output' : previewIdx === 0 ? 'KMD Disposition - BPML' : 'KMD Disposition - Readiness Check'}
        />
      )}

      {/* BPML Analysis Full Page */}
      {showBpmlAnalysis && bpmlAnalysisData && (
        <BpmlAnalysisPage
          data={bpmlAnalysisData}
          onClose={() => setShowBpmlAnalysis(false)}
          bpmlChartData={bpmlChartData}
          ocmChartData={ocmChartData}
          changeCategoryChartData={changeCategoryChartData}
          stackedCategoryL1Keys={stackedCategoryL1Keys}
          ocmL5ChartData={ocmL5ChartData}
        />
      )}

      {/* Testing Scope Dashboard Full Page */}
      {showTestingScope && testingScopeData && (
        <TestingScopePage
          data={testingScopeData}
          onClose={() => setShowTestingScope(false)}
          ocmL5ChartData={ocmL5ChartData}
        />
      )}
    </div>
  );
}

// Chart colors
const CHART_COLORS = ['#FD5108', '#FF7F3E', '#FFB347', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'];

// Extract leading capital letters from name as abbreviation
function getBusinessAbbreviation(name) {
  if (!name) return '';
  // Extract leading uppercase letters (e.g., "OTC Order to Cash" → "OTC")
  const match = name.match(/^([A-Z]+)/);
  if (match && match[1]) {
    return match[1];
  }
  // Fallback: first 3 characters uppercase
  return name.substring(0, 3).toUpperCase();
}

// Custom bar shape that enforces a minimum visible height for stacked segments
const MIN_BAR_HEIGHT = 6;
function StackedBarWithMinHeight(props) {
  const { x, y, width, height, fill } = props;
  if (!height) return null;
  const displayHeight = Math.max(height, MIN_BAR_HEIGHT);
  const displayY = y + height - displayHeight;
  return <rect x={x} y={displayY} width={width} height={displayHeight} fill={fill} />;
}

function SummaryCard({ title, label, output, onPreview }) {
  const progress = output?.progress ?? 0;
  const status = output
    ? output.done
      ? 'Completed'
      : 'In progress'
    : 'Pending';
  const badgeStyle = output?.done ? styles.badgeSuccess : styles.badgePending;
  return (
    <div style={styles.summaryCard}>
      <div style={styles.summaryTop}>
        <span style={styles.summaryTitle}>{title}</span>
        <span style={badgeStyle}>{status}</span>
      </div>
      <div style={styles.progressBarOuter}>
        <div style={{ ...styles.progressBarInner, width: `${progress}%` }} />
      </div>
      <div style={styles.summaryFooter}>
        <span style={styles.progressLabel}>{progress}%</span>
        <span style={styles.summaryValue}>{label}</span>
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginTop: 10 }}>
        {output && (
          <div style={styles.analysisBox}>
            <div style={styles.analysisRow}>
              <span style={styles.tableLabel}>File</span>
              <span style={styles.tableValue}>
                {output.filename || '—'} ({output.size ? formatBytes(output.size) : '—'})
              </span>
            </div>
            {output.sheets && output.sheets.length > 1 && (
              <div style={styles.analysisRow}>
                <span style={styles.tableLabel}>Sheets</span>
                <span style={styles.tableValue}>
                  {output.sheets.length} sheets
                </span>
              </div>
            )}
            <div style={styles.analysisRow}>
              <span style={styles.tableLabel}>Status</span>
              <span style={styles.tableValue}>
                {output.done ? 'Ready for download' : 'Processing...'}
              </span>
            </div>
          </div>
        )}
        <button
          style={{
            ...styles.actionButton,
            ...(output && progress >= 100 ? {} : styles.disabledButton),
          }}
          onClick={onPreview}
          disabled={!output || progress < 100}
        >
          View & Download
        </button>
      </div>
    </div>
  );
}

function PreviewPage({
  output,
  onClose,
  onDownload,
  downloading,
  page,
  setPage,
  pageSize,
  title = 'Merged Output',
}) {
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);

  // Check if this is multi-sheet output
  const isMultiSheet = output?.multiSheet && output?.sheets;

  // Get header and rows based on whether it's multi-sheet or single-sheet
  // For multi-sheet, each sheet can have its own header
  const header = isMultiSheet
    ? (output.sheets[activeSheetIdx]?.header || [])
    : (output?.header || []);

  const rows = isMultiSheet
    ? (output.sheets[activeSheetIdx]?.rows || [])
    : (output?.rows || []);

  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  const start = (page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);
  const rowCount = rows.length;
  const colCount = header.length;

  return (
    <div style={styles.previewOverlay}>
      <div style={styles.previewCard}>
        <div style={styles.previewHeader}>
          <div>
            <div style={styles.sectionTitle}>{title}</div>
            <div style={styles.outputsSub}>Columns and rows from data</div>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button
              style={{
                ...styles.actionButton,
                ...(downloading ? styles.disabledButton : {}),
              }}
              onClick={onDownload}
              disabled={downloading}
            >
              {downloading ? 'Downloading…' : 'Download Excel'}
            </button>
            <button style={styles.closeButton} onClick={onClose} aria-label="Close">
              ✕
            </button>
          </div>
        </div>

        {isMultiSheet && (
          <div style={styles.sheetTabs}>
            {output.sheets.map((sheet, idx) => (
              <button
                key={idx}
                style={{
                  ...styles.sheetTab,
                  ...(activeSheetIdx === idx ? styles.sheetTabActive : {}),
                }}
                onClick={() => {
                  setActiveSheetIdx(idx);
                  setPage(1);
                }}
              >
                {sheet.name} ({sheet.rows.length})
              </button>
            ))}
          </div>
        )}

        <div style={styles.tableWrap}>
          <table style={styles.table}>
            <thead>
              <tr>
                {header.map((h, idx) => (
                  <th key={idx} style={styles.th}>
                    {h || '-'}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {pageRows.map((row, idx) => {
                // Check if this is Fiori output - get Fiori IDs from column index 1
                const isFioriOutput = title === 'Fiori Mapping Output';
                const fioriIdStr = isFioriOutput ? (row[1] || '').toString().trim() : '';
                // Split Fiori IDs in case there are multiple (comma or newline separated)
                const fioriIds = fioriIdStr.split(/[\n,]/).map(s => s.trim()).filter(s => s);

                return (
                  <tr key={idx}>
                    {row.map((cell, cIdx) => {
                      const cellStr = (cell || '').toString();
                      // Check if cell contains URLs (could be multiple, newline or comma separated)
                      const urls = cellStr.split(/[\n,]/).map(s => s.trim()).filter(s =>
                        s.startsWith('http://') || s.startsWith('https://')
                      );

                      // Helper to extract SAP Note number from URL
                      const getSapNoteLabel = (url) => {
                        // Match SAP Note URLs like launchpad.support.sap.com/#/notes/123456 or me.sap.com/notes/123456
                        const sapNoteMatch = url.match(/\/notes\/(\d+)/i);
                        if (sapNoteMatch) {
                          return `SAP Note ${sapNoteMatch[1]}`;
                        }
                        return null;
                      };

                      // Get link label based on context
                      const getLinkLabel = (url, linkIdx) => {
                        const sapNoteLabel = getSapNoteLabel(url);
                        if (sapNoteLabel) return sapNoteLabel;
                        // For Fiori output, use corresponding Fiori ID by index
                        if (isFioriOutput && fioriIds.length > 0) {
                          return fioriIds[linkIdx] || fioriIds[0] || `App ${linkIdx + 1}`;
                        }
                        return `Open KMD${urls.length > 1 ? ` ${linkIdx + 1}` : ''}`;
                      };

                      return (
                        <td key={cIdx} style={styles.td}>
                          {urls.length > 0 ? (
                            <div style={styles.linkContainer}>
                              {urls.map((url, linkIdx) => (
                                <a
                                  key={linkIdx}
                                  href={url}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  style={styles.tableLink}
                                >
                                  {getLinkLabel(url, linkIdx)}
                                </a>
                              ))}
                            </div>
                          ) : (
                            cell
                          )}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
              {pageRows.length === 0 && (
                <tr>
                  <td style={styles.td} colSpan={header.length || 1}>
                    No data
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        <div style={styles.pagination}>
          <button
            style={styles.secondaryButton}
            onClick={() => setPage(Math.max(1, page - 1))}
            disabled={page === 1}
          >
            Prev
          </button>
          <span style={{ fontWeight: 700 }}>
            Page {page} of {totalPages}
          </span>
          <button
            style={styles.secondaryButton}
            onClick={() => setPage(Math.min(totalPages, page + 1))}
            disabled={page === totalPages}
          >
            Next
          </button>
        </div>
      </div>
    </div>
  );
}

function SearchableDropdown({ value, options, onChange, placeholder = 'Select...', multiSelect = false }) {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [hoveredIdx, setHoveredIdx] = useState(null);
  const wrapperRef = useRef(null);
  const inputRef = useRef(null);

  // For multi-select, value should be an array
  const selectedValues = multiSelect ? (Array.isArray(value) ? value : []) : null;

  // Filter options based on search term
  const filteredOptions = useMemo(() => {
    if (!searchTerm) return options;
    return options.filter(opt =>
      opt.toString().toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [options, searchTerm]);

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClickOutside = (event) => {
      if (wrapperRef.current && !wrapperRef.current.contains(event.target)) {
        setIsOpen(false);
        setSearchTerm('');
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleInputChange = (e) => {
    const val = e.target.value;
    setSearchTerm(val);
    if (!multiSelect) {
      onChange(val);
    }
    if (!isOpen) setIsOpen(true);
  };

  const handleSelect = (opt) => {
    if (multiSelect) {
      // Toggle selection in multi-select mode
      const newSelection = selectedValues.includes(opt)
        ? selectedValues.filter(v => v !== opt)
        : [...selectedValues, opt];
      onChange(newSelection);
      setSearchTerm('');
      // Keep dropdown open in multi-select mode
    } else {
      onChange(opt);
      setSearchTerm('');
      setIsOpen(false);
    }
  };

  const handleClear = (e) => {
    e.stopPropagation();
    onChange(multiSelect ? [] : '');
    setSearchTerm('');
    setIsOpen(false);
  };

  const handleRemoveTag = (optToRemove, e) => {
    e.stopPropagation();
    const newSelection = selectedValues.filter(v => v !== optToRemove);
    onChange(newSelection);
  };

  const handleInputFocus = () => {
    setIsOpen(true);
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Escape') {
      setIsOpen(false);
      setSearchTerm('');
    } else if (e.key === 'Enter' && filteredOptions.length > 0 && !multiSelect) {
      handleSelect(filteredOptions[0]);
    }
  };

  const isSelected = (opt) => {
    if (multiSelect) {
      return selectedValues.includes(opt);
    }
    return value === opt;
  };

  const getDisplayValue = () => {
    if (multiSelect) {
      return searchTerm;
    }
    return value;
  };

  const getPlaceholderText = () => {
    if (multiSelect && selectedValues.length > 0) {
      return `${selectedValues.length} selected`;
    }
    return placeholder;
  };

  return (
    <div ref={wrapperRef} style={styles.searchableDropdown}>
      <div
        style={{
          ...styles.dropdownInputWrapper,
          ...(isOpen ? styles.dropdownInputWrapperFocused : {}),
          ...(multiSelect && selectedValues.length > 0 ? { flexWrap: 'wrap', minHeight: '36px', height: 'auto' } : {}),
        }}
        onClick={() => inputRef.current?.focus()}
      >
        {multiSelect && selectedValues.length > 0 && (
          <div style={styles.selectedTagsContainer}>
            {selectedValues.map((val, idx) => (
              <span key={idx} style={styles.selectedTag}>
                <span style={styles.selectedTagText}>
                  {val.length > 20 ? val.substring(0, 20) + '...' : val}
                </span>
                <button
                  style={styles.selectedTagRemove}
                  onClick={(e) => handleRemoveTag(val, e)}
                  aria-label={`Remove ${val}`}
                >
                  ×
                </button>
              </span>
            ))}
          </div>
        )}
        <input
          ref={inputRef}
          type="text"
          style={{
            ...styles.dropdownInput,
            ...(multiSelect && selectedValues.length > 0 ? { flex: '1 1 100px', minWidth: '100px' } : {}),
          }}
          value={getDisplayValue()}
          onChange={handleInputChange}
          onFocus={handleInputFocus}
          onKeyDown={handleKeyDown}
          placeholder={getPlaceholderText()}
        />
        {((multiSelect && selectedValues.length > 0) || (!multiSelect && value)) && (
          <button style={styles.dropdownClearBtn} onClick={handleClear} aria-label="Clear">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
              <line x1="18" y1="6" x2="6" y2="18"></line>
              <line x1="6" y1="6" x2="18" y2="18"></line>
            </svg>
          </button>
        )}
        <button
          style={styles.dropdownToggleBtn}
          onClick={(e) => { e.stopPropagation(); setIsOpen(!isOpen); }}
          aria-label="Toggle dropdown"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            style={{ transform: isOpen ? 'rotate(180deg)' : 'rotate(0deg)', transition: 'transform 0.2s' }}
          >
            <polyline points="6 9 12 15 18 9"></polyline>
          </svg>
        </button>
      </div>
      {isOpen && (
        <div style={styles.dropdownMenu}>
          {multiSelect && (
            <div
              style={{
                ...styles.dropdownOption,
                ...((selectedValues.length === 0) ? styles.dropdownOptionSelected : {}),
                ...(hoveredIdx === -1 ? styles.dropdownOptionHovered : {}),
                fontWeight: '600',
              }}
              onClick={() => { onChange([]); setSearchTerm(''); }}
              onMouseEnter={() => setHoveredIdx(-1)}
              onMouseLeave={() => setHoveredIdx(null)}
            >
              <span style={styles.dropdownOptionAll}>Clear All</span>
            </div>
          )}
          {!multiSelect && (
            <div
              style={{
                ...styles.dropdownOption,
                ...(value === '' ? styles.dropdownOptionSelected : {}),
                ...(hoveredIdx === -1 ? styles.dropdownOptionHovered : {}),
              }}
              onClick={() => handleSelect('')}
              onMouseEnter={() => setHoveredIdx(-1)}
              onMouseLeave={() => setHoveredIdx(null)}
            >
              <span style={styles.dropdownOptionAll}>All</span>
            </div>
          )}
          {filteredOptions.length > 0 ? (
            <div style={styles.dropdownOptionsList}>
              {filteredOptions.slice(0, 100).map((opt, idx) => (
                <div
                  key={idx}
                  style={{
                    ...styles.dropdownOption,
                    ...(isSelected(opt) ? styles.dropdownOptionSelected : {}),
                    ...(hoveredIdx === idx ? styles.dropdownOptionHovered : {}),
                    ...(multiSelect ? { display: 'flex', alignItems: 'center', gap: '8px' } : {}),
                  }}
                  onClick={() => handleSelect(opt)}
                  onMouseEnter={() => setHoveredIdx(idx)}
                  onMouseLeave={() => setHoveredIdx(null)}
                >
                  {multiSelect && (
                    <input
                      type="checkbox"
                      checked={isSelected(opt)}
                      readOnly
                      style={{ pointerEvents: 'none', margin: 0 }}
                    />
                  )}
                  <span style={{ flex: 1 }}>
                    {opt.length > 60 ? opt.substring(0, 60) + '...' : opt}
                  </span>
                </div>
              ))}
              {filteredOptions.length > 100 && (
                <div style={styles.dropdownMoreText}>
                  +{filteredOptions.length - 100} more items...
                </div>
              )}
            </div>
          ) : (
            <div style={styles.dropdownNoResults}>No matches found</div>
          )}
        </div>
      )}
    </div>
  );
}

function BpmlAnalysisPage({
  data,
  onClose,
  bpmlChartData,
  ocmChartData,
  changeCategoryChartData,
  stackedCategoryL1Keys,
  ocmL5ChartData,
}) {
  const [page, setPage] = useState(1);
  const [filters, setFilters] = useState({});
  const [sortColumn, setSortColumn] = useState(null);
  const [sortDirection, setSortDirection] = useState('asc'); // 'asc' or 'desc'
  const pageSize = 25;

  const rawHeader = data?.header || [];
  const rawRows = data?.rows || [];

  // Find column indices to hide (Source Transaction Code)
  const hiddenColumnIndices = useMemo(() => {
    const hidden = new Set();
    rawHeader.forEach((col, idx) => {
      const colName = (col || '').toString().trim();
      if (colName === 'Source Transaction Code') {
        hidden.add(idx);
      }
    });
    return hidden;
  }, [rawHeader]);

  // Find Sequence No. column index for filtering rows
  const sequenceColumnIdx = useMemo(() => {
    return rawHeader.findIndex(col => {
      const colName = (col || '').toString().trim();
      return colName === 'Sequence No.';
    });
  }, [rawHeader]);

  // Filter out rows where Sequence Number is null/empty
  const rowsWithSequence = useMemo(() => {
    if (sequenceColumnIdx === -1) return rawRows;
    return rawRows.filter(row => {
      const seqValue = (row[sequenceColumnIdx] || '').toString().trim();
      return seqValue !== '' && seqValue.toLowerCase() !== 'null';
    });
  }, [rawRows, sequenceColumnIdx]);

  // Create header and rows without hidden columns
  const header = useMemo(() => {
    return rawHeader.filter((_, idx) => !hiddenColumnIndices.has(idx));
  }, [rawHeader, hiddenColumnIndices]);

  const allRows = useMemo(() => {
    return rowsWithSequence.map(row =>
      row.filter((_, idx) => !hiddenColumnIndices.has(idx))
    );
  }, [rowsWithSequence, hiddenColumnIndices]);

  // Define which columns should have filters
  const filterColumnNames = [
    'level 1',
    'level 2',
    'level 3',
    'level 4',
    'task',
    'l5',
    'ecc transaction',
    's/4hana transaction',
    's4hana transaction',
    'ocm valid',
    'change category',
  ];

  // Get indices of columns that should have filters
  const filterColumnIndices = useMemo(() => {
    return header.map((col, idx) => {
      const colLower = (col || '').toString().toLowerCase();
      const shouldShowFilter = filterColumnNames.some(name => colLower.includes(name));
      return shouldShowFilter ? idx : null;
    }).filter(idx => idx !== null);
  }, [header]);

  // Get unique values for each column (for dropdown filters)
  const uniqueValues = useMemo(() => {
    const values = {};
    header.forEach((col, idx) => {
      const colValues = new Set();
      allRows.forEach(row => {
        const val = (row[idx] || '').toString().trim();
        if (val) colValues.add(val);
      });
      values[idx] = Array.from(colValues).sort();
    });
    return values;
  }, [header, allRows]);

  // Filter rows based on selected filters (supports partial text matching and multi-select)
  const filteredRows = useMemo(() => {
    let rows = allRows.filter(row => {
      return Object.entries(filters).every(([colIdx, filterValue]) => {
        // If filterValue is empty string/array, show all
        if (!filterValue || (Array.isArray(filterValue) && filterValue.length === 0)) return true;

        const cellValue = (row[parseInt(colIdx)] || '').toString().trim().toLowerCase();

        // Handle array (multi-select): OR logic - match if cell value matches ANY selected option
        if (Array.isArray(filterValue)) {
          return filterValue.some(val => {
            const searchValue = val.toLowerCase();
            return cellValue.includes(searchValue);
          });
        }

        // Handle single value (legacy/text search)
        const searchValue = filterValue.toLowerCase();
        return cellValue.includes(searchValue);
      });
    });

    // Apply sorting if a column is selected
    if (sortColumn !== null) {
      rows = [...rows].sort((a, b) => {
        const aVal = (a[sortColumn] || '').toString().trim();
        const bVal = (b[sortColumn] || '').toString().trim();

        // Try numeric comparison first
        const aNum = parseFloat(aVal);
        const bNum = parseFloat(bVal);
        if (!isNaN(aNum) && !isNaN(bNum)) {
          return sortDirection === 'asc' ? aNum - bNum : bNum - aNum;
        }

        // Fallback to string comparison
        const comparison = aVal.localeCompare(bVal, undefined, { numeric: true, sensitivity: 'base' });
        return sortDirection === 'asc' ? comparison : -comparison;
      });
    }

    return rows;
  }, [allRows, filters, sortColumn, sortDirection]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const start = (page - 1) * pageSize;
  const pageRows = filteredRows.slice(start, start + pageSize);

  const handleFilterChange = (colIdx, value) => {
    setFilters(prev => ({
      ...prev,
      [colIdx]: value,
    }));
    setPage(1); // Reset to first page when filter changes
  };

  const handleSort = (colIdx) => {
    if (sortColumn === colIdx) {
      // Toggle direction if clicking same column
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      // New column, default to ascending
      setSortColumn(colIdx);
      setSortDirection('asc');
    }
    setPage(1); // Reset to first page when sorting changes
  };

  const clearAllFilters = () => {
    setFilters({});
    setPage(1);
  };

  const hasActiveFilters = Object.values(filters).some(v => v);

  // Download filtered data as Excel
  const handleDownloadFiltered = () => {
    // Create workbook and worksheet
    const wb = XLSX.utils.book_new();

    // Prepare data with header
    const wsData = [header, ...filteredRows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Auto-size columns
    const colWidths = header.map((h, idx) => {
      const maxLength = Math.max(
        (h || '').toString().length,
        ...filteredRows.slice(0, 100).map(row => (row[idx] || '').toString().length)
      );
      return { wch: Math.min(Math.max(maxLength, 10), 50) };
    });
    ws['!cols'] = colWidths;

    // Style header row - bold only
    header.forEach((_, idx) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: idx });
      if (ws[cellRef]) {
        ws[cellRef].s = { font: { bold: true } };
      }
    });

    // Add hyperlinks for URL cells
    filteredRows.forEach((row, rowIdx) => {
      row.forEach((cell, colIdx) => {
        const cellStr = (cell || '').toString();
        if (cellStr.startsWith('http://') || cellStr.startsWith('https://')) {
          const cellRef = XLSX.utils.encode_cell({ r: rowIdx + 1, c: colIdx });
          if (ws[cellRef]) {
            ws[cellRef].l = { Target: cellStr, Tooltip: cellStr };
          }
        }
      });
    });

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Filtered BPML Data');

    // Generate filename with filter info
    const filterCount = Object.values(filters).filter(v => v).length;
    const filename = filterCount > 0
      ? `BPML_Analysis_Filtered_${filteredRows.length}_rows.xlsx`
      : `BPML_Analysis_All_${filteredRows.length}_rows.xlsx`;

    // Download
    XLSX.writeFile(wb, filename);
  };

  return (
    <div style={styles.fullPageOverlay}>
      <div style={styles.fullPageContainer}>
        {/* Header */}
        <div style={styles.fullPageHeader}>
          <div>
            <h2 style={styles.fullPageTitle}>BPML Analysis</h2>
            <p style={styles.fullPageSubtitle}>
              Recommended BPML data with filters and analytics
            </p>
          </div>
          <button style={styles.closeButtonLarge} onClick={onClose} aria-label="Close">
            ✕
          </button>
        </div>

        {/* Scrollable Content */}
        <div style={styles.fullPageContent}>
          {/* Data Table Section */}
          <section style={styles.dataTableSection}>
            <div style={styles.tableHeaderRow}>
              <h3 style={styles.sectionTitle}>
                Recommended BPML Data ({filteredRows.length} of {allRows.length} rows)
              </h3>
              <div style={styles.tableHeaderActions}>
                {hasActiveFilters && (
                  <button style={styles.clearFiltersBtn} onClick={clearAllFilters}>
                    Clear All Filters
                  </button>
                )}
                <button style={styles.downloadFilteredBtn} onClick={handleDownloadFiltered}>
                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                  </svg>
                  Download Excel
                </button>
              </div>
            </div>

            {/* Filter Row */}
            <div style={styles.filterRow}>
              {filterColumnIndices.map((idx) => (
                <div key={idx} style={styles.filterItem}>
                  <label style={styles.filterLabel}>{header[idx] || `Column ${idx + 1}`}</label>
                  <SearchableDropdown
                    value={filters[idx] || []}
                    options={uniqueValues[idx] || []}
                    onChange={(val) => handleFilterChange(idx, val)}
                    placeholder="All"
                    multiSelect={true}
                  />
                </div>
              ))}
            </div>

            {/* Table */}
            <div style={styles.fullPageTableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    {header.map((h, idx) => (
                      <th
                        key={idx}
                        style={{
                          ...styles.th,
                          cursor: 'pointer',
                          userSelect: 'none',
                          position: 'relative',
                          paddingRight: '24px',
                        }}
                        onClick={() => handleSort(idx)}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <span>{h || '-'}</span>
                          <span style={{
                            fontSize: '12px',
                            opacity: sortColumn === idx ? 1 : 0.3,
                            transition: 'opacity 0.2s'
                          }}>
                            {sortColumn === idx && sortDirection === 'asc' ? '▲' : '▼'}
                          </span>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pageRows.map((row, idx) => (
                    <tr key={idx}>
                      {row.map((cell, cIdx) => (
                        <td key={cIdx} style={styles.td}>
                          {cell || '-'}
                        </td>
                      ))}
                    </tr>
                  ))}
                  {pageRows.length === 0 && (
                    <tr>
                      <td style={styles.td} colSpan={header.length || 1}>
                        No data matching filters
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            <div style={styles.pagination}>
              <button
                style={styles.secondaryButton}
                onClick={() => setPage(Math.max(1, page - 1))}
                disabled={page === 1}
              >
                Prev
              </button>
              <span style={{ fontWeight: 700 }}>
                Page {page} of {totalPages}
              </span>
              <button
                style={styles.secondaryButton}
                onClick={() => setPage(Math.min(totalPages, page + 1))}
                disabled={page === totalPages}
              >
                Next
              </button>
            </div>
          </section>

          {/* BPML Analytics Charts Section - Below the data table */}
          {(bpmlChartData.length > 0 || ocmChartData.length > 0 || changeCategoryChartData.length > 0) && (
            <section style={styles.analyticsSection}>
              <h3 style={styles.sectionTitle}>BPML Analytics</h3>
              <div style={styles.chartContainer}>
                {bpmlChartData.length > 0 && (
                  <div style={styles.chartCard}>
                    <h4 style={styles.chartTitle}>Unique Sub-Processes by Business Area</h4>
                    <p style={styles.chartSubtitle}>Level 1 vs Unique Level 4 Count</p>
                    <ResponsiveContainer width="100%" height={350}>
                      <PieChart>
                        <Pie
                          data={bpmlChartData}
                          cx="50%"
                          cy="50%"
                          labelLine={true}
                          label={({ name, value }) => {
                            const abbr = getBusinessAbbreviation(name);
                            return `${abbr}: ${value}`;
                          }}
                          outerRadius={120}
                          fill="#8884d8"
                          dataKey="value"
                        >
                          {bpmlChartData.map((entry, index) => (
                            <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                          ))}
                        </Pie>
                        <Tooltip
                          formatter={(value, name) => [`${value} unique sub-processes`, name]}
                        />
                      </PieChart>
                    </ResponsiveContainer>
                    {/* Custom Legend with abbreviations and full names */}
                    <div style={styles.pieChartLegend}>
                      {bpmlChartData.map((entry, index) => (
                        <div key={index} style={styles.pieChartLegendItem}>
                          <span
                            style={{
                              ...styles.pieChartLegendColor,
                              backgroundColor: CHART_COLORS[index % CHART_COLORS.length]
                            }}
                          />
                          <span style={styles.pieChartLegendAbbr}>{getBusinessAbbreviation(entry.name)}</span>
                          <span style={styles.pieChartLegendName}>- {entry.name}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {ocmChartData.length > 0 && (
                  <div style={styles.chartCard}>
                    <h4 style={styles.chartTitle}>OCM Valid Distribution by Business Area</h4>
                    <p style={styles.chartSubtitle}>Count of New, Change and Existing per Level 1 – Business/Enterprise Area</p>
                    <ResponsiveContainer width="100%" height={450}>
                      <BarChart data={ocmChartData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }} barCategoryGap="20%" barGap={2}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis
                          dataKey="shortCode"
                          interval={0}
                          tick={{ fontSize: 11, fontWeight: 600 }}
                          height={60}
                          tickMargin={10}
                        />
                        <YAxis allowDecimals={false} />
                        <Tooltip
                          content={({ active, payload }) => {
                            if (active && payload && payload.length) {
                              const d = payload[0].payload;
                              const total = d.New + d.Change + d.Existing;
                              return (
                                <div style={styles.customTooltip}>
                                  <div style={styles.tooltipTitle}>{d.fullName}</div>
                                  {payload.map((p, i) => (
                                    <div key={i} style={{ ...styles.tooltipValue, color: p.color }}>
                                      {p.name}: {p.value}
                                    </div>
                                  ))}
                                  <div style={{ ...styles.tooltipValue, borderTop: '1px solid #555', marginTop: 4, paddingTop: 4, color: '#fff', fontWeight: 'bold' }}>
                                    Total: {total}
                                  </div>
                                </div>
                              );
                            }
                            return null;
                          }}
                        />
                        <Legend wrapperStyle={{ paddingTop: 10, fontSize: 12 }} />
                        <Bar dataKey="New" fill="#5088e0ff" radius={[4, 4, 0, 0]} />
                        <Bar dataKey="Change" fill="#f02222ff" radius={[4, 4, 0, 0]} />
                        <Bar dataKey="Existing" fill="#64ce8bff" radius={[4, 4, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                    {/* Abbreviation legend (no colors - colors are for New/Change/Existing) */}
                    <div style={styles.pieChartLegend}>
                      {ocmChartData.map((entry, index) => (
                        <div key={index} style={styles.pieChartLegendItem}>
                          <span style={styles.pieChartLegendAbbr}>{entry.shortCode}</span>
                          <span style={styles.pieChartLegendName}>- {entry.fullName}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
                {changeCategoryChartData.length > 0 && stackedCategoryL1Keys.length > 0 && (
                  <div style={styles.chartCard}>
                    <h4 style={styles.chartTitle}>Change Category Distribution</h4>
                    <p style={styles.chartSubtitle}>Count by Change Category, stacked by Level 1 – Business/Enterprise Area</p>
                    <ResponsiveContainer width="100%" height={400}>
                      <BarChart data={changeCategoryChartData} margin={{ top: 20, right: 30, left: 20, bottom: 40 }}>
                        <CartesianGrid strokeDasharray="3 3" />
                        <XAxis
                          dataKey="code"
                          interval={0}
                          tick={{ fontSize: 11, fontWeight: 600 }}
                          height={50}
                          tickMargin={10}
                        />
                        <YAxis
                          allowDecimals={false}
                          ticks={(() => {
                            const max = Math.max(...changeCategoryChartData.map(d =>
                              stackedCategoryL1Keys.reduce((s, k) => s + (d[k] || 0), 0)
                            ));
                            return Array.from({ length: max + 1 }, (_, i) => i);
                          })()}
                        />
                        <Tooltip
                          content={({ active, payload }) => {
                            if (active && payload && payload.length) {
                              const d = payload[0]?.payload;
                              const total = payload.reduce((sum, p) => sum + (p.value || 0), 0);
                              return (
                                <div style={styles.customTooltip}>
                                  <div style={styles.tooltipTitle}>{d?.name || d?.code}</div>
                                  {payload.map((p, i) => (
                                    p.value > 0 && (
                                      <div key={i} style={{ ...styles.tooltipValue, color: p.color }}>
                                        {getBusinessAbbreviation(p.name)}: {p.value}
                                      </div>
                                    )
                                  ))}
                                  <div style={{ ...styles.tooltipValue, borderTop: '1px solid #555', marginTop: 4, paddingTop: 4, color: '#fff', fontWeight: 'bold' }}>
                                    Total: {total}
                                  </div>
                                </div>
                              );
                            }
                            return null;
                          }}
                        />
                        {stackedCategoryL1Keys.map((l1, index) => (
                          <Bar key={l1} dataKey={l1} stackId="a" fill={CHART_COLORS[index % CHART_COLORS.length]} shape={<StackedBarWithMinHeight />} />
                        ))}
                      </BarChart>
                    </ResponsiveContainer>
                    {/* Code legend below the chart */}
                    <div style={styles.pieChartLegend}>
                      {changeCategoryChartData.map((entry, index) => (
                        <div key={index} style={styles.pieChartLegendItem}>
                          <span style={styles.pieChartLegendAbbr}>{entry.code}</span>
                          <span style={styles.pieChartLegendName}>- {entry.name}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}

function TestingScopePage({
  data,
  onClose,
  ocmL5ChartData,
}) {
  const [selectedL1, setSelectedL1] = useState('');
  const [page, setPage] = useState(1);
  const [sortColumn, setSortColumn] = useState(null);
  const [sortDirection, setSortDirection] = useState('asc');
  const pageSize = 25;

  const { uniqueL1Values = [], rows = [], l1ColName, l4ColName, l5ColName } = data || {};

  // Filter rows by selected L1 (data is already filtered for OCM Valid = Change and valid Change Categories)
  const filteredData = useMemo(() => {
    if (!selectedL1) return [];
    return rows.filter(row => row.l1 === selectedL1);
  }, [selectedL1, rows]);

  // Apply sorting
  const sortedData = useMemo(() => {
    if (sortColumn === null) return filteredData;

    return [...filteredData].sort((a, b) => {
      const aVal = sortColumn === 'l1' ? a.l1 : sortColumn === 'l4' ? a.l4 : a.l5;
      const bVal = sortColumn === 'l1' ? b.l1 : sortColumn === 'l4' ? b.l4 : b.l5;

      const comparison = aVal.localeCompare(bVal, undefined, { numeric: true, sensitivity: 'base' });
      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }, [filteredData, sortColumn, sortDirection]);

  const totalPages = Math.max(1, Math.ceil(sortedData.length / pageSize));
  const start = (page - 1) * pageSize;
  const pageRows = sortedData.slice(start, start + pageSize);

  const handleSort = (col) => {
    if (sortColumn === col) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc');
    } else {
      setSortColumn(col);
      setSortDirection('asc');
    }
  };

  // Download filtered data as Excel
  const handleDownload = () => {
    if (sortedData.length === 0) return;

    const headerRow = [l1ColName || 'Level 1', l4ColName || 'Level 4', l5ColName || 'Task (L5)'];
    const wsData = [headerRow, ...sortedData.map(row => [row.l1, row.l4, row.l5])];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set column widths
    ws['!cols'] = headerRow.map((h, idx) => ({
      wch: Math.max(h.length, ...sortedData.slice(0, 100).map(row =>
        (idx === 0 ? row.l1 : idx === 1 ? row.l4 : row.l5).length
      )) + 2
    }));

    // Style header row - bold only
    headerRow.forEach((_, idx) => {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c: idx });
      if (ws[cellRef]) {
        ws[cellRef].s = { font: { bold: true } };
      }
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Testing Scope');
    XLSX.writeFile(wb, `Testing_Scope_${selectedL1.replace(/[^a-zA-Z0-9]/g, '_')}_${sortedData.length}_tasks.xlsx`);
  };

  return (
    <div style={styles.previewOverlay}>
      <div style={{ ...styles.previewCard, overflowY: 'auto' }}>
        <div style={styles.previewHeader}>
          <div>
            <div style={styles.sectionTitle}>Testing Scope Dashboard</div>
            <div style={styles.outputsSub}>Select a Business Area to view unique tasks</div>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            {selectedL1 && sortedData.length > 0 && (
              <button style={styles.actionButton} onClick={handleDownload}>
                Download Excel
              </button>
            )}
            <button style={styles.closeButton} onClick={onClose} aria-label="Close">
              ✕
            </button>
          </div>
        </div>

        {/* L1 Filter Dropdown */}
        <div style={{ marginBottom: 16 }}>
          <label style={{ fontWeight: 600, marginRight: 12, color: '#1A1A1A' }}>
            {l1ColName || 'Level 1 - Business/Enterprise Area'}:
          </label>
          <select
            value={selectedL1}
            onChange={(e) => {
              setSelectedL1(e.target.value);
              setPage(1);
              setSortColumn(null);
            }}
            style={{
              padding: '10px 16px',
              borderRadius: 8,
              border: '1px solid #FFCDA8',
              background: '#fff',
              fontSize: 14,
              minWidth: 300,
              cursor: 'pointer',
            }}
          >
            <option value="">-- Select Business Area --</option>
            {uniqueL1Values.map((l1, idx) => (
              <option key={idx} value={l1}>{l1}</option>
            ))}
          </select>
          {selectedL1 && (
            <span style={{ marginLeft: 16, color: '#666', fontSize: 14 }}>
              {sortedData.length} unique tasks found
            </span>
          )}
        </div>

        {/* Data Table */}
        {selectedL1 && (
          <div style={{ ...styles.tableWrap, maxHeight: '35vh', flex: 'none' }}>
            <table style={styles.table}>
              <thead>
                <tr>
                  {[
                    { key: 'l1', label: l1ColName || 'Level 1' },
                    { key: 'l4', label: l4ColName || 'Level 4' },
                    { key: 'l5', label: l5ColName || 'Task (L5)' },
                  ].map(({ key, label }) => (
                    <th
                      key={key}
                      style={{
                        ...styles.th,
                        cursor: 'pointer',
                        userSelect: 'none',
                      }}
                      onClick={() => handleSort(key)}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <span>{label}</span>
                        <span style={{
                          fontSize: 12,
                          opacity: sortColumn === key ? 1 : 0.3,
                        }}>
                          {sortColumn === key && sortDirection === 'asc' ? '▲' : '▼'}
                        </span>
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pageRows.map((row, idx) => (
                  <tr key={idx}>
                    <td style={styles.td}>{row.l1}</td>
                    <td style={styles.td}>{row.l4}</td>
                    <td style={styles.td}>{row.l5}</td>
                  </tr>
                ))}
                {pageRows.length === 0 && (
                  <tr>
                    <td style={styles.td} colSpan={3}>
                      No tasks found for selected business area
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Pagination */}
        {selectedL1 && sortedData.length > 0 && (
          <div style={styles.pagination}>
            <button
              style={styles.secondaryButton}
              onClick={() => setPage(Math.max(1, page - 1))}
              disabled={page === 1}
            >
              Prev
            </button>
            <span style={{ fontWeight: 700 }}>
              Page {page} of {totalPages}
            </span>
            <button
              style={styles.secondaryButton}
              onClick={() => setPage(Math.min(totalPages, page + 1))}
              disabled={page === totalPages}
            >
              Next
            </button>
          </div>
        )}

        {/* L1 vs Unique L5 Bar Chart */}
        {ocmL5ChartData && ocmL5ChartData.length > 0 && (
          <section style={{ marginTop: 16, flex: 'none' }}>
            <div style={styles.chartCard}>
              <h4 style={styles.chartTitle}>OCM Change Impact - Level 5 Tasks</h4>
              <p style={styles.chartSubtitle}>Unique L5 count by Business Area (OCM Valid = 'Change')</p>
              <ResponsiveContainer width="100%" height={300}>
                <BarChart data={ocmL5ChartData} margin={{ top: 15, right: 25, left: 15, bottom: 40 }} barCategoryGap="15%">
                  <CartesianGrid strokeDasharray="3 3" />
                  <XAxis
                    dataKey="shortCode"
                    interval={0}
                    tick={{ fontSize: 10, fontWeight: 600 }}
                    height={40}
                    tickMargin={8}
                  />
                  <YAxis />
                  <Tooltip
                    content={({ active, payload }) => {
                      if (active && payload && payload.length) {
                        const d = payload[0].payload;
                        return (
                          <div style={styles.customTooltip}>
                            <div style={styles.tooltipTitle}>{d.fullName}</div>
                            <div style={styles.tooltipValue}>{d.count} unique tasks (L5)</div>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Bar dataKey="count" fill="#FD5108" radius={[4, 4, 0, 0]}>
                    {ocmL5ChartData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                    ))}
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </section>
        )}
      </div>
    </div>
  );
}

const styles = {
  shell: {
    minHeight: '100vh',
    background: '#FFF5ED',
    padding: '32px 40px 56px',
    fontFamily: 'Inter, system-ui, -apple-system, sans-serif',
    color: '#1A1A1A',
    boxSizing: 'border-box',
  },
  hero: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    background: '#FAF7F4',
    borderRadius: 20,
    padding: 24,
    boxShadow: '0 8px 24px rgba(0,0,0,0.06)',
    marginBottom: 20,
    gap: 16,
  },
  badge: {
    display: 'inline-block',
    background: '#FFE8D4',
    color: '#FD5108',
    padding: '6px 10px',
    borderRadius: 999,
    fontWeight: 600,
    fontSize: 12,
    border: '1px solid #FFCDA8',
  },
  title: {
    margin: '12px 0 8px',
    fontSize: 28,
    fontWeight: 700,
    color: '#1A1A1A',
  },
  subtitle: {
    margin: 0,
    color: '#4A4A4A',
    maxWidth: 840,
    lineHeight: 1.4,
  },
  logo: {
    width: 96,
    height: 64,
    objectFit: 'contain',
    background: '#FAF7F4',
    borderRadius: 12,
    padding: 8,
  },
  card: {
    background: '#F7F3EF',
    borderRadius: 16,
    padding: 20,
    boxShadow: '0 8px 22px rgba(0,0,0,0.06)',
    marginBottom: 20,
  },
  sectionHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 12,
  },
  sectionTitle: {
    margin: 0,
    fontSize: 18,
    fontWeight: 700,
  },
  collapse: {
    background: 'transparent',
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 10,
    padding: '8px 12px',
    cursor: 'pointer',
    color: '#4A4A4A',
    fontWeight: 600,
  },
  cardGrid: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: 12,
  },
  dropCard: {
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 14,
    padding: 14,
    background: '#F7F3EF',
    boxShadow: '0 6px 16px rgba(0,0,0,0.04)',
    cursor: 'pointer',
  },
  cardTopRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 6,
  },
  cardTitle: {
    fontWeight: 700,
    color: '#1A1A1A',
  },
  ready: {
    background: '#ecfdf3',
    color: '#16a34a',
    borderRadius: 8,
    padding: '4px 8px',
    fontSize: 12,
    fontWeight: 700,
  },
  pending: {
    background: '#FFE8D4',
    color: '#FD5108',
    borderRadius: 8,
    padding: '4px 8px',
    fontSize: 12,
    fontWeight: 700,
  },
  cardHelp: {
    margin: '0 0 10px',
    color: '#4A4A4A',
    fontSize: 13,
  },
  dropZone: {
    border: '1px dashed rgba(0,0,0,0.12)',
    borderRadius: 12,
    padding: 12,
    display: 'flex',
    gap: 10,
    alignItems: 'center',
    background: '#FAF7F4',
  },
  dropIcon: {
    width: 28,
    height: 28,
    borderRadius: 12,
    background: '#F7F3EF',
    display: 'grid',
    placeItems: 'center',
    color: '#8A8A8A',
    fontWeight: 700,
  },
  dropText: {
    fontWeight: 700,
    color: '#1A1A1A',
  },
  dropSub: {
    fontSize: 12,
    color: '#4A4A4A',
  },
  fileName: {
    marginTop: 6,
    fontSize: 12,
    color: '#FD5108',
    fontWeight: 700,
  },
  outputsCard: {
    background: '#F7F3EF',
    borderRadius: 16,
    padding: 20,
    boxShadow: '0 10px 26px rgba(0,0,0,0.08)',
    display: 'flex',
    flexDirection: 'column',
    gap: 14,
  },
  outputsHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 12,
  },
  outputsSub: {
    margin: 0,
    color: '#4A4A4A',
    fontSize: 14,
  },
  runBtn: {
    background: '#FD5108',
    color: '#fff',
    border: 'none',
    padding: '12px 16px',
    borderRadius: 12,
    fontWeight: 700,
    cursor: 'pointer',
    boxShadow: '0 10px 20px rgba(253,81,8,0.24)',
    minWidth: 150,
  },
  progressWrap: {
    background: '#FFE8D4',
    borderRadius: 12,
    padding: 14,
    border: '1px solid #FFCDA8',
  },
  progressMeta: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 8,
  },
  progressLabel: { color: '#4A4A4A', fontWeight: 600 },
  progressPct: { color: '#FD5108', fontWeight: 800 },
  progressBarOuter: {
    width: '100%',
    height: 10,
    background: '#FFCDA8',
    borderRadius: 999,
    overflow: 'hidden',
  },
  progressBarInner: {
    height: '100%',
    background: 'linear-gradient(90deg, #FD5108, #FE7C39)',
    borderRadius: 999,
  },
  summaryRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
    gap: 12,
  },
  summaryCard: {
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 14,
    padding: 14,
    background: '#FAF7F4',
    boxShadow: '0 6px 16px rgba(0,0,0,0.04)',
  },
  summaryTop: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 8,
  },
  summaryTitle: { fontWeight: 700, color: '#1A1A1A' },
  badgeSuccess: {
    background: '#ecfdf3',
    color: '#16a34a',
    borderRadius: 10,
    padding: '6px 10px',
    fontSize: 12,
    fontWeight: 800,
  },
  badgePending: {
    background: '#FFE8D4',
    color: '#FD5108',
    borderRadius: 10,
    padding: '6px 10px',
    fontSize: 12,
    fontWeight: 800,
  },
  summaryFooter: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginTop: 8,
    color: '#4A4A4A',
  },
  summaryValue: { fontWeight: 800, color: '#1A1A1A' },
  actionsRow: {
    display: 'flex',
    gap: 12,
    alignItems: 'center',
    flexWrap: 'wrap',
  },
  actionButton: {
    background: '#FD5108',
    color: '#fff',
    border: 'none',
    padding: '12px 16px',
    borderRadius: 12,
    fontWeight: 700,
    cursor: 'pointer',
    boxShadow: '0 10px 20px rgba(253,81,8,0.24)',
  },
  disabledButton: {
    opacity: 0.5,
    cursor: 'not-allowed',
    pointerEvents: 'none',
  },
  secondaryButton: {
    background: '#F7F3EF',
    color: '#1A1A1A',
    border: '1px solid rgba(0,0,0,0.12)',
    padding: '10px 14px',
    borderRadius: 10,
    fontWeight: 700,
    cursor: 'pointer',
  },
  closeButton: {
    background: '#F7F3EF',
    color: '#1A1A1A',
    border: '1px solid rgba(0,0,0,0.12)',
    width: 40,
    height: 40,
    borderRadius: 10,
    fontWeight: 700,
    cursor: 'pointer',
    fontSize: 18,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  backButton: {
    background: '#F7F3EF',
    color: '#1A1A1A',
    border: '1px solid rgba(0,0,0,0.12)',
    padding: '10px 20px',
    borderRadius: 10,
    fontWeight: 700,
    cursor: 'pointer',
    fontSize: 14,
    display: 'flex',
    alignItems: 'center',
    gap: 6,
  },
  secondaryLink: {
    color: '#FD5108',
    fontWeight: 700,
    textDecoration: 'none',
  },
  error: {
    background: '#fef2f2',
    border: '1px solid #fecdd3',
    color: '#b91c1c',
    padding: 12,
    borderRadius: 12,
    fontWeight: 700,
  },
  detailCard: {
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 14,
    padding: 14,
    background: '#FAF7F4',
  },
  detailHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 10,
  },
  detailTitle: { fontWeight: 800, color: '#1A1A1A' },
  detailSub: { color: '#4A4A4A', fontSize: 13 },
  badgeSoft: {
    background: '#FFE8D4',
    color: '#FD5108',
    padding: '6px 10px',
    borderRadius: 10,
    fontWeight: 700,
  },
  tableRow: {
    display: 'flex',
    justifyContent: 'space-between',
    padding: '8px 0',
    borderBottom: '1px solid rgba(0,0,0,0.12)',
    fontSize: 14,
  },
  tableLabel: { fontWeight: 700, color: '#1A1A1A' },
  tableValue: { color: '#4A4A4A' },
  analysisBox: {
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 10,
    padding: 10,
    background: '#FAF7F4',
  },
  analysisRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: 8,
    padding: '4px 0',
    fontSize: 13,
  },
  previewOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100%',
    height: '100%',
    background: '#FFF5ED',
    display: 'flex',
    flexDirection: 'column',
    boxSizing: 'border-box',
    zIndex: 1000,
  },
  previewCard: {
    background: '#FFF5ED',
    padding: '20px 32px',
    width: '100%',
    height: '100%',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    gap: 12,
    boxSizing: 'border-box',
  },
  previewHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
  },
  tableWrap: {
    border: '1px solid #FFCDA8',
    borderRadius: 12,
    overflow: 'auto',
    flex: 1,
    minHeight: 0,
  },
  table: {
    width: '100%',
    borderCollapse: 'collapse',
    fontSize: 14,
  },
  th: {
    textAlign: 'left',
    background: '#FFE8D4',
    color: '#FD5108',
    padding: '10px 12px',
    borderBottom: '1px solid #FFCDA8',
    position: 'sticky',
    top: 0,
  },
  td: {
    padding: '10px 12px',
    borderBottom: '1px solid rgba(0,0,0,0.12)',
  },
  pagination: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginTop: 8,
  },
  previewChips: {
    display: 'flex',
    gap: 8,
    flexWrap: 'wrap',
  },
  previewChip: {
    background: '#FFE8D4',
    color: '#FD5108',
    border: '1px solid #FFCDA8',
    borderRadius: 999,
    padding: '6px 12px',
    fontWeight: 700,
    fontSize: 12,
  },
  sheetTabs: {
    display: 'flex',
    gap: 8,
    borderBottom: '2px solid #FFCDA8',
    marginBottom: 12,
  },
  sheetTab: {
    background: 'transparent',
    border: 'none',
    borderBottom: '3px solid transparent',
    padding: '10px 16px',
    cursor: 'pointer',
    fontWeight: 600,
    color: '#4A4A4A',
    transition: 'all 0.2s',
  },
  sheetTabActive: {
    borderBottomColor: '#FD5108',
    color: '#FD5108',
    fontWeight: 700,
  },
  dropdownWrapper: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
  },
  dropdown: {
    width: '100%',
    maxWidth: 500,
    padding: '12px 16px',
    fontSize: 14,
    fontWeight: 600,
    color: '#1A1A1A',
    background: '#FAF7F4',
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 12,
    cursor: 'pointer',
    outline: 'none',
    boxShadow: '0 4px 12px rgba(0,0,0,0.04)',
    transition: 'all 0.2s',
  },
  loadingText: {
    color: '#4A4A4A',
    fontSize: 14,
    fontWeight: 600,
  },
  deleteFileBtn: {
    background: '#fef2f2',
    color: '#dc2626',
    border: '1px solid #fecaca',
    borderRadius: 6,
    width: 24,
    height: 24,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    cursor: 'pointer',
    fontSize: 14,
    fontWeight: 700,
    transition: 'all 0.2s',
    padding: 0,
  },
  fioriBtn: {
    background: '#FD5108',
    color: '#fff',
    border: 'none',
    padding: '12px 16px',
    borderRadius: 12,
    fontWeight: 700,
    cursor: 'pointer',
    boxShadow: '0 10px 20px rgba(253,81,8,0.24)',
    minWidth: 150,
  },
  chartContainer: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fit, minmax(550px, 1fr))',
    gap: 20,
    marginTop: 12,
  },
  chartCard: {
    background: '#FAF7F4',
    borderRadius: 14,
    padding: 24,
    border: '1px solid rgba(0,0,0,0.12)',
    boxShadow: '0 6px 16px rgba(0,0,0,0.04)',
    minHeight: 450,
  },
  chartTitle: {
    margin: '0 0 4px',
    fontSize: 16,
    fontWeight: 700,
    color: '#1A1A1A',
  },
  chartSubtitle: {
    margin: '0 0 16px',
    fontSize: 13,
    color: '#4A4A4A',
  },
  pieChartLegend: {
    display: 'flex',
    flexDirection: 'column',
    gap: 8,
    marginTop: 16,
    padding: '12px 16px',
    background: '#fff',
    borderRadius: 8,
    border: '1px solid rgba(0,0,0,0.08)',
  },
  pieChartLegendItem: {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    fontSize: 13,
  },
  pieChartLegendColor: {
    width: 14,
    height: 14,
    borderRadius: 3,
    flexShrink: 0,
  },
  pieChartLegendAbbr: {
    fontWeight: 700,
    color: '#1A1A1A',
    minWidth: 45,
  },
  pieChartLegendName: {
    color: '#4A4A4A',
  },
  customTooltip: {
    background: '#1A1A1A',
    borderRadius: 8,
    padding: '10px 14px',
    boxShadow: '0 4px 12px rgba(0,0,0,0.15)',
  },
  tooltipTitle: {
    color: '#fff',
    fontWeight: 600,
    fontSize: 13,
    marginBottom: 4,
  },
  tooltipValue: {
    color: '#FFB347',
    fontWeight: 700,
    fontSize: 14,
  },
  tableLink: {
    color: '#FD5108',
    fontWeight: 600,
    textDecoration: 'none',
    padding: '4px 10px',
    background: '#FFE8D4',
    borderRadius: 6,
    fontSize: 12,
    display: 'inline-block',
    transition: 'all 0.2s',
  },
  linkContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: 6,
  },
  // BPML Analysis Full Page Styles
  fullPageOverlay: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100vw',
    height: '100vh',
    background: '#FFF5ED',
    zIndex: 1000,
    overflow: 'hidden',
  },
  fullPageContainer: {
    width: '100%',
    height: '100%',
    display: 'flex',
    flexDirection: 'column',
  },
  fullPageHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '20px 32px',
    background: '#FAF7F4',
    borderBottom: '1px solid #FFCDA8',
    flexShrink: 0,
  },
  fullPageTitle: {
    margin: 0,
    fontSize: 24,
    fontWeight: 700,
    color: '#1A1A1A',
  },
  fullPageSubtitle: {
    margin: '4px 0 0',
    color: '#4A4A4A',
    fontSize: 14,
  },
  closeButtonLarge: {
    background: '#F7F3EF',
    color: '#1A1A1A',
    border: '1px solid rgba(0,0,0,0.12)',
    width: 48,
    height: 48,
    borderRadius: 12,
    fontWeight: 700,
    cursor: 'pointer',
    fontSize: 20,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    transition: 'all 0.2s',
  },
  fullPageContent: {
    flex: 1,
    overflow: 'auto',
    padding: '24px 32px',
  },
  analyticsSection: {
    marginTop: 32,
    marginBottom: 32,
  },
  dataTableSection: {
    background: '#F7F3EF',
    borderRadius: 16,
    padding: 20,
    boxShadow: '0 8px 22px rgba(0,0,0,0.06)',
  },
  tableHeaderRow: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 16,
    flexWrap: 'wrap',
    gap: 12,
  },
  tableHeaderActions: {
    display: 'flex',
    alignItems: 'center',
    gap: 10,
    flexWrap: 'wrap',
  },
  downloadFilteredBtn: {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    background: '#FD5108',
    color: '#fff',
    border: 'none',
    padding: '10px 16px',
    borderRadius: 10,
    fontWeight: 600,
    cursor: 'pointer',
    fontSize: 13,
    boxShadow: '0 4px 12px rgba(253, 81, 8, 0.25)',
    transition: 'all 0.2s ease',
  },
  filterRow: {
    display: 'grid',
    gridTemplateColumns: 'repeat(auto-fill, minmax(180px, 1fr))',
    gap: 12,
    marginBottom: 16,
    padding: 16,
    background: '#FAF7F4',
    borderRadius: 12,
    border: '1px solid #FFCDA8',
  },
  filterItem: {
    display: 'flex',
    flexDirection: 'column',
    gap: 4,
  },
  filterLabel: {
    fontSize: 11,
    fontWeight: 700,
    color: '#FD5108',
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
  },
  filterSelect: {
    padding: '8px 12px',
    fontSize: 13,
    fontWeight: 500,
    color: '#1A1A1A',
    background: '#fff',
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 8,
    cursor: 'pointer',
    outline: 'none',
  },
  comboboxWrapper: {
    position: 'relative',
    display: 'flex',
    alignItems: 'center',
  },
  comboboxInput: {
    width: '100%',
    padding: '8px 32px 8px 12px',
    fontSize: 13,
    fontWeight: 500,
    color: '#1A1A1A',
    background: '#fff',
    border: '1px solid rgba(0,0,0,0.12)',
    borderRadius: 8,
    outline: 'none',
    boxSizing: 'border-box',
  },
  clearInputBtn: {
    position: 'absolute',
    right: 8,
    background: 'transparent',
    border: 'none',
    color: '#999',
    cursor: 'pointer',
    fontSize: 12,
    padding: 4,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
  },
  // Modern Searchable Dropdown Styles
  searchableDropdown: {
    position: 'relative',
    width: '100%',
  },
  dropdownInputWrapper: {
    display: 'flex',
    alignItems: 'center',
    background: '#fff',
    border: '1px solid #e0e0e0',
    borderRadius: 10,
    padding: '0 4px 0 0',
    transition: 'all 0.2s ease',
    boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
  },
  dropdownInputWrapperFocused: {
    borderColor: '#FD5108',
    boxShadow: '0 0 0 3px rgba(253, 81, 8, 0.1), 0 1px 3px rgba(0,0,0,0.08)',
  },
  dropdownInput: {
    flex: 1,
    border: 'none',
    outline: 'none',
    padding: '10px 12px',
    fontSize: 13,
    fontWeight: 500,
    color: '#1A1A1A',
    background: 'transparent',
    borderRadius: 10,
    minWidth: 0,
  },
  selectedTagsContainer: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: '4px',
    padding: '4px 0 4px 8px',
    flex: '1 1 auto',
    minWidth: 0,
  },
  selectedTag: {
    display: 'inline-flex',
    alignItems: 'center',
    gap: '4px',
    background: '#FFF5ED',
    color: '#FD5108',
    border: '1px solid #FFCDA8',
    borderRadius: '6px',
    padding: '2px 4px 2px 8px',
    fontSize: '12px',
    fontWeight: '500',
    maxWidth: '150px',
    whiteSpace: 'nowrap',
  },
  selectedTagText: {
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    whiteSpace: 'nowrap',
  },
  selectedTagRemove: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'transparent',
    border: 'none',
    color: '#FD5108',
    cursor: 'pointer',
    fontSize: '18px',
    fontWeight: 'bold',
    padding: '0',
    width: '20px',
    height: '20px',
    borderRadius: '4px',
    transition: 'background 0.15s ease',
  },
  dropdownClearBtn: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: '#f5f5f5',
    border: 'none',
    borderRadius: 6,
    width: 24,
    height: 24,
    cursor: 'pointer',
    color: '#666',
    transition: 'all 0.15s ease',
    marginRight: 4,
  },
  dropdownToggleBtn: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'transparent',
    border: 'none',
    borderRadius: 6,
    width: 28,
    height: 28,
    cursor: 'pointer',
    color: '#888',
    transition: 'all 0.15s ease',
  },
  dropdownMenu: {
    position: 'absolute',
    top: 'calc(100% + 6px)',
    left: 0,
    right: 0,
    background: '#fff',
    border: '1px solid #e0e0e0',
    borderRadius: 12,
    boxShadow: '0 10px 40px rgba(0,0,0,0.12), 0 2px 8px rgba(0,0,0,0.08)',
    zIndex: 1000,
    overflow: 'hidden',
  },
  dropdownOptionsList: {
    maxHeight: 240,
    overflowY: 'auto',
  },
  dropdownOption: {
    padding: '10px 14px',
    fontSize: 13,
    color: '#333',
    cursor: 'pointer',
    transition: 'all 0.1s ease',
    borderBottom: '1px solid #f5f5f5',
  },
  dropdownOptionSelected: {
    background: '#FFF5ED',
    color: '#FD5108',
    fontWeight: 600,
  },
  dropdownOptionHovered: {
    background: '#f8f8f8',
  },
  dropdownOptionAll: {
    color: '#888',
    fontStyle: 'italic',
  },
  dropdownNoResults: {
    padding: '14px',
    fontSize: 13,
    color: '#999',
    textAlign: 'center',
    fontStyle: 'italic',
  },
  dropdownMoreText: {
    padding: '10px 14px',
    fontSize: 12,
    color: '#888',
    textAlign: 'center',
    background: '#fafafa',
    borderTop: '1px solid #eee',
  },
  clearFiltersBtn: {
    background: '#FFE8D4',
    color: '#FD5108',
    border: '1px solid #FFCDA8',
    padding: '10px 16px',
    borderRadius: 10,
    fontWeight: 700,
    cursor: 'pointer',
    fontSize: 13,
  },
  fullPageTableWrap: {
    border: '1px solid #FFCDA8',
    borderRadius: 12,
    overflow: 'auto',
    maxHeight: '50vh',
  },
};

