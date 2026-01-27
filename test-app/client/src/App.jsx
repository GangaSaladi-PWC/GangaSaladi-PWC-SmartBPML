import { useMemo, useRef, useState, useEffect } from 'react';
import pwcLogo from './assets/pwc-logo.png';
import * as XLSX from 'xlsx';
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
  const pageSize = 10;
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
  const [changeCategoryChartData, setChangeCategoryChartData] = useState([]); // Bar chart data for Change Category

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
        return;
      }

      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      if (rows.length < 2) return;

      const headers = rows[0];

      // Find column indices
      const l1Idx = headers.findIndex(h => h && h.toString().toLowerCase().includes('level 1'));
      const l4Idx = headers.findIndex(h => h && h.toString().toLowerCase().includes('level 4'));
      const ocmIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('ocm valid'));

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

      // Chart 2: Level 1 vs OCM Valid (not 'Same' or 'Existing')
     if (l1Idx !== -1 && ocmIdx !== -1) {
  const l1OcmMap = new Map();

  rows.slice(1).forEach((row) => {
    const l1 = cleanL1Name((row[l1Idx] || '').toString().trim());

    // normalize OCM Valid safely
    const rawOcm = row[ocmIdx];
    const ocmValid = (rawOcm ?? '').toString().trim().toLowerCase();

    //  skip blanks / nulls / placeholders
    const isBlank =
      !ocmValid ||
      ocmValid === '-' ||
      ocmValid === 'na' ||
      ocmValid === 'n/a' ||
      ocmValid === 'null' ||
      ocmValid === 'undefined';

    //  skip "same" and "existing"
    const isExcluded = ocmValid === 'same' || ocmValid === 'existing';

    if (!l1 || isBlank || isExcluded) return;

    l1OcmMap.set(l1, (l1OcmMap.get(l1) || 0) + 1);
  });

        // Standard business process abbreviations
        const businessAbbreviations = {
          'procure to pay': 'P2P',
          'purchase to pay': 'P2P',
          'procurement': 'P2P',
          'order to cash': 'OTC',
          'record to report': 'RTR',
          'hire to retire': 'H2R',
          'plan to produce': 'P2P',
          'plan to product': 'P2P',
          'source to pay': 'S2P',
          'acquire to retire': 'A2R',
          'finance': 'FIN',
          'finance & controlling': 'FICO',
          'financial accounting': 'FI',
          'controlling': 'CO',
          'sales': 'SD',
          'sales & distribution': 'SD',
          'materials management': 'MM',
          'production planning': 'PP',
          'plant maintenance': 'PM',
          'quality management': 'QM',
          'human resources': 'HR',
          'human capital': 'HCM',
          'supply chain': 'SCM',
          'supply chain management': 'SCM',
          'warehouse management': 'WM',
          'extended warehouse': 'EWM',
          'transportation': 'TM',
          'project system': 'PS',
          'asset management': 'EAM',
          'enterprise asset': 'EAM',
          'customer service': 'CS',
          'global trade': 'GTS',
          'treasury': 'TRM',
          'real estate': 'RE',
          'environment health safety': 'EHS',
        };

        const getAbbreviation = (name, index) => {
          const lower = name.toLowerCase();
          for (const [key, abbr] of Object.entries(businessAbbreviations)) {
            if (lower.includes(key)) {
              return abbr;
            }
          }
          // Fallback: take first letter of each word
          const words = name.split(/[\s&,\-()]+/).filter(w => w.length > 1);
          if (words.length >= 2) {
            return words.slice(0, 3).map(w => w[0]).join('').toUpperCase();
          }
          return name.substring(0, 3).toUpperCase();
        };

        // Track used abbreviations to avoid duplicates
        const usedAbbrs = new Map();
        const ocmData = Array.from(l1OcmMap.entries())
          .filter(([l1, count]) => l1 && count > 0)
          .map(([l1, count], index) => {
            let abbr = getAbbreviation(l1, index);
            // Handle duplicates by adding a number
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
              count: count,
            };
          });
        setOcmChartData(ocmData);
        console.log('Bar chart data (OCM):', ocmData);
      }

      // Chart 3: Change Category counts with descriptions
      const changeCategoryIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes('change category'));
      if (changeCategoryIdx !== -1) {
        const categoryDescriptions = {
          'NCG': 'No Change',
          'SCH_R': 'Screen Change – Revised',
          'SCH_N': 'Screen Change – New',
          'TRN_N': 'New Transaction – New',
          'TRN_C': 'New Transaction – Changed',
          'PRC': 'Process Change',
          'FRC': 'Fiori Change',
        };

        const categoryCountMap = new Map();
        rows.slice(1).forEach(row => {
          const category = (row[changeCategoryIdx] || '').toString().trim().toUpperCase();
          if (category) {
            categoryCountMap.set(category, (categoryCountMap.get(category) || 0) + 1);
          }
        });

        const changeCategoryData = Array.from(categoryCountMap.entries())
          .filter(([code, count]) => code && count > 0)
          .map(([code, count]) => ({
            name: categoryDescriptions[code] || code,
            code: code,
            count: count,
          }));
        setChangeCategoryChartData(changeCategoryData);
        console.log('Bar chart data (Change Category):', changeCategoryData);
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

    // Clear chart data if BPML file is deleted
    if (cardIdx === 0) {
      setBpmlChartData([]);
      setOcmChartData([]);
      setChangeCategoryChartData([]);
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
          title={previewIdx === 'fiori' ? 'Fiori Mapping Preview' : 'Merged preview'}
        />
      )}

      {/* BPML Visualizations */}
      {(bpmlChartData.length > 0 || ocmChartData.length > 0 || changeCategoryChartData.length > 0) && (
        <section style={styles.card}>
          <div style={styles.sectionHeader}>
            <h3 style={styles.sectionTitle}>BPML Analytics</h3>
          </div>
          <div style={styles.chartContainer}>
            {bpmlChartData.length > 0 && (
              <div style={styles.chartCard}>
                <h4 style={styles.chartTitle}>Unique Sub-Processes by Business Area</h4>
                <p style={styles.chartSubtitle}>Level 1 vs Unique Level 4 Count</p>
                <ResponsiveContainer width="100%" height={400}>
                  <PieChart>
                    <Pie
                      data={bpmlChartData}
                      cx="50%"
                      cy="50%"
                      labelLine={true}
                      label={({ name, value, percent }) => `${name}: ${value} (${(percent * 100).toFixed(0)}%)`}
                      outerRadius={120}
                      fill="#8884d8"
                      dataKey="value"
                    >
                      {bpmlChartData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                      ))}
                    </Pie>
                    <Tooltip formatter={(value) => [`${value} unique sub-processes`, 'Count']} />
                    <Legend />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            )}
            {ocmChartData.length > 0 && (
              <div style={styles.chartCard}>
                <h4 style={styles.chartTitle}>OCM Changes by Business Area</h4>
                <p style={styles.chartSubtitle}>Level 1 where OCM Valid is not Same or Existing (hover for details)</p>
                <ResponsiveContainer width="100%" height={450}>
                  <BarChart data={ocmChartData} margin={{ top: 20, right: 30, left: 20, bottom: 60 }} barCategoryGap="15%">
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis
                      dataKey="shortCode"
                      interval={0}
                      tick={{ fontSize: 11, fontWeight: 600 }}
                      height={60}
                      tickMargin={10}
                    />
                    <YAxis />
                    <Tooltip
                      content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          const data = payload[0].payload;
                          return (
                            <div style={styles.customTooltip}>
                              <div style={styles.tooltipTitle}>{data.fullName}</div>
                              <div style={styles.tooltipValue}>{data.count} items</div>
                            </div>
                          );
                        }
                        return null;
                      }}
                    />
                    <Bar dataKey="count" fill="#FD5108" radius={[4, 4, 0, 0]}>
                      {ocmChartData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
            {changeCategoryChartData.length > 0 && (
              <div style={styles.chartCard}>
                <h4 style={styles.chartTitle}>Change Category Distribution</h4>
                <p style={styles.chartSubtitle}>Count of items by Change Category</p>
                <ResponsiveContainer width="100%" height={400}>
                  <BarChart data={changeCategoryChartData} margin={{ top: 20, right: 30, left: 20, bottom:60 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis
                      dataKey="name"
                      angle={-35}
                      textAnchor="end"
                      interval={0}
                      height={120}
                      tick={{ fontSize: 11 }}
                    />
<YAxis
  domain={[
    0,
    Math.ceil(Math.max(...changeCategoryChartData.map(d => d.count)) / 10) * 10
  ]}
  ticks={Array.from(
    {
      length:
        Math.ceil(Math.max(...changeCategoryChartData.map(d => d.count)) / 10) + 1
    },
    (_, i) => i * 10
  )}
  allowDecimals={false}
/>
                    <Tooltip formatter={(value) => [`${value} items`, 'Count']} />
                    <Bar dataKey="count" fill="#FD5108" radius={[4, 4, 0, 0]}>
                      {changeCategoryChartData.map((entry, index) => (
                        <Cell key={`cell-${index}`} fill={CHART_COLORS[index % CHART_COLORS.length]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        </section>
      )}
    </div>
  );
}

// Chart colors
const CHART_COLORS = ['#FD5108', '#FF7F3E', '#FFB347', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'];

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
          Preview & download
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
  title = 'Merged preview',
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
              {downloading ? 'Downloading…' : 'Download XLSX'}
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
              {pageRows.map((row, idx) => (
                <tr key={idx}>
                  {row.map((cell, cIdx) => {
                    const cellStr = (cell || '').toString();
                    // Check if cell contains URLs (could be multiple, newline or comma separated)
                    const urls = cellStr.split(/[\n,]/).map(s => s.trim()).filter(s =>
                      s.startsWith('http://') || s.startsWith('https://')
                    );

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
                                Open App {urls.length > 1 ? linkIdx + 1 : ''}
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
              ))}
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
    background: 'rgba(255, 245, 237, 0.95)',
    backdropFilter: 'blur(4px)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: 24,
    boxSizing: 'border-box',
  },
  previewCard: {
    background: '#F7F3EF',
    borderRadius: 16,
    padding: 20,
    width: '90%',
    maxWidth: 1100,
    maxHeight: '90vh',
    overflow: 'hidden',
    display: 'flex',
    flexDirection: 'column',
    gap: 12,
    border: '1px solid #FFCDA8',
    boxShadow: '0 12px 32px rgba(0,0,0,0.12)',
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
    maxHeight: '60vh',
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
};

