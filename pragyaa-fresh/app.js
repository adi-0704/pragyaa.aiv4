/* ═══════════════════════════════════════════════
   Pragyaa.AI — App Logic (Client-Side Excel Processing)
   Uses SheetJS (xlsx) for browser-based Excel parsing
   ═══════════════════════════════════════════════ */

// Load SheetJS from CDN
const script = document.createElement('script');
script.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
document.head.appendChild(script);

// ─── State ────────────────────────────────────
let state = {
  file: null,
  audioFile: null,
  rawData: null,
  analysis: null,
  deltas: null,
  optimizedPrompt: '',
  history: [],
  engineMode: 'gemini',
};

// ─── DOM References ───────────────────────────
const $ = id => document.getElementById(id);
const dropzone = $('dropzone');
const fileInput = $('fileInput');
const fileInfo = $('fileInfo');
const fileName = $('fileName');
const removeFile = $('removeFile');
const analyzeBtn = $('analyzeBtn');
const resultsSection = $('resultsSection');
const promptSection = $('promptSection');
const testSection = $('testSection');
const historySection = $('historySection');
const loadingOverlay = $('loadingOverlay');
const toast = $('toast');
const engineModeSelect = $('engineMode');

// Audio References (Replaced by Transcript Test)
const runTestBtn = $('runTestBtn');
const testResults = $('testResults');
const testTranscriptInput = $('testTranscriptInput');
const oldAuditResult = $('oldAuditResult');
const newAuditResult = $('newAuditResult');

// Toolbar References
const toggleBtns = document.querySelectorAll('.toggle-btn');
const promptView = $('promptView');
const optimizedPromptEl = $('optimizedPrompt');
const changesView = $('changesView');
const reportView = $('reportView');
const promptVersionBadge = $('promptVersionBadge');
const transcriptFileInput = $('transcriptFileInput');

// ─── Event Listeners ──────────────────────────
engineModeSelect.addEventListener('change', (e) => {
  state.engineMode = e.target.value;
  if (state.engineMode === 'vertex') {
    showToast('🚀 Backend Engine selected (Cloud ID not required)');
  }
});

// ─── Upload Handlers ──────────────────────────
// Fix: Use a single robust click handler for the dropzone
dropzone.onclick = (e) => {
  console.log('Dropzone clicked');
  fileInput.click();
};

fileInput.onchange = (e) => {
  console.log('File input changed');
  if (e.target.files.length > 0) {
    handleFile(e.target.files[0]);
  }
};

dropzone.ondragover = (e) => {
  e.preventDefault();
  dropzone.classList.add('drag-over');
};

dropzone.ondragleave = () => {
  dropzone.classList.remove('drag-over');
};

dropzone.ondrop = (e) => {
  e.preventDefault();
  dropzone.classList.remove('drag-over');
  if (e.dataTransfer.files.length > 0) {
    handleFile(e.dataTransfer.files[0]);
  }
};

function handleFile(file) {
  if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
    showToast('❌ Please upload an Excel file (.xlsx)');
    return;
  }
  state.file = file;
  fileName.textContent = file.name;
  fileInfo.style.display = 'flex';
  analyzeBtn.disabled = false;
  showToast('📊 File ready for analysis');
  console.log('File handled:', file.name);
}

removeFile.onclick = (e) => {
  e.stopPropagation();
  state.file = null;
  fileInput.value = '';
  fileInfo.style.display = 'none';
  analyzeBtn.disabled = true;
};

// ─── Transcript File Upload Handler ───────────
if (transcriptFileInput) {
  transcriptFileInput.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      testTranscriptInput.value = event.target.result;
      showToast('📄 Transcript loaded successfully');
    };
    reader.readAsText(file);
  };
}

// ─── View Toggle Logic ────────────────────────
toggleBtns.forEach(btn => {
  btn.addEventListener('click', (e) => {
    toggleBtns.forEach(b => b.classList.remove('active'));
    e.target.classList.add('active');
    
    optimizedPromptEl.style.display = 'none';
    changesView.style.display = 'none';
    reportView.style.display = 'none';
    
    const view = e.target.getAttribute('data-view');
    if (view === 'optimized') optimizedPromptEl.style.display = 'block';
    if (view === 'changes') changesView.style.display = 'block';
    if (view === 'report') reportView.style.display = 'block';
  });
});

// Audio handlers removed — text transcript pasting is now used directly.

// ─── Analysis Pipeline ────────────────────────
analyzeBtn.onclick = async () => {
  if (!state.file) return;
  showLoading('Processing...');
  
  try {
    if (state.engineMode === 'vertex') {
      updateLoader('Calling Backend API (Vertex AI)...');
      await runBackendAnalysis();
    } else {
      updateLoader('Analyzing locally...');
      await runLocalAnalysis();
    }
  } catch (err) {
    hideLoading();
    showToast('❌ Error: ' + err.message);
    console.error(err);
  }
};

async function safeJson(response) {
  const text = await response.text();
  try {
    return JSON.parse(text);
  } catch (err) {
    console.error('Failed to parse JSON:', text.substring(0, 500));
    throw new Error(`Invalid JSON response: ${text.substring(0, 100)}...`);
  }
}

async function runBackendAnalysis() {
  const reader = new FileReader();
  const fileBase64 = await new Promise((resolve) => {
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.readAsDataURL(state.file);
  });

  const response = await fetch('/api/analyze', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      file_content: fileBase64,
      current_prompt: $('currentPrompt').value,
      generate_prompt: true
    })
  });

  const result = await safeJson(response);
  if (!response.ok) {
    throw new Error(result.error || `Backend error (${response.status})`);
  }

  state.analysis = result.analysis;
  state.deltas = result.deltas;
  
  if (result.optimized_prompt) {
    state.optimizedPrompt = result.optimized_prompt;
    state.promptSource = 'ai';
  } else {
    throw new Error('AI FAILED: ' + (result.vertex_status || 'Empty response'));
  }

  finalizeAnalysis();
}

async function runLocalAnalysis() {
    const data = await readExcel(state.file);
    state.rawData = data;
    updateLoader('Analyzing discrepancies...');
    
    state.analysis = analyzeRootCauses(data);
    updateLoader('Generating prompt deltas...');
    state.deltas = generateDeltas(state.analysis);
    
    const currentPrompt = $('currentPrompt').value;
    
    try {
      updateLoader('🤖 Calling Vertex AI...');
      state.optimizedPrompt = await generatePromptWithVertex(state.analysis, state.deltas, currentPrompt, 2);
      state.promptSource = 'ai';
    } catch (apiError) {
      console.error('Frontend Vertex AI failed, trying backend fallback...', apiError);
      showToast('⚠️ Direct AI call failed, trying backend...');
      try {
        updateLoader('🤖 Trying Backend Vertex AI...');
        state.optimizedPrompt = await generatePromptViaBackend(state.analysis, state.deltas, currentPrompt);
        state.promptSource = 'ai';
      } catch (backendError) {
        console.error('Backend fallback failed:', backendError);
        throw new Error('AI FAILED: ' + backendError.message);
      }
    }
    
    finalizeAnalysis();
}

async function generatePromptViaBackend(analysis, deltas, currentPrompt) {
  if (!state.file) throw new Error('No file available for backend fallback');
  
  const reader = new FileReader();
  const fileBase64 = await new Promise((resolve) => {
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.readAsDataURL(state.file);
  });

  const response = await fetch('/api/analyze', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      file_content: fileBase64,
      current_prompt: currentPrompt,
      generate_prompt: true
    })
  });

  const result = await safeJson(response);
  if (!response.ok) throw new Error(result.error || `Backend error (${response.status})`);
  if (!result.optimized_prompt) throw new Error(result.vertex_status || 'Empty response');
  
  return result.optimized_prompt;
}

function finalizeAnalysis() {
    state.history.push({
      timestamp: new Date().toISOString(),
      file: state.file.name,
      cases: state.analysis.summary.total,
      agreement: state.analysis.summary.agreementRate,
      falseReworks: state.analysis.summary.falseReworkCount,
      source: state.promptSource,
    });

    renderResults();
    renderPrompt();
    renderHistory();
    hideLoading();
    
    showToast(`✅ Analysis complete — 🤖 AI prompt ready`);
    
    resultsSection.style.display = 'block';
    promptSection.style.display = 'block';
    testSection.style.display = 'block';
    historySection.style.display = 'block';
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── Live Test Pipeline ───────────────────────
runTestBtn.onclick = async () => {
  const transcriptText = testTranscriptInput.value.trim();
  const currentPrompt = $('currentPrompt').value.trim() || 'You are an auditor. Audit the call.';
  
  if (!transcriptText || !state.optimizedPrompt) {
    showToast('⚠️ Please generate an optimized prompt and paste a transcript first.');
    return;
  }
  
  showLoading('Running Variance Test...');
  
  try {
    // Helper to format AI response
    const formatResponse = (text) => {
      let cleaned = text.replace(/```json\s*/, '').replace(/```\s*$/, '').trim();
      try {
        const parsed = JSON.parse(cleaned);
        return `<pre class="json-output">${JSON.stringify(parsed, null, 2)}</pre>`;
      } catch (e) {
        return `<pre class="json-output">${cleaned}</pre>`;
      }
    };

    updateLoader('⚖️ Auditing with OLD Prompt...');
    const oldEvalResponse = await fetch('/api/vertex', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt: `${currentPrompt}\n\n[TRANSCRIPT TO AUDIT]:\n${transcriptText}`,
        model: "gemini-2.5-flash-lite"
      })
    });
    
    if (!oldEvalResponse.ok) throw new Error('Old Audit failed');
    const oldEvalResult = await oldEvalResponse.json();
    const oldRawText = oldEvalResult.text || oldEvalResult.response || JSON.stringify(oldEvalResult);
    oldAuditResult.innerHTML = formatResponse(oldRawText);

    updateLoader('⚖️ Auditing with NEW Optimized Prompt...');
    const newEvalResponse = await fetch('/api/vertex', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt: `${state.optimizedPrompt}\n\n[TRANSCRIPT TO AUDIT]:\n${transcriptText}`,
        model: "gemini-2.5-flash-lite"
      })
    });
    
    if (!newEvalResponse.ok) throw new Error('New Audit failed');
    const newEvalResult = await newEvalResponse.json();
    const newRawText = newEvalResult.text || newEvalResult.response || JSON.stringify(newEvalResult);
    newAuditResult.innerHTML = formatResponse(newRawText);
    
    testResults.style.display = 'grid';
    testResults.scrollIntoView({ behavior: 'smooth' });
    hideLoading();
    showToast('✅ Variance Test Complete');
    
  } catch (err) {
    hideLoading();
    showToast('❌ Test Error: ' + err.message);
    console.error(err);
  }
};

// ─── Excel Reader ─────────────────────────────
async function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const sheetName = wb.SheetNames.includes('Raw Data') ? 'Raw Data' : wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        resolve(json);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

// ─── Root Cause Analyzer ──────────────────────
function analyzeRootCauses(data) {
  const norm = v => String(v || '').trim().toLowerCase();
  const total = data.length;
  const cols = Object.keys(data[0]);

  // ── Detect Process ────────────────────────────
  const hasCallStatus = cols.some(c => c.includes('Call Status AI'));
  const hasAccurateDisposition = cols.some(c => c === 'Accurate Disposition');
  const hasMismatchRemarks = cols.some(c => c === 'Mismatch Remarks');
  const processType = hasCallStatus ? 'cc_dc' : (hasAccurateDisposition ? 'insta' : 'generic');

  let agree = 0, mismatch = [], allRows = data;

  if (processType === 'cc_dc') {
    // CC/DC: compare Call Status AI vs Call Status Verifier
    const aiCol = cols.find(c => c.includes('Call Status AI'));
    const verCol = cols.find(c => c.includes('Call Status Verifier'));
    data.forEach(row => {
      const ai = norm(row[aiCol] || '');
      const ver = norm(row[verCol] || '');
      if (ai === ver) agree++;
      else mismatch.push(row);
    });
  } else if (processType === 'insta') {
    // Insta: use Accurate Disposition Yes/No as the mismatch flag
    const accCol = 'Accurate Disposition';
    data.forEach(row => {
      const v = norm(row[accCol] || '');
      if (v === 'yes') agree++;
      else if (v === 'no') mismatch.push(row);
      else agree++; // treat blanks as correct to avoid inflation
    });
  } else {
    // Generic fallback: compare Disposition vs Finmech Disposition
    const aiCol = cols.find(c => c.toLowerCase() === 'disposition') || cols[0];
    const verCol = cols.find(c => c.toLowerCase() === 'finmech disposition') || cols[1];
    data.forEach(row => {
      const ai = norm(row[aiCol] || '');
      const ver = norm(row[verCol] || '');
      if (ai === ver) agree++;
      else mismatch.push(row);
    });
  }

  // ── Parameter Failure Analysis ─────────────────
  // Find all score columns
  const scoreCols = cols.filter(c => c.includes(' Score') || c.includes(' Met'));
  const binaryCols = cols.filter(c => ['Accurate Disposition','Junk Lead','Complete Information Provided','Professional Behavior','No Disconnection Avoid'].includes(c));
  const paramFailures = {};

  scoreCols.forEach(col => {
    // Get max possible score for this column across all data
    const allVals = allRows.map(r => parseFloat(r[col]) || 0);
    const maxScore = Math.max(...allVals);
    if (maxScore === 0) return; // skip always-zero columns
    
    const fails = mismatch.filter(r => {
      const v = r[col];
      const num = parseFloat(v);
      if (!isNaN(num)) return num === 0;
      return norm(v) === 'no' || norm(v) === '0';
    }).length;
    if (fails > 0) {
      paramFailures[col] = { count: fails, pct: Math.round(fails / (mismatch.length || 1) * 1000) / 10 };
    }
  });

  binaryCols.forEach(col => {
    const fails = mismatch.filter(r => norm(r[col] || '') === 'no' || norm(r[col] || '') === 'incorrect').length;
    if (fails > 0) {
      paramFailures[col] = { count: fails, pct: Math.round(fails / (mismatch.length || 1) * 1000) / 10 };
    }
  });

  // ── Mismatch Remarks Pattern Analysis ──────────
  const remarksCols = cols.filter(c => c.toLowerCase().includes('mismatch') || c.toLowerCase().includes('remark'));
  const remarkPatterns = {};
  mismatch.forEach(row => {
    remarksCols.forEach(rc => {
      const remark = norm(row[rc] || '');
      if (!remark) return;
      // Split on comma and count each segment
      remark.split(',').forEach(segment => {
        const s = segment.trim();
        if (s.length > 2) {
          remarkPatterns[s] = (remarkPatterns[s] || 0) + 1;
        }
      });
    });
  });

  // Sort and keep top patterns
  const sortedRemarks = Object.entries(remarkPatterns)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8)
    .reduce((acc, [k, v]) => ({ ...acc, [k]: v }), {});

  // ── Score computation ──────────────────────────
  const agreementRate = Math.round(agree / (total || 1) * 1000) / 10;
  const mismatchRate = Math.round(mismatch.length / (total || 1) * 1000) / 10;

  return {
    summary: {
      total,
      agree,
      agreementRate,
      mismatchRate,
      aiApprovalRate: agreementRate,
      verApprovalRate: 100 - mismatchRate,
      falseReworkCount: mismatch.length,
      falseApproveCount: 0,
      processType,
    },
    paramFailures,
    consentPatterns: sortedRemarks,
    chargesPatterns: {},
    faReasons: [],
  };
}


function generateDeltas(analysis) {
  const deltas = [];
  const pf = analysis.paramFailures;
  const sorted = Object.entries(pf).sort((a, b) => b[1].pct - a[1].pct);
  sorted.forEach(([param, info]) => {
    if (info.pct < 5) return;
    const severity = info.pct > 50 ? 'CRITICAL' : info.pct > 20 ? 'HIGH' : 'MEDIUM';
    let rootCause = '', fix = '';
    if (param === 'Consent Taken Met') {
      rootCause = 'AI rejects passive Hindi consent ("Okay", "Haan ji", "Theek hai")';
      fix = 'Implement 3-Tier consent: Tier 1 (explicit), Tier 2 (contextual — valid after full pitch), Tier 3 (refusal only)';
    } else if (param === 'Charges Explained Met') {
      rootCause = 'AI penalizes delivery speed instead of factual accuracy';
      fix = 'Switch to content-based evaluation — pass if ₹699+GST stated correctly regardless of pace';
    } else if (param === 'Pitch Pace') {
      rootCause = 'AI pace threshold stricter than human verifier tolerance';
      fix = 'Only fail if customer explicitly asks to repeat or slow down';
    } else if (param === 'Benefits Explained Met') {
      rootCause = 'AI requires too many benefits to be mentioned';
      fix = 'Pass if agent mentions ≥2 core benefits accurately';
    } else if (param === 'Card Variant Met') {
      rootCause = 'AI fails on informal card name variations';
      fix = 'Add fuzzy mapping: "Coral card"/"Updated Coral"/"Coral Visa" → Coral Debit Card';
    } else {
      rootCause = `AI too strict on ${param}`;
      fix = `Align ${param} threshold with human verifier standards`;
    }
    deltas.push({ param, ...info, severity, rootCause, fix });
  });
  return deltas;
}

async function generatePromptWithVertex(analysis, deltas, currentPrompt, maxRetries = 1) {
  const s = analysis.summary;
  const metaPrompt = `You are an expert prompt engineer specializing in compliance audit automation.
I have analyzed ${s.total} audits. AI Approval: ${s.aiApprovalRate}% | Verifier Approval: ${s.verApprovalRate}%
Agreement: ${s.agreementRate}%. TOP FAILURES:
${deltas.map(d => `- ${d.param}: ${d.pct}% failure rate | Fix: ${d.fix}`).join('\n')}

TASK: Write a compliance audit prompt that matches human judgment. Use 3-tier consent and content-based charges (₹699+GST).
Output ONLY the audit prompt content.`;

  const VERTEX_GENERATE_URL = '/api/vertex';
  let lastError;
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch(VERTEX_GENERATE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt: metaPrompt, model: "gemini-2.5-flash-lite" })
      });
      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`Vertex AI error (${response.status}): ${errText.substring(0, 100)}`);
      }
      const result = await response.json();
      let text = result.text || result.response || result.content || result.generated_text || result.result || result.data ||
                 result.candidates?.[0]?.content?.parts?.[0]?.text || 
                 (typeof result === 'string' ? result : null);
      if (typeof text === 'object' && text !== null) text = JSON.stringify(text);
      if (!text && result.candidates?.[0]?.output) text = result.candidates[0].output;
      if (!text) throw new Error('Vertex AI returned empty response');
      return `# AI-GENERATED PROMPT — Vertex AI\n# Generated: ${new Date().toLocaleString()}\n\n${text}`;
    } catch (err) {
      lastError = err;
      if (attempt < maxRetries) await sleep(1500);
    }
  }
  throw lastError;
}

function buildTemplatePrompt(analysis, deltas, currentPrompt) {
  return "/* RECOVERY TEMPLATE */";
}

function renderResults() {
  const s = state.analysis.summary;
  const gap = Math.round((s.verApprovalRate - s.aiApprovalRate) * 10) / 10;
  const processLabel = s.processType === 'cc_dc' ? 'CC/DC Upgrade' : s.processType === 'insta' ? 'Insta (NCA)' : 'Generic';
  
  // Update process detected label if it exists
  const processDetectedEl = document.querySelector('.results-card p');
  if (processDetectedEl) processDetectedEl.innerHTML = `Process Detected: <strong style="color: var(--accent-color)">${processLabel}</strong>`;

  $('statsGrid').innerHTML = [
    { value: s.total, label: 'Total Cases', cls: '' },
    { value: s.agreementRate + '%', label: 'AI Accuracy', cls: s.agreementRate > 70 ? 'success' : 'danger' },
    { value: s.falseReworkCount, label: 'Mismatches', cls: s.falseReworkCount > 0 ? 'danger' : 'success' },
    { value: (isNaN(gap) ? (100 - s.agreementRate) : Math.abs(gap)) + '%', label: 'Error Rate', cls: 'danger' },
  ].map(s => `<div class="stat-card"><div class="stat-value ${s.cls}">${s.value}</div><div class="stat-label">${s.label}</div></div>`).join('');

  const sorted = Object.entries(state.analysis.paramFailures).sort((a, b) => b[1].pct - a[1].pct);
  if (sorted.length === 0) {
    $('failureBars').innerHTML = '<p style="opacity:0.5; text-align:center; padding: 20px;">No parameter failures detected in mismatched rows.</p>';
  } else {
    $('failureBars').innerHTML = sorted.map(([param, info]) => {
      const cls = info.pct > 50 ? 'critical' : info.pct > 20 ? 'high' : 'medium';
      return `<div class="failure-row"><div class="failure-label">${param.replace(' Met', '').replace(' Score', '')}</div><div class="failure-bar-bg"><div class="failure-bar ${cls}" style="width: ${info.pct}%"></div></div><div class="failure-pct">${info.pct}%</div></div>`;
    }).join('');
  }

  const remarksEntries = Object.entries(state.analysis.consentPatterns || {});
  const patternTitle = s.processType === 'insta' ? '📋 Top Mismatch Patterns (from Remarks)' : '🔒 Consent Patterns';
  $('patternsGrid').innerHTML = `<div class="pattern-box" style="grid-column: 1/-1"><h4>${patternTitle}</h4>${
    remarksEntries.length > 0
      ? remarksEntries.map(([k, v]) => `<div class="pattern-item"><span>${k}</span><span class="badge-count">${v}</span></div>`).join('')
      : '<p style="opacity:0.5">No patterns found in mismatch remarks.</p>'
  }</div>`;
}


function renderPrompt() {
  const version = `v1.${state.history.length}`;
  const badge = '<span class="ai-badge">🤖 Vertex AI</span>';
  $('promptSection').querySelector('h2').innerHTML = `Optimized Prompt ${badge} <span class="badge-dot" style="font-size: 0.5em; vertical-align: middle; padding: 2px 6px; border-radius: 4px; background: var(--accent-color); color: #000; margin-left: 10px;">${version}</span>`;
  
  // Format version in prompt header
  const versionHeader = `# AI-GENERATED PROMPT — Vertex AI\n# Version: ${version}\n# Generated: ${new Date().toLocaleString()}\n\n`;
  const cleanPrompt = state.optimizedPrompt.replace(/^# AI-GENERATED.*?\n# Generated:.*?\n\n/m, '');
  $('optimizedPrompt').textContent = versionHeader + cleanPrompt;
  
  // Render Changes View
  if (state.deltas && state.deltas.length > 0) {
    changesView.innerHTML = `<h3>Modifications (Version ${version})</h3>
      <ul style="list-style: none; padding: 0;">
        ${state.deltas.map(d => `
          <li style="margin-bottom: 15px; padding: 15px; background: rgba(255,255,255,0.05); border-left: 4px solid var(--accent-color); border-radius: 4px;">
            <div style="font-weight: 600; color: var(--accent-color); margin-bottom: 5px;">${d.param} (${d.severity})</div>
            <div style="font-size: 0.9em; margin-bottom: 5px;"><strong>Root Cause:</strong> ${d.rootCause}</div>
            <div style="font-size: 0.9em; color: #a5d6a7;"><strong>Fix Applied:</strong> ${d.fix}</div>
          </li>
        `).join('')}
      </ul>`;
  } else {
    changesView.innerHTML = '<p>No specific parameter gaps identified or fixed in this iteration.</p>';
  }
}

function renderHistory() {
  $('timeline').innerHTML = state.history.map((h, i) => `<div class="timeline-item"><div class="timeline-date">${new Date(h.timestamp).toLocaleString()}</div><div class="timeline-content"><strong>Run #${i + 1}:</strong> ${h.file}</div></div>`).join('');
}

function showLoading(text) { loadingOverlay.style.display = 'flex'; $('loaderSub').textContent = text; }
function updateLoader(text) { $('loaderSub').textContent = text; }
function hideLoading() { loadingOverlay.style.display = 'none'; }
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function showToast(msg) {
  toast.textContent = msg;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 3000);
}
