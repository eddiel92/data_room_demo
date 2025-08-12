
// Finance Data Room Copilot — self-contained demo logic
(function(){
  // Mock project data
  const project = {
    name: "Project Atlas — Series B Acquisition",
    files: [
      { id: "f1", path: "Finance/Model_v3.xlsx", type: "xlsx", tags:["Finance","Model"], symbols:[
        { kind:"Sheet", name:"P&L" }, { kind:"Sheet", name:"Bridge" },
        { kind:"NamedRange", name:"EBITDA_2022", loc:"P&L!C12" },
        { kind:"NamedRange", name:"EBITDA_2023", loc:"P&L!D12" },
        { kind:"Table", name:"Bridge FY22→FY23", loc:"Bridge!B4:G18" }
      ]},
      { id: "f2", path: "Finance/Historicals_2019-2024.xlsx", type: "xlsx", tags:["Finance"], symbols:[
        { kind:"Sheet", name:"P&L (hist)" }, { kind:"Sheet", name:"Balance" }, { kind:"Sheet", name:"CashFlow" }
      ]},
      { id: "f3", path: "Legal/Credit_Agreement.pdf", type: "pdf", tags:["Legal","Covenants"], symbols:[
        { kind:"Clause", name:"Covenant 4.2: Max Net Leverage", loc:"p112" },
        { kind:"Clause", name:"Covenant 5.1: Min Liquidity", loc:"p118" },
        { kind:"DefinedTerm", name:"EBITDA", loc:"p16" }
      ]},
      { id: "f4", path: "Legal/Shareholders_Agreement_v2.pdf", type: "pdf", tags:["Legal"], symbols:[
        { kind:"Clause", name:"Transfer Restrictions", loc:"p37" }
      ]},
      { id: "f5", path: "Market/Industry_Report_2024.pdf", type: "pdf", tags:["Market"], symbols:[
        { kind:"Figure", name:"TAM/SAM/SOM", loc:"p12" }
      ]},
      { id: "f6", path: "Technical/IP/Patent_Portfolio.xlsx", type: "xlsx", tags:["Technical","IP"], symbols:[
        { kind:"Sheet", name:"Claims Map" }
      ]},
      { id: "f7", path: "DD Reports/Customer_Cohorts.docx", type: "docx", tags:["DD Report"], symbols:[]},
      { id: "f8", path: "Term Sheets/Proposed_Terms.pdf", type: "pdf", tags:["Term Sheet"], symbols:[
        { kind:"Term", name:"Purchase Price", loc:"p2" }
      ]}
    ],
    im: {
      sections: [
        { id:"sec_market", title:"Market & Competitive Landscape" },
        { id:"sec_financials", title:"Financial Analysis" },
        { id:"sec_risks", title:"Risks & Mitigations" }
      ],
      content: {
        sec_market: {
          v1: `# Market & Competitive Landscape

The target participates in a resilient mid-market segment with an estimated TAM of $3.4B growing ~14% CAGR. Share gains are driven by superior workflow coverage and a differentiated data network effect. Competitive intensity is moderate; two scaled incumbents and a long tail of vertical-specific tools. Pricing power is supported by contract sticky-ness (net revenue retention >115%).

**Top Risks.** New entrants building AI-native alternatives and the potential for hyperscaler bundling in the enterprise. Early signals indicate switching costs remain high due to integration depth and trained user behavior. [Industry_Report_2024 p12]{cite:Market/Industry_Report_2024.pdf#p12}`,
          v2: `# Market & Competitive Landscape

The target operates in an expanding mid-market with a TAM of $3.8B growing ~16% CAGR. Share gains are sustained by end-to-end workflow coverage and a proprietary data moat. Competition remains moderate; two scaled incumbents plus niche specialists. Pricing power persists (NRR >118%) with multi-year contracts.

**Top Risks.** AI-native challengers and hyperscaler feature bundling. However, switching costs remain high due to integration depth and trained user behavior. [Industry_Report_2024 p12]{cite:Market/Industry_Report_2024.pdf#p12}`
        },
        sec_financials: {
          v1: `# Financial Analysis

FY2022 EBITDA of $24.1M bridges to FY2023 EBITDA of $28.9M, led by price/mix and opex discipline, partially offset by COGS inflation. See table below. [Model_v3.xlsx Bridge!B4:G18]{cite:Finance/Model_v3.xlsx#Bridge!B4:G18}

**Covenants.** The strictest debt covenant is Maximum Net Leverage ≤ 3.5x (tested quarterly), followed by Minimum Liquidity ≥ $10.0M. [Credit_Agreement.pdf p112]{cite:Legal/Credit_Agreement.pdf#p112}`,
          v2: `# Financial Analysis

FY2022 EBITDA of $24.3M bridges to FY2023 EBITDA of $29.1M, with price/mix uplift and opex control offsetting residual COGS inflation. [Model_v3.xlsx Bridge!B4:G18]{cite:Finance/Model_v3.xlsx#Bridge!B4:G18}

**Covenants.** Maximum Net Leverage ≤ 3.5x; Minimum Liquidity ≥ $10.0M (quarterly). [Credit_Agreement.pdf p112]{cite:Legal/Credit_Agreement.pdf#p112}`
        },
        sec_risks: {
          v1: `# Risks & Mitigations

• Customer concentration (top 10 ≈ 42% of revenue) — mitigate via enterprise expansion and SLG pipeline.  
• Model brittleness under mix volatility — scenario testing added to the model.  
• Data residency and KYC/AML obligations for international growth — policy controls in place.  
[Customer_Cohorts.docx]{cite:DD Reports/Customer_Cohorts.docx}`,
          v2: `# Risks & Mitigations

• Customer concentration (top 10 ≈ 38% of revenue) — enterprise expansion progressing; SLG pipeline expanding.  
• Mix-driven volatility — added sensitivity / scenario testing.  
• Data residency & KYC/AML for int'l growth — policy controls & roadmap.  
[Customer_Cohorts.docx]{cite:DD Reports/Customer_Cohorts.docx}`
        }
      }
    }
  };

  // Utility: escape HTML
  const esc = (s) => String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

  // State
  let paletteMode = "command"; // 'command' or 'open'
  let currentSection = project.im.sections[0].id;
  let redlineMode = false;

  // Elements
  const elRepo = document.getElementById('repoTree');
  const elOutline = document.getElementById('outline');
  const elSectionSelect = document.getElementById('sectionSelect');
  const elDoc = document.getElementById('canvasDoc');
  const elContext = document.getElementById('contextContent');
  const elChat = document.getElementById('chatBox');
  const elChatInput = document.getElementById('chatInput');

  // Init
  function init(){
    // Populate section select
    project.im.sections.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s.id; opt.textContent = s.title;
      elSectionSelect.appendChild(opt);
    });
    elSectionSelect.value = currentSection;
    renderSection();
    renderRepoTree();
    renderOutline();
    renderContext(null);
    addChatMsg("system", "Project loaded. Answers are grounded to sources; use /show sources to render snapshots.");

    // Wire buttons
    document.getElementById('btnRedline').addEventListener('click', toggleRedline);
    document.getElementById('btnInsertTable').addEventListener('click', insertDemoTable);
    document.getElementById('btnInsertCitation').addEventListener('click', insertCitationTag);
    document.getElementById('btnExport').addEventListener('click', exportSection);
    document.getElementById('btnTheme').addEventListener('click', toggleTheme);
    document.getElementById('btnCommand').addEventListener('click', () => openPalette("command"));
    document.getElementById('btnQuickOpen').addEventListener('click', () => openPalette("open"));
    document.getElementById('btnNewProject').addEventListener('click', simulateIngest);

    // Chat
    document.getElementById('chatSend').addEventListener('click', handleChatSend);
    elChatInput.addEventListener('keydown', e => { if(e.key === 'Enter'){ handleChatSend(); } });

    // Section change
    elSectionSelect.addEventListener('change', e => {
      currentSection = e.target.value; redlineMode = false; renderSection();
    });

    // Key bindings
    document.addEventListener('keydown', (e)=>{
      const mod = e.metaKey || e.ctrlKey;
      if(mod && e.key.toLowerCase() === 'k'){ e.preventDefault(); openPalette("command"); }
      if(mod && e.key.toLowerCase() === 'p'){ e.preventDefault(); openPalette("open"); }
      if(mod && e.key.toLowerCase() === 'j'){ e.preventDefault(); toggleTheme(); }
      if(mod && e.key.toLowerCase() === 'n'){ e.preventDefault(); simulateIngest(); }
    });

    // Click on citations (Go to Definition / Peek)
    elDoc.addEventListener('click', (e)=>{
      const tgt = e.target.closest('.citation');
      if(tgt){
        const loc = tgt.getAttribute('data-loc');
        openPeekFor(loc);
      }
    });

    // Command palette interactions
    const pal = document.getElementById('paletteModal');
    pal.addEventListener('keydown', e => {
      if(e.key === 'Escape'){ closePalette(); }
    });
  }

  // Renderers
  function renderRepoTree(){
    elRepo.innerHTML = '';
    const grouped = {};
    for(const f of project.files){
      const dir = f.path.split('/')[0];
      if(!grouped[dir]) grouped[dir] = [];
      grouped[dir].push(f);
    }
    Object.entries(grouped).forEach(([dir, files])=>{
      const header = document.createElement('div');
      header.className = 'node';
      header.innerHTML = `<span class="chev">▾</span><span class="name"><strong>${esc(dir)}</strong></span>`;
      elRepo.appendChild(header);
      files.forEach(f=>{
        const n = document.createElement('div');
        n.className = 'node';
        n.setAttribute('role','treeitem');
        n.innerHTML = `<span class="chev">•</span><span class="name">${esc(f.path.split('/')[1])}</span>`+
                      `<span class="tag">${f.tags.join(', ')}</span>`;
        n.addEventListener('click', ()=> selectFile(f));
        elRepo.appendChild(n);
      });
    });
  }

  function renderOutline(){
    elOutline.innerHTML = '';
    const syms = project.files.flatMap(f=> (f.symbols||[]).map(s=> ({...s, file:f})) );
    syms.forEach(sym => {
      const row = document.createElement('div');
      row.className = 'sym';
      row.innerHTML = `<div class="kind">${esc(sym.kind)}</div>`+
                      `<div class="what">${esc(sym.name)}</div>`+
                      `<div class="loc">${esc(sym.file.path)}${sym.loc?(' — '+esc(sym.loc)):""}</div>`;
      row.addEventListener('click', ()=> openPeekFor(`${sym.file.path}#${sym.loc||''}`));
      elOutline.appendChild(row);
    });
  }

  function renderSection(){
    const content = project.im.content[currentSection];
    const v = redlineMode ? diffMarkup(content.v1, content.v2) : content.v2;
    elDoc.innerHTML = renderCitations(inlineMarkdown(v));
    renderContext({section: currentSection});
  }

  function renderContext(sel){
    const kv = [
      ["Project", project.name],
      ["Section", project.im.sections.find(s=>s.id===currentSection).title],
      ["Routing Policy", "Local LLM (demo)"],
      ["Citations", "Live page/cell references"],
      ["Security", "Single-tenant | RBAC | Audit"],
    ];
    elContext.innerHTML = kv.map(([k,v])=> `<div class="row"><div class="key">${esc(k)}</div><div class="val">${esc(v)}</div></div>`).join('');
  }

  // Simple Markdown-ish renderer for # headers + **bold** + code spans + line breaks
  function inlineMarkdown(s){
    let out = s.replace(/^# (.*)$/gm, '<h2>$1</h2>');
    out = out.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    out = out.replace(/`(.*?)`/g, '<code>$1</code>');
    out = out.replace(/\n/g, '<br/>');
    return out;
  }

  // Render [Text]{cite:file#loc} to clickable citations
  function renderCitations(html){
    return html.replace(/\[([^\]]+)\]\{cite:([^}]+)\}/g, (m, label, loc)=> {
      return `${esc(label)} <span class="citation" title="Go to Definition / Peek" data-loc="${esc(loc)}">[${esc(loc.split('#')[1]||'cite')}]</span>`;
    });
  }

  // Redline (very simple token diff)
  function diffMarkup(a, b){
    const ta = tokenize(a), tb = tokenize(b);
    // LCS dynamic programming (simple)
    const dp = Array(ta.length+1).fill(null).map(()=>Array(tb.length+1).fill(0));
    for(let i=ta.length-1;i>=0;i--){
      for(let j=tb.length-1;j>=0;j--){
        dp[i][j] = ta[i] === tb[j] ? dp[i+1][j+1]+1 : Math.max(dp[i+1][j], dp[i][j+1]);
      }
    }
    let i=0,j=0,res=[];
    while(i<ta.length && j<tb.length){
      if(ta[i]===tb[j]){ res.push(ta[i]); i++; j++; }
      else if(dp[i+1][j]>=dp[i][j+1]){ res.push(`<span class="deleted">${esc(ta[i])}</span>`); i++; }
      else { res.push(`<span class="added">${esc(tb[j])}</span>`); j++; }
    }
    while(i<ta.length){ res.push(`<span class="deleted">${esc(ta[i++])}</span>`); }
    while(j<tb.length){ res.push(`<span class="added">${esc(tb[j++])}</span>`); }
    return res.join('').replace(/\n/g,'\n');
  }
  function tokenize(s){ return s.replace(/\s+/g,' ').split(' ').map(t=>t+' '); }

  // Insert demo table
  function insertDemoTable(){
    const tbl = `<table>
      <thead><tr><th>Bridge Driver</th><th>Impact ($M)</th></tr></thead>
      <tbody>
        <tr><td>Price</td><td>+3.1</td></tr>
        <tr><td>Volume</td><td>+1.4</td></tr>
        <tr><td>Mix</td><td>+0.8</td></tr>
        <tr><td>COGS</td><td>-0.9</td></tr>
        <tr><td>Opex</td><td>+1.6</td></tr>
      </tbody>
    </table> <span class="citation" data-loc="Finance/Model_v3.xlsx#Bridge!B4:G18">[Bridge!B4:G18]</span>`;
    insertHTMLAtCursor(tbl);
  }

  function insertCitationTag(){
    insertHTMLAtCursor(` <span class="citation" data-loc="Legal/Credit_Agreement.pdf#p112">[p112]</span> `);
  }

  function insertHTMLAtCursor(html){
    elDoc.focus();
    document.execCommand('insertHTML', false, html);
  }

  // Palette
  function openPalette(mode){
    paletteMode = mode;
    const modal = document.getElementById('paletteModal');
    modal.setAttribute('aria-hidden','false');
    const input = document.getElementById('paletteInput');
    input.value='';
    renderPaletteResults('');
    input.focus();
    input.oninput = (e)=> renderPaletteResults(e.target.value);
    input.onkeydown = (e)=> {
      if(e.key==='Escape'){ closePalette(); }
      if(e.key==='Enter'){ activatePaletteSelection(); }
      if(e.key==='ArrowDown' || e.key==='ArrowUp'){ movePaletteSelection(e.key==='ArrowDown'?1:-1); e.preventDefault();}
    };
  }
  function closePalette(){
    document.getElementById('paletteModal').setAttribute('aria-hidden','true');
  }
  function paletteItems(query){
    if(paletteMode==='command'){
      const cmds = [
        { id:'cmd_extract', label:'/extract table — Insert table from Model', meta:'Action' },
        { id:'cmd_explain', label:'/explain formula — Explain EBITDA formula', meta:'Action' },
        { id:'cmd_sources', label:'/show sources — Toggle source snapshots', meta:'Action' },
        { id:'cmd_findrefs', label:'Find References: "EBITDA"', meta:'Search' }
      ];
      return cmds.filter(c=> c.label.toLowerCase().includes(query.toLowerCase()));
    } else {
      const files = project.files.flatMap(f=>{
        const arr = [{ id: 'open:'+f.path, label: f.path, meta: f.type.toUpperCase() }];
        (f.symbols||[]).forEach(s => arr.push({ id:'peek:'+f.path+'#'+(s.loc||''), label: s.name, meta: s.kind }));
        return arr;
      });
      return files.filter(x=> x.label.toLowerCase().includes(query.toLowerCase()));
    }
  }
  function renderPaletteResults(query){
    const items = paletteItems(query);
    const list = document.getElementById('paletteResults');
    list.innerHTML = items.map((it,i)=> `<div class="item ${i===0?'active':''}" data-id="${esc(it.id)}"><div class="meta">${esc(it.meta)}</div><div>${esc(it.label)}</div></div>`).join('');
  }
  function movePaletteSelection(dir){
    const list = document.getElementById('paletteResults');
    const items = Array.from(list.querySelectorAll('.item'));
    let idx = items.findIndex(el=> el.classList.contains('active'));
    idx = Math.max(0, Math.min(items.length-1, idx+dir));
    items.forEach(el=> el.classList.remove('active'));
    if(items[idx]) items[idx].classList.add('active');
  }
  function activatePaletteSelection(){
    const list = document.getElementById('paletteResults');
    const active = list.querySelector('.item.active');
    if(!active) return;
    const id = active.getAttribute('data-id');
    if(id==='cmd_extract'){ insertDemoTable(); }
    else if(id==='cmd_explain'){ explainFormula(); }
    else if(id==='cmd_sources'){ toggleSourcesInChat(); }
    else if(id==='cmd_findrefs'){ findReferences('EBITDA'); }
    else if(id.startsWith('open:')){ addChatMsg('system', `Opened ${id.slice(5)} (demo)`); }
    else if(id.startsWith('peek:')){ openPeekFor(id.slice(5)); }
    closePalette();
  }

  // Peek preview
  function openPeekFor(loc){
    const [file, anchor] = loc.split('#');
    const title = `${file}${anchor?(' — '+anchor):''}`;
    document.getElementById('peekTitle').textContent = title;
    const body = document.getElementById('peekBody');
    let img = null;
    if(file.includes('Credit_Agreement') && anchor && anchor.includes('p112')){
      img = 'assets/snap_credit_agreement_p112.svg';
    } else if(file.includes('Model_v3.xlsx') && anchor && anchor.includes('Bridge')){
      img = 'assets/snap_model_bridge_B4_G18.svg';
    } else if(file.includes('Model_v3.xlsx') && anchor && anchor.includes('P&L!C12')){
      img = 'assets/snap_model_formula_C12.svg';
    }
    if(img){
      body.innerHTML = `<img alt="Snapshot" src="${img}"/>`;
    } else {
      body.innerHTML = `<div style="padding:16px;color:#9bb4d1;">No snapshot cached for <code>${esc(loc)}</code>. (Demo)</div>`;
    }
    const modal = document.getElementById('peekModal');
    modal.setAttribute('aria-hidden','false');
    document.getElementById('btnPeekClose').onclick = ()=> modal.setAttribute('aria-hidden','true');
  }

  // Chat
  function addChatMsg(role, text, cites){
    const div = document.createElement('div');
    div.className = 'msg';
    div.innerHTML = `<div class="role">${esc(role)}</div><div class="bubble">${esc(text)}${
      cites?`<div style="margin-top:6px;">Sources: ${cites.map(c=> `<span class="cite" data-loc="${esc(c)}">${esc(c.split('#')[1]||c)}</span>`).join(' • ')}</div>`:''
    }</div>`;
    elChat.appendChild(div);
    elChat.scrollTop = elChat.scrollHeight;
    // clicks on cites
    div.querySelectorAll('.cite').forEach(el => el.addEventListener('click', ()=> openPeekFor(el.getAttribute('data-loc')) ));
  }

  function handleChatSend(){
    const q = elChatInput.value.trim();
    if(!q) return;
    addChatMsg('you', q);
    elChatInput.value = '';
    // Recognize demo queries / slash commands
    if(q.toLowerCase().includes('fy2022 ebida') || q.toLowerCase().includes('fy2022 ebitda') || q.toLowerCase().includes('bridge')){
      addChatMsg('copilot', 'FY2022 EBITDA of $24.3M bridges to FY2023 EBITDA of $29.1M. Drivers: price, volume, mix, COGS, opex.', [
        'Finance/Model_v3.xlsx#Bridge!B4:G18'
      ]);
    } else if(q.toLowerCase().includes('covenant') || q.toLowerCase().includes('strictest')){
      addChatMsg('copilot', 'Strictest covenant: Maximum Net Leverage ≤ 3.5x; also Minimum Liquidity ≥ $10.0M (quarterly tests).', [
        'Legal/Credit_Agreement.pdf#p112'
      ]);
    } else if(q.startsWith('/extract')){
      insertDemoTable();
      addChatMsg('copilot', 'Inserted table from Model_v3.xlsx Bridge range.', [
        'Finance/Model_v3.xlsx#Bridge!B4:G18'
      ]);
    } else if(q.startsWith('/explain')){
      explainFormula();
    } else if(q.startsWith('/show sources')){
      toggleSourcesInChat();
    } else if(q.startsWith('/insert citations')){
      insertCitationTag();
      addChatMsg('copilot', 'Inserted citation tag at cursor.', []);
    } else {
      addChatMsg('copilot', 'This demo answers with grounded citations. Try: "Show FY2022 EBITDA and bridge to FY2023" or "/explain formula".');
    }
  }

  function explainFormula(){
    addChatMsg('copilot', 'EBITDA = SUM(C5:C9) - SUM(C10:C11). Non-cash add-backs excluded in values-only mode.', [
      'Finance/Model_v3.xlsx#P&L!C12'
    ]);
    openPeekFor('Finance/Model_v3.xlsx#P&L!C12');
  }

  let showSources = true;
  function toggleSourcesInChat(){
    showSources = !showSources;
    // In this demo, sources are always shown per message; we simulate toggle by adding an info message
    addChatMsg('system', 'Show Sources is ' + (showSources? 'ON':'OFF') + ' (snapshots available for key cites).');
  }

  // Find References (simple demo)
  function findReferences(term){
    const refs = [
      { file:'Legal/Credit_Agreement.pdf', where:'p16', snippet:'"EBITDA" means Earnings Before Interest, Taxes, Depreciation and Amortization…' },
      { file:'Finance/Model_v3.xlsx', where:"P&L!C12", snippet:'Named range EBITDA_2022 refers to P&L!C12' }
    ];
    const lines = refs.map(r=> `• ${r.file} — ${r.where}: ${r.snippet}`).join('\n');
    addChatMsg('copilot', `References for ${term}:\n${lines}`, refs.map(r=> `${r.file}#${r.where}`));
  }

  // Redline toggle
  function toggleRedline(){
    redlineMode = !redlineMode; renderSection();
  }

  // Export current section as .md
  function exportSection(){
    const content = project.im.content[currentSection].v2;
    const blob = new Blob([content], {type:'text/markdown'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    const name = project.im.sections.find(s=>s.id===currentSection).title.replace(/\s+/g,'_').toLowerCase();
    a.download = `${name}.md`;
    a.click();
    URL.revokeObjectURL(a.href);
  }

  // File selection (updates context)
  function selectFile(f){
    renderContext({file: f.path});
    addChatMsg('system', `Selected ${f.path}. Use Peek to preview cited pages/ranges.`);
  }

  // Theme toggle
  function toggleTheme(){
    document.body.classList.toggle('light');
  }

  // Ingest simulation
  function simulateIngest(){
    const modal = document.getElementById('ingestModal');
    modal.setAttribute('aria-hidden','false');
    const bar = document.getElementById('ingestBar');
    const pct = document.getElementById('ingestPct');
    const log = document.getElementById('ingestLog');
    log.innerHTML = '';
    const steps = [
      "Unzipping archive…",
      "Detecting nested ZIPs…",
      "Unlocking passworded PDFs…",
      "OCR (layout-aware, CJK)…",
      "Parsing Excel values + formula ASTs…",
      "Dedup & idempotency checks…",
      "Building lexical/semantic/structure indices…",
      "Classifying (Legal/Finance/Tech/Market/Risk/ESG)…",
      "Mapping to IM template & ranking by materiality…",
      "Snapshotting cited pages & ranges…",
      "Ready."
    ];
    let i=0;
    const timer = setInterval(()=>{
      if(i<steps.length){
        const line = document.createElement('div');
        line.innerHTML = `<span class="${i<steps.length-1?'ok':''}">• ${steps[i]}</span>`;
        log.appendChild(line);
        const p = Math.round((i+1)/steps.length*100);
        bar.style.width = p+'%'; pct.textContent = p+'%';
        i++;
      } else { clearInterval(timer); }
    }, 500);
    document.getElementById('btnIngestClose').onclick = ()=> modal.setAttribute('aria-hidden','true');
  }

  // Boot
  init();
})();
