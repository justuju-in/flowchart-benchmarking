/** Standalone bound Apps Script for a COPY of the spreadsheet. */
const FC_HEADERS = ['Run ID','Source row','Problem','Image snapshot','Image SHA256','Prompt SHA256','Model','State','Code URL','Code SHA256','Raw response URL','Model version','Prompt tokens','Output tokens','Thought tokens','Total tokens','API ms','Estimated USD','Pricing date','DOMjudge submission','Submitted at','Verdict','Error','Created at','Pricing basis JSON'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Model Comparison')
    .addItem('Run rows…', 'fcrShowDialog')
    .addItem('Open row results', 'fcrOpenResults')
    .addItem('Stop row comparison', 'fcrStop').addToUi();
}

function fcConfig() {
  const p = PropertiesService.getScriptProperties().getProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!p.TEST_SPREADSHEET_ID || p.TEST_SPREADSHEET_ID !== ss.getId()) {
    throw new Error('Set TEST_SPREADSHEET_ID to the separate test-copy spreadsheet ID.');
  }
  if (ss.getId() === '1fWrmaiPJcM463E6FIQsf-wMoVYN7erSni7ga58XBsRk') throw new Error('Do not run in the original source spreadsheet. Make a separate copy.');
  if (ScriptApp.getScriptId() === '1Py-kvsz3Af7pkXhh6KiEAH1ldKt_cZg0-q3dNasBFty4J7gKdKxJW7ij') throw new Error('Use a separate bound Apps Script project, not the original project.');
  // Baseline is imported from Master, never regenerated or resubmitted.
  const models = [...new Set((p.COMPARISON_MODELS || 'gemini-3.8-flash').split(',').map(x => x.trim()).filter(x => x && x !== 'gemini-2.5-pro'))];
  if (!models.length) throw new Error('Select at least one new comparison model.');
  return {p, ss, models};
}

function fcSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(headers); s.setFrozenRows(1); }
  const actual = s.getRange(1,1,1,headers.length).getValues()[0];
  if (JSON.stringify(actual) !== JSON.stringify(headers)) throw new Error(name + ': unexpected headers');
  return s;
}

function fcLocked(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another comparison action is running.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function fcHash(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes).map(b => ((b+256)%256).toString(16).padStart(2,'0')).join('');
}

function fcWrite(s, row, values) {
  Object.keys(values).forEach(k => {
    const i = FC_HEADERS.indexOf(k);
    if (i < 0) throw new Error('Unknown result field ' + k);
    s.getRange(row,i+1).setValue(values[k]);
  });
  SpreadsheetApp.flush();
}

function fcObject(row) { return Object.fromEntries(FC_HEADERS.map((k,i) => [k,row[i]])); }
function fcFileId(url) {
  const match = String(url).match(/[-\w]{25,}/);
  if (!match) throw new Error('Invalid Drive URL');
  return match[0];
}

function fcSavedBaseline(head, row, folder, shared) {
  const read = name => { const i=head.indexOf(name); return i<0 ? '' : row[i]; };
  const notes=['Historical baseline: prompt/image bytes and judge tests at original generation are not verified; blank values mean unknown.'];
  const saved = {...shared, Model:'gemini-2.5-pro', State:'BASELINE_SAVED','Prompt SHA256':'','Image SHA256':'','API ms':'','Estimated USD':''};
  const model=String(read('Model Used')).trim();
  const record=folder.createFile('row-'+shared['Source row']+'-historical-record.json',JSON.stringify({headers:head,values:row}),'application/json');
  saved['Raw response URL']=record.getUrl(); // Source record, not a fabricated Gemini API response.
  if (model !== 'gemini-2.5-pro') {
    saved.State='BASELINE_MISSING'; saved.Error='No stored gemini-2.5-pro result for this source row.'; return saved;
  }
  saved['Model version']=read('Model version') || read('Model Version') || model;
  const codeURL=read('Code File URL');
  if (codeURL) {
    try {
      const content=DriveApp.getFileById(fcFileId(codeURL)).getBlob().getDataAsString();
      const snapshot=folder.createFile('row-'+shared['Source row']+'-saved-2.5-pro.py',content,'text/plain');
      saved['Code URL']=snapshot.getUrl(); saved['Code SHA256']=fcHash(content);
    } catch(e) { notes.push('Saved code unavailable: '+String(e)); }
  } else notes.push('Saved code URL missing.');
  const verdict=String(read('Verdict')).trim();
  // Never fish a judgement out of a displaced timestamp/token/other column.
  if (/^(AC|WA|TLE|RTE|RE|CE|MLE|OLE|PE|NO|JE|ACCEPTED|CORRECT|WRONG_ANSWER|WRONG ANSWER|TIMELIMIT|RUN_ERROR|COMPILER_ERROR)$/i.test(verdict)) saved.Verdict=verdict;
  else notes.push('Stored Verdict missing or invalid; no guessed result.');
  const numeric=x=>x!=='' && x!==null && !(x instanceof Date) && Number.isFinite(Number(x)) && Number(x)>=0;
  const input=read('Input Tokens'),output=read('Output Tokens'),total=read('Total Tokens'),thought=read('Thoughts Token Count');
  if ([input,output,total].every(numeric) && Number(total)>=Number(input)+Number(output) && (!numeric(thought)||Number(total)>=Number(input)+Number(output)+Number(thought))) {
    saved['Prompt tokens']=Number(input); saved['Output tokens']=Number(output); saved['Total tokens']=Number(total);
    if(numeric(thought))saved['Thought tokens']=Number(thought);
    const ms=read('Response Time (ms)'); if(numeric(ms))saved['API ms']=Number(ms);
  } else notes.push('Historical token/time fields missing or inconsistent; left unknown.');
  const id=read('Submission ID'); if(/^\d+$/.test(String(id)))saved['DOMjudge submission']=String(id);
  saved.Error=notes.join(' ');
  return saved;
}

/** Snapshot input and import baseline; only new models are queued. */
function fcPrepare() {
  return fcLocked(() => {
    const {ss,p,models} = fcConfig();
    const results = fcSheet(ss,'ModelResults',FC_HEADERS);
    const manifest = fcSheet(ss,'ComparisonRuns',['Run ID','Folder ID','Prompt file ID','Settings JSON']);
    const pricing = fcSheet(ss,'ModelPricing',['Model','Input USD per 1M','Output USD per 1M including thinking','Price verified date','Source URL']);
    models.forEach(m => {
      if (!pricing.getDataRange().getValues().slice(1).some(r => r[0] === m)) pricing.appendRow([m,'','','','']);
    });
    const source = ss.getSheetByName('Master');
    if (!source) throw new Error('Copy the Master and Meta tabs into this test spreadsheet first.');
    const data = source.getDataRange().getValues(), head = data[0];
    ['Problem Number','Upload your Flowchart'].forEach(h => { if (!head.includes(h)) throw new Error('Missing column: ' + h); });
    const selected = p.SOURCE_ROWS || '2';
    const all = selected.toUpperCase() === 'ALL';
    const requested = selected.split(',').map(Number);
    const chosen = data.slice(1).map((row,i) => ({row,n:i+2})).filter(x => (all || requested.includes(x.n)) && x.row[head.indexOf('Upload your Flowchart')]);
    if (!chosen.length) throw new Error('No selected rows with flowchart URLs. Set SOURCE_ROWS to row numbers or ALL.');
    const cfgText = JSON.stringify({baseline:'saved-only-v1',models,sourceRows:chosen.map(x=>x.n),temperature:0,prompt:getPrompt()});
    const key = fcHash(cfgText);
    if (manifest.getDataRange().getValues().slice(1).some(r => JSON.parse(r[3]).key === key)) throw new Error('This selection is already queued. Continue existing rows; do not prepare again.');
    const run = Utilities.getUuid();
    const folder = DriveApp.createFolder('FlowchartComparison-' + run);
    const promptFile = folder.createFile('prompt.txt',getPrompt(),'text/plain');
    const settings = {key, models, temperature:0, judgeBase:p.DOMJUDGE_BASE_URL || '',contest:p.DOMJUDGE_CONTEST_ID || '',language:p.DOMJUDGE_LANGUAGE_ID || 'python3',team:p.DOMJUDGE_TEAM_ID || ''};
    manifest.appendRow([run,folder.getId(),promptFile.getId(),JSON.stringify(settings)]);
    chosen.forEach(({row,n}) => {
      try {
        // One immutable input snapshot shared by every model for this source row.
        const blob = DriveApp.getFileById(fcFileId(row[head.indexOf('Upload your Flowchart')])).getBlob();
        const snapshot = folder.createFile(blob).setName('row-' + n + '-' + blob.getName());
        const shared={'Run ID':run,'Source row':n,Problem:row[head.indexOf('Problem Number')],'Image snapshot':snapshot.getUrl(),'Created at':new Date()};
        const baseline=fcSavedBaseline(head,row,folder,shared);
        results.appendRow(FC_HEADERS.map(h => baseline[h] === undefined ? '' : baseline[h]));
        models.forEach(model => {
          const value = {'Run ID':run,'Source row':n,Problem:row[head.indexOf('Problem Number')],'Image snapshot':snapshot.getUrl(),'Image SHA256':fcHash(blob.getBytes()),'Prompt SHA256':fcHash(getPrompt()),Model:model,State:'QUEUED','Created at':new Date()};
          results.appendRow(FC_HEADERS.map(h => value[h] === undefined ? '' : value[h]));
        });
      } catch(e) {
        models.forEach(model => {
          const value = {'Run ID':run,'Source row':n,Model:model,State:'INPUT_ERROR',Error:String(e),'Created at':new Date()};
          results.appendRow(FC_HEADERS.map(h => value[h] === undefined ? '' : value[h]));
        });
      }
    });
    fcReport();
  });
}

function fcRun(ss, id) {
  const r = ss.getSheetByName('ComparisonRuns').getDataRange().getValues().slice(1).find(x => x[0] === id);
  if (!r) throw new Error('Missing run manifest');
  return {folder:DriveApp.getFolderById(r[1]),prompt:DriveApp.getFileById(r[2]).getBlob().getDataAsString(),settings:JSON.parse(r[3])};
}

/** Ordinary paid synchronous Gemini calls; no batch-price assumption. */
function fcCost(usage, rate) {
  if (typeof usage.promptTokenCount !== 'number' || typeof usage.candidatesTokenCount !== 'number') return '';
  if (!rate || rate.slice(1,3).some(x => x === '' || !Number.isFinite(Number(x)) || Number(x)<0) || !rate[3] || !rate[4]) return '';
  // Thoughts are billed output in addition to visible candidate tokens.
  return ((usage.promptTokenCount || 0)*Number(rate[1]) + ((usage.candidatesTokenCount || 0)+(usage.thoughtsTokenCount || 0))*Number(rate[2]))/1000000;
}

function fcGenerate() {
  return fcLocked(() => {
    const {ss,p} = fcConfig();
    if (!p.GEMINI_API_KEY) throw new Error('Missing GEMINI_API_KEY Script Property');
    const s = fcSheet(ss,'ModelResults',FC_HEADERS), data = s.getDataRange().getValues();
    const rates = ss.getSheetByName('ModelPricing').getDataRange().getValues().slice(1);
    const start = Date.now(); let count = 0;
    for (let i=1;i<data.length && count<3 && Date.now()-start<180000;i++) {
      const job = fcObject(data[i]); if (job.State !== 'QUEUED') continue;
      // Covers stale queues prepared by the earlier version as well.
      if(job.Model==='gemini-2.5-pro') {fcWrite(s,i+1,{State:'BASELINE_NOT_RERUN',Error:'Old queue: prepare saved-baseline comparison instead. No Gemini call made.'});continue;}
      const run = fcRun(ss,job['Run ID']);
      fcWrite(s,i+1,{State:'GENERATING'}); count++;
      try {
        const blob = DriveApp.getFileById(fcFileId(job['Image snapshot'])).getBlob();
        if (fcHash(blob.getBytes()) !== job['Image SHA256'] || fcHash(run.prompt) !== job['Prompt SHA256']) throw new Error('Input or prompt changed after preparation');
        const tick = Date.now();
        const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(job.Model)+':generateContent', {
          method:'post',contentType:'application/json',headers:{'x-goog-api-key':p.GEMINI_API_KEY},muteHttpExceptions:true,
          payload:JSON.stringify({contents:[{parts:[{inline_data:{mime_type:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())}},{text:run.prompt}]}],generationConfig:{temperature:run.settings.temperature}})
        });
        const ms=Date.now()-tick, text=resp.getContentText();
        const raw=run.folder.createFile('row-'+job['Source row']+'-'+job.Model+'-response.json',text,'application/json');
        fcWrite(s,i+1,{'Raw response URL':raw.getUrl(),'API ms':ms});
        if (resp.getResponseCode()<200 || resp.getResponseCode()>=300) throw new Error('Gemini HTTP '+resp.getResponseCode()+'; inspect raw response');
        const json=JSON.parse(text), candidate=(json.candidates || [])[0] || {}, usage=json.usageMetadata || {};
        const rate=rates.find(r=>r[0]===job.Model);
        fcWrite(s,i+1,{'Model version':json.modelVersion || job.Model,'Prompt tokens':usage.promptTokenCount ?? '','Output tokens':usage.candidatesTokenCount ?? '','Thought tokens':usage.thoughtsTokenCount ?? '','Total tokens':usage.totalTokenCount ?? '','Estimated USD':fcCost(usage,rate),'Pricing date':rate ? rate[3] : '','Pricing basis JSON':rate ? JSON.stringify(rate) : ''});
        if (candidate.finishReason !== 'STOP') throw new Error('Incomplete/blocked output: '+candidate.finishReason);
        const code=((candidate.content || {}).parts || []).filter(x=>!x.thought && typeof x.text==='string').map(x=>x.text).join('\n').replace(/^\s*```(?:python|py)?\s*\n/i,'').replace(/\n?```\s*$/,'').trim();
        if (!code) throw new Error('Empty generated code');
        const file=run.folder.createFile('row-'+job['Source row']+'-'+job.Model+'.py',code,'text/plain');
        fcWrite(s,i+1,{'Code URL':file.getUrl(),'Code SHA256':fcHash(code),State:'GENERATED'});
      } catch(e) { fcWrite(s,i+1,{State:'GENERATION_ERROR',Error:String(e)}); }
    }
    fcReport();
  });
}

function fcJudge(p,settings,path,method,payload) {
  if (!settings.judgeBase || !settings.contest || !p.DOMJUDGE_USERNAME || !p.DOMJUDGE_PASSWORD) throw new Error('Missing DOMjudge Script Properties; configure them before preparing a run');
  const base=settings.judgeBase.replace(/\/+$/,'').replace(/\/api\/v4$/,'')+'/api/v4';
  const response=UrlFetchApp.fetch(base+'/contests/'+encodeURIComponent(settings.contest)+path,{
    method,headers:{Authorization:'Basic '+Utilities.base64Encode(p.DOMJUDGE_USERNAME+':'+p.DOMJUDGE_PASSWORD)},contentType:'application/json',muteHttpExceptions:true,...(payload ? {payload:JSON.stringify(payload)} : {})
  });
  if (response.getResponseCode()<200 || response.getResponseCode()>=300) throw new Error('DOMjudge HTTP '+response.getResponseCode());
  return JSON.parse(response.getContentText());
}

function fcSubmit() {
  return fcLocked(() => {
    const {ss,p}=fcConfig(), s=fcSheet(ss,'ModelResults',FC_HEADERS), data=s.getDataRange().getValues();
    const meta=ss.getSheetByName('Meta'); if (!meta) throw new Error('Missing Meta mapping');
    const problems=meta.getDataRange().getValues();
    let count=0;
    for (let i=1;i<data.length && count<3;i++) {
      const job=fcObject(data[i]); if (job.State!=='GENERATED') continue;
      if(job.Model==='gemini-2.5-pro') {fcWrite(s,i+1,{State:'BASELINE_NOT_RESUBMITTED',Error:'Historical baseline must not be resubmitted.'});continue;}
      try {
        const run=fcRun(ss,job['Run ID']);
        const mapping=problems.slice(1).find(r=>String(r[0]).trim()===String(job.Problem).split(' ')[0].trim());
        if (!mapping || mapping[1]==='') throw new Error('Problem not found in Meta');
        const code=DriveApp.getFileById(fcFileId(job['Code URL'])).getBlob().getDataAsString();
        if (fcHash(code)!==job['Code SHA256']) throw new Error('Generated code modified');
        const zip=Utilities.zip([Utilities.newBlob(code,'text/x-python','solution.py')],'solution.zip');
        const payload={problem_id:String(mapping[1]),language_id:run.settings.language,files:[{filename:'solution.zip',data:Utilities.base64Encode(zip.getBytes())}]};
        if (run.settings.team) payload.team_id=run.settings.team;
        fcWrite(s,i+1,{State:'SUBMITTING'}); count++;
        const result=fcJudge(p,run.settings,'/submissions','post',payload);
        if (!result.id) throw new Error('No submission ID returned; inspect judge before retrying');
        fcWrite(s,i+1,{'DOMjudge submission':String(result.id),'Submitted at':new Date(),State:'JUDGE_SUBMITTED'});
      } catch(e) { fcWrite(s,i+1,{State:'SUBMISSION_ERROR',Error:String(e)}); }
    }
    fcReport();
  });
}

function fcPickVerdict(judgements,id) {
  if (!Array.isArray(judgements)) throw new Error('Unexpected judgement response');
  const matching=judgements.filter(j=>String(j.submission_id)===String(id));
  // A newer unfinished rejudging must not return a previous completed verdict.
  matching.sort((a,b)=>Date.parse(b.start_time || 0)-Date.parse(a.start_time || 0));
  const latest=matching[0];
  return latest && latest.end_time && latest.judgement_type_id ? latest.judgement_type_id : '';
}

function fcPoll() {
  return fcLocked(() => {
    const {ss,p}=fcConfig(),s=fcSheet(ss,'ModelResults',FC_HEADERS),data=s.getDataRange().getValues();
    const start=Date.now();
    for(let i=1;i<data.length && Date.now()-start<180000;i++) {
      const job=fcObject(data[i]); if(job.State!=='JUDGE_SUBMITTED') continue;
      try {
        const run=fcRun(ss,job['Run ID']);
        const response=fcJudge(p,run.settings,'/judgements?submission_id='+encodeURIComponent(job['DOMjudge submission']),'get');
        const verdict=fcPickVerdict(response,job['DOMjudge submission']);
        if(verdict) fcWrite(s,i+1,{Verdict:verdict,State:'VERDICT_READY',Error:''});
      } catch(e) { fcWrite(s,i+1,{Error:String(e)}); }
    }
    fcReport();
  });
}

function fcReport() {
  const {ss}=fcConfig(), source=ss.getSheetByName('ModelResults'); if(!source) return;
  const rows=source.getDataRange().getValues().slice(1).map(fcObject),groups=new Map();
  rows.forEach(j=>{const k=j['Run ID']+'|'+j['Source row'];if(!groups.has(k))groups.set(k,[]);groups.get(k).push(j);});
  const headers=['Run ID','Source row','Problem','Model A','Model B','Same image and prompt','Same code','Verdict A','Verdict B','Same verdict','API ms A','API ms B','Estimated USD A','Estimated USD B','State A','State B'];
  const out=[];
  groups.forEach(g=>{for(let a=0;a<g.length;a++)for(let b=a+1;b<g.length;b++){
    const x=g[a],y=g[b];
    out.push([x['Run ID'],x['Source row'],x.Problem,x.Model,y.Model,x['Image SHA256']&&y['Image SHA256']&&x['Prompt SHA256']&&y['Prompt SHA256']?x['Image SHA256']===y['Image SHA256']&&x['Prompt SHA256']===y['Prompt SHA256']:'unknown',x['Code SHA256']&&y['Code SHA256']?x['Code SHA256']===y['Code SHA256']:'',x.Verdict,y.Verdict,x.Verdict&&y.Verdict?x.Verdict===y.Verdict:'',x['API ms'],y['API ms'],x['Estimated USD'],y['Estimated USD'],x.State,y.State]);
  }});
  const report=fcSheet(ss,'ModelComparison',headers);
  if(report.getLastRow()>1)report.getRange(2,1,report.getLastRow()-1,headers.length).clearContent();
  if(out.length)report.getRange(2,1,out.length,headers.length).setValues(out);
  const summaryHeaders=['Run ID','Model','Jobs','Verdicts ready','Accepted','Accepted / completed','Errors','Average successful API ms','Known estimated USD','Rows with known cost'];
  const aggregate=new Map();
  rows.forEach(j=>{
    const key=j['Run ID']+'|'+j.Model;
    if(!aggregate.has(key))aggregate.set(key,{run:j['Run ID'],model:j.Model,total:0,done:0,accepted:0,errors:0,times:[],cost:0,priced:0});
    const a=aggregate.get(key); a.total++;
    if(j.State==='VERDICT_READY'||(j.State==='BASELINE_SAVED'&&j.Verdict)){a.done++;if(['AC','ACCEPTED','CORRECT'].includes(String(j.Verdict).toUpperCase()))a.accepted++;}
    if(String(j.State).endsWith('_ERROR'))a.errors++;
    if(j['Code URL'] && typeof j['API ms']==='number')a.times.push(j['API ms']);
    if(typeof j['Estimated USD']==='number'){a.cost+=j['Estimated USD'];a.priced++;}
  });
  const summary=fcSheet(ss,'ModelSummary',summaryHeaders);
  if(summary.getLastRow()>1)summary.getRange(2,1,summary.getLastRow()-1,summaryHeaders.length).clearContent();
  const stats=[...aggregate.values()].map(a=>[a.run,a.model,a.total,a.done,a.accepted,a.done?a.accepted/a.done:'',a.errors,a.times.length?a.times.reduce((x,y)=>x+y,0)/a.times.length:'',a.priced?a.cost:'',a.priced]);
  if(stats.length)summary.getRange(2,1,stats.length,summaryHeaders.length).setValues(stats);
}

/** One-file pilot using the FCP010 image verified in the user's browser. */
function fcPrepareFcp010() {
  fcConfig();
  PropertiesService.getScriptProperties().setProperty('SOURCE_ROWS', '3');
  return fcPrepare();
}

/** Verify the dev-contest mapping before submitting the FCP010 pilot. */
function fcVerifyFcp010Judge() {
  const {ss,p}=fcConfig();
  const jobs=ss.getSheetByName('ModelResults').getDataRange().getValues().slice(1).map(fcObject);
  const job=jobs.find(j=>j.Model==='gemini-3.8-flash' && j['Source row']===3 && j.State==='GENERATED');
  if(!job)throw new Error('No generated FCP010 pilot found');
  const run=fcRun(ss,job['Run ID']);
  const problems=fcJudge(p,run.settings,'/problems','get');
  if(!Array.isArray(problems))throw new Error('Unexpected problem-list response');
  const check=fcSheet(ss,'JudgeCheck',['ID','Label','Name','External ID']);
  const values=problems.map(x=>[String(x.id),x.label||'',x.name||'',x.externalid||x.external_id||'']);
  if(values.length)check.getRange(2,1,values.length,4).setValues(values);
  const exact=problems.filter(x=>[x.label,x.name,x.externalid,x.external_id].some(v=>/^FCP010(?:\b|\s|$)/i.test(String(v||''))));
  if(exact.length!==1)throw new Error('Verify FCP010 mapping in JudgeCheck; no submission made');
  const meta=ss.getSheetByName('Meta'),rows=meta.getDataRange().getValues();
  const i=rows.findIndex(r=>String(r[0]).trim()==='FCP010');
  if(i<1)throw new Error('Missing FCP010 Meta row');
  meta.getRange(i+1,2).setValue(String(exact[0].id));
  console.log('Verified FCP010 dev problem ID: '+exact[0].id);
}

/** Display the completed pilot honestly, including the unmatched judge target. */
function fcShowFcp010Pilot() {
  const {ss}=fcConfig(),results=ss.getSheetByName('ModelResults');
  const all=results.getDataRange().getValues().slice(1).map(fcObject);
  const current=all.find(j=>j['Source row']===3&&j.Model==='gemini-3.8-flash'&&j['Code URL']);
  if(!current)throw new Error('FCP010 generation not complete');
  const old=all.find(j=>j['Run ID']===current['Run ID']&&j.Model==='gemini-2.5-pro');
  const row=all.indexOf(current)+2;
  const rate=['gemini-3.8-flash',0.75,3.75,'2026-10-24','https://ai.google.dev/gemini-api/docs/pricing#gemini-3.8-flash'];
  const cost=fcCost({promptTokenCount:current['Prompt tokens'],candidatesTokenCount:current['Output tokens'],thoughtsTokenCount:current['Thought tokens']},rate);
  fcWrite(results,row,{'Estimated USD':cost,'Pricing date':rate[3],'Pricing basis JSON':JSON.stringify(rate),Error:'Configured dev contest does not contain FCP010. No DOMjudge submission made.'});
  const prices=ss.getSheetByName('ModelPricing');
  const priceRows=prices.getDataRange().getValues();
  const pi=priceRows.findIndex(r=>r[0]===rate[0]);
  if(pi>0)prices.getRange(pi+1,1,1,5).setValues([rate]);
  const originalCode=old&&old['Code URL']?DriveApp.getFileById(fcFileId(old['Code URL'])).getBlob().getDataAsString():'';
  const newCode=DriveApp.getFileById(fcFileId(current['Code URL'])).getBlob().getDataAsString();
  const out=[
    ['FCP010 pilot','Saved 2.5 Pro result','New 3.8 Flash result'],
    ['Code file',old?old['Code URL']:'',current['Code URL']],
    ['Input tokens',old?old['Prompt tokens']:'unknown',current['Prompt tokens']],
    ['Output tokens',old?old['Output tokens']:'unknown',current['Output tokens']],
    ['Thinking tokens',old?old['Thought tokens']:'unknown',current['Thought tokens']],
    ['Total tokens',old?old['Total tokens']:'unknown',current['Total tokens']],
    ['API time (seconds)',old&&typeof old['API ms']==='number'?old['API ms']/1000:'unknown',current['API ms']/1000],
    ['Estimated paid-standard USD','Historical price unknown',cost],
    ['Verdict for these code files','Master verdict missing/misaligned','Not submitted: FCP010 absent in configured dev contest'],
    ['Historical reference for SAME IMAGE','AC; original judge submission 2187; generated code linkage not verified','https://docs.google.com/spreadsheets/d/1VHFvbcIukZhSB9wmTMSFXuSpRu49ySVq63_SpUz_KUs/edit#gid=317642154'],
    ['Python code',originalCode,newCode],
    ['Comparison limits','Old prompt bytes and original tests unverified','Current prompt used; historical verdict is not attributed to the stored 2.5 code'],
    ['Pricing source','Standard input $0.75 / 1M; output including thinking $3.75 / 1M as checked 2026-10-24',rate[4]],
    ['Billing note','Estimated token cost only','Actual billing/free-tier/caching/taxes may differ']
  ];
  let summary=ss.getSheetByName('PilotSummary'); if(!summary)summary=ss.insertSheet('PilotSummary');
  summary.getRange(1,1,out.length,3).setValues(out);
  summary.setFrozenRows(1);summary.setColumnWidth(1,230);summary.setColumnWidths(2,2,420);
  summary.getRange(1,1,out.length,3).setWrap(true);
  summary.getRange(1,1,1,3).setFontWeight('bold');
  fcReport();
  ss.setActiveSheet(summary);
}

/** Row-range workflow. Restricted to the user's comparison copy and dev judge. */
const FCR_COPY = '1U-FahEyZvmMa5Y2uaLmoBUDL-q-Zqp14ss2fYqFwWU4';
const FCR_COLS = ['Run ID','Source row','Problem','Status','Old model','New model','Old dev verdict','New dev verdict','Old API seconds (saved)','New API seconds','Old total tokens (saved)','New total tokens','New estimated USD','Old code URL','New code URL','Old dev submission','New dev submission','Notes','Updated at','Source snapshot URL','Image snapshot URL','Image SHA256','Prompt SHA256','Old code SHA256','New code SHA256','Raw response URL','New input tokens','New output tokens','New thinking tokens','Returned model','Judge problem ID','Started at','Old input tokens (saved)','Old output tokens (saved)','Old thinking tokens (saved)','Judge metadata JSON'];
const FCR_RUN_COLS = ['Run ID','State','Start row','End row','New model','Folder ID','Prompt file ID','Settings JSON','Selection SHA256','Created at'];

function fcrEnvironment() {
  const p=PropertiesService.getScriptProperties().getProperties();
  if(p.TEST_SPREADSHEET_ID!==FCR_COPY) throw new Error('This workflow is restricted to the comparison copy.');
  if(ScriptApp.getScriptId()==='1Py-kvsz3Af7pkXhh6KiEAH1ldKt_cZg0-q3dNasBFty4J7gKdKxJW7ij') throw new Error('Original project is blocked.');
  const base=String(p.DOMJUDGE_BASE_URL||'').replace(/\/+$/,'');
  if(base!=='https://judgedev.csbasics.in') throw new Error('Only the dev DOMjudge is allowed.');
  if(!p.DOMJUDGE_CONTEST_ID||!p.DOMJUDGE_USERNAME||!p.DOMJUDGE_PASSWORD||!p.GEMINI_API_KEY) throw new Error('Missing API credentials in Script Properties.');
  return {p,ss:SpreadsheetApp.openById(FCR_COPY)};
}
function fcrTable(ss,name,headers) {
  const fresh=!ss.getSheetByName(name), s=fcSheet(ss,name,headers);
  if(fresh) {
    s.setFrozenRows(1); s.getRange(1,1,1,headers.length).setBackground('#17365d').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
    s.setColumnWidths(1,headers.length,140); s.setRowHeight(1,48);
    if(name==='RowComparisons') {s.setColumnWidth(3,280);s.setColumnWidth(4,200);s.setColumnWidth(18,420);s.setFrozenColumns(3);}
  }
  return s;
}
function fcrRecord(values) {return Object.fromEntries(FCR_COLS.map((k,i)=>[k,values[i]]));}
function fcrPut(s,n,changes) {
  const v=s.getRange(n,1,1,FCR_COLS.length).getValues()[0];
  Object.keys(changes).forEach(k=>{const i=FCR_COLS.indexOf(k);if(i<0)throw new Error('Unknown column '+k);v[i]=fcrLiteral(changes[k]);});
  v[FCR_COLS.indexOf('Updated at')]=new Date(); s.getRange(n,1,1,v.length).setValues([v]); SpreadsheetApp.flush();
}
function fcrLiteral(v) {return typeof v==='string'&&/^[=+@]/.test(v)?"'"+v:v;}
function fcrNumber(v) {return typeof v==='number'&&Number.isFinite(v)&&v>=0?v:'';}
function fcrNormal(s) {return String(s||'').replace(/^FCP\d+\s*/i,'').toLowerCase().replace(/[^a-z0-9]/g,'');}
function fcrMatch(name,problems) {
  // Reviewed PDF D-46: J<=I, boundary tests J==1, J==I, I==N: hollow right triangle.
  // Match both dev ID and its name; do not infer mappings for other generic Pattern rows.
  if(String(name).trim()==='FCP046 Pattern') {
    const verified=problems.filter(p=>p.id==='hollow-right-triangle-star'&&fcrNormal(p.name)==='hollowrighttrianglestarpattern');
    return verified.length===1?verified[0]:null;
  }
  const found=problems.filter(p=>fcrNormal(p.name)===fcrNormal(name));
  return found.length===1?found[0]:null;
}
function fcrJudge(p,settings,path,payload) {
  if(settings.judgeBase!=='https://judgedev.csbasics.in')throw new Error('Production judge blocked');
  const response=UrlFetchApp.fetch(settings.judgeBase+'/api/v4/contests/'+encodeURIComponent(settings.contest)+path,{
    method:payload?'post':'get',contentType:'application/json',muteHttpExceptions:true,
    headers:{Authorization:'Basic '+Utilities.base64Encode(p.DOMJUDGE_USERNAME+':'+p.DOMJUDGE_PASSWORD)},
    ...(payload?{payload:JSON.stringify(payload)}:{})
  });
  const status=response.getResponseCode();
  if(status<200||status>=300)throw new Error('Dev DOMjudge HTTP '+status);
  return JSON.parse(response.getContentText());
}
function fcrDone(state) {return /^(DONE|SKIPPED|ERROR|REVIEW_REQUIRED|CANCELLED)$/.test(state);}

function fcrShowDialog() {
  fcrEnvironment();
  const html=`<!doctype html><html><head><base target="_top"><style>body{font:15px Arial;padding:18px;color:#17365d}label{display:block;margin:14px 0 5px}input{box-sizing:border-box;padding:10px;width:100%;font-size:15px}button{margin-top:18px;padding:12px 22px;background:#1765cc;color:white;border:0;border-radius:5px;font-size:15px}p{line-height:1.5;color:#444}#message{white-space:pre-wrap}</style></head><body><h2>Run comparison rows</h2><p>Source: <b>Master</b> tab. Row 1 contains headers.<br>Uses the saved old code; generates only the new model's code. Both codes are tested on the same dev problem.</p><label for="start">Start row</label><input id="start" type="number" min="2" value="33"><label for="end">End row</label><input id="end" type="number" min="2" value="33"><label for="model">New model API name</label><input id="model" value="gemini-3.8-flash"><p>Generation uses your Gemini API account. Rows without an exact dev problem match are skipped before generation. Up to 100 rows per run.</p><button id="run" onclick="startRun()">Run comparison</button><p id="message" role="status"></p><script>function startRun(){var b=document.getElementById('run'),m=document.getElementById('message');b.disabled=true;m.textContent='Checking and queuing rows…';google.script.run.withSuccessHandler(function(r){m.textContent=r.message;b.textContent='Queued';}).withFailureHandler(function(e){m.textContent=e.message;b.disabled=false;}).fcrStart({start:document.getElementById('start').value,end:document.getElementById('end').value,model:document.getElementById('model').value});}</script></body></html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(620),'Model Comparison');
}
function fcrInstall() {
  const {ss}=fcrEnvironment();
  fcrTable(ss,'RowComparisons',FCR_COLS);fcrTable(ss,'RowComparisonRuns',FCR_RUN_COLS);
  console.log('Row result tabs installed. Reload the copied Sheet to load the menu. No Gemini call or judge submission made.');
}
function fcrOpenResults() {const {ss}=fcrEnvironment();ss.setActiveSheet(fcrTable(ss,'RowComparisons',FCR_COLS));}
function fcrTriggers(remove) {
  const mine=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='fcrWorker');
  if(remove)mine.forEach(t=>ScriptApp.deleteTrigger(t));
  else if(!mine.length)ScriptApp.newTrigger('fcrWorker').timeBased().everyMinutes(1).create();
}
function fcrStart(args) {
 return fcLocked(()=>{
  const {ss,p}=fcrEnvironment(), start=Number(args.start),end=Number(args.end),model=String(args.model||'').trim();
  const source=ss.getSheetByName('Master');if(!source)throw new Error('Missing Master tab');
  if(!Number.isInteger(start)||!Number.isInteger(end)||start<2||end<start||end>source.getLastRow()||end-start>=100)throw new Error('Choose valid Master rows, starting at 2; maximum 100 rows.');
  if(!/^gemini-[a-z0-9.-]+$/.test(model))throw new Error('Enter a Gemini model API name.');
  if(model==='gemini-2.5-pro')throw new Error('2.5 Pro fresh generation is unavailable for this project. Its saved code is used automatically.');
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('FCR_ACTIVE_RUN'))throw new Error('A row comparison is already active. Open results or stop it first.');
  const runs=fcrTable(ss,'RowComparisonRuns',FCR_RUN_COLS),results=fcrTable(ss,'RowComparisons',FCR_COLS);
  const settings={judgeBase:'https://judgedev.csbasics.in',contest:p.DOMJUDGE_CONTEST_ID,language:p.DOMJUDGE_LANGUAGE_ID||'python3',temperature:0};
  const problems=fcrJudge(p,settings,'/problems');if(!Array.isArray(problems))throw new Error('Unexpected problem list');
  const head=source.getRange(1,1,1,source.getLastColumn()).getValues()[0],data=source.getRange(start,1,end-start+1,head.length).getValues();
  ['Problem Number','Upload your Flowchart','Code File URL','Model Used'].forEach(h=>{if(!head.includes(h))throw new Error('Missing Master column: '+h);});
  const prompt=getPrompt(), problemMappings=data.map(row=>{const m=fcrMatch(row[head.indexOf('Problem Number')],problems);return m?m.id:null;}), key=fcHash(JSON.stringify({start,end,model,data,prompt,settings,problemMappings}));
  if(runs.getDataRange().getValues().slice(1).some(r=>r[8]===key))throw new Error('This exact selection was already queued. Use its existing results; duplicate calls are blocked.');
  const id=Utilities.getUuid(),folder=DriveApp.createFolder('FlowchartRows-'+id),promptFile=folder.createFile('prompt.txt',prompt,'text/plain');
  folder.createFile('source-rows.json',JSON.stringify({headers:head,start,rows:data,problems}),'application/json');
  const priceSheet=ss.getSheetByName('ModelPricing');
  settings.rate=priceSheet?priceSheet.getDataRange().getValues().slice(1).find(r=>r[0]===model)||null:null;
  runs.appendRow([id,'QUEUED',start,end,model,folder.getId(),promptFile.getId(),JSON.stringify(settings),key,new Date()]);
  const values=data.map((row,i)=>{
    const read=h=>row[head.indexOf(h)],problem=String(read('Problem Number')||''),match=fcrMatch(problem,problems);
    const state=match?'QUEUED':'SKIPPED';
    const job={'Run ID':id,'Source row':start+i,Problem:problem,Status:state,'Old model':String(read('Model Used')||''),'New model':model,'Judge problem ID':match?String(match.id):'','Judge metadata JSON':match?JSON.stringify(match):'','Prompt SHA256':fcHash(prompt),'Updated at':new Date(),Notes:match?'Old tokens/time are historical; old prompt bytes are unverified. Both saved and new code will be judged now.':'No unique exact dev-contest problem-name match. No Gemini call or submission made.'};
    return FCR_COLS.map(h=>fcrLiteral(job[h]===undefined?'':job[h]));
  });
  results.getRange(results.getLastRow()+1,1,values.length,FCR_COLS.length).setValues(values);
  if(values.some(r=>r[3]==='QUEUED')){
    try {fcrTriggers(false);props.setProperty('FCR_ACTIVE_RUN',id);}catch(e){runs.getRange(runs.getLastRow(),2).setValue('SETUP_ERROR');throw e;}
  }else runs.getRange(runs.getLastRow(),2).setValue('COMPLETE');
  ss.setActiveSheet(results);SpreadsheetApp.flush();
  return {runId:id,message:'Rows '+start+'–'+end+' queued. Open RowComparisons for live status. Processing starts within about a minute; you may close this dialog. Rows without a match are marked SKIPPED.'};
 });
}
function fcrStop() {
 const {ss}=fcrEnvironment(),props=PropertiesService.getScriptProperties(),id=props.getProperty('FCR_ACTIVE_RUN');
 if(id){props.setProperty('FCR_STOP_RUN',id);ss.toast('Stop requested. The current request will finish; no further stage will start.','Model Comparison',8);}
 else {fcrTriggers(true);ss.toast('No row comparison is active.','Model Comparison',5);}
}
function fcrCancelQueued(ss,id) {
 const props=PropertiesService.getScriptProperties(),s=ss.getSheetByName('RowComparisons');
 s.getDataRange().getValues().slice(1).forEach((r,i)=>{if(r[0]===id&&!fcrDone(r[3]))fcrPut(s,i+2,{Status:'CANCELLED',Notes:'Stopped by user. Existing calls/submissions are retained; no automatic retry.'});});
 fcrRunState(ss,id,'CANCELLED');props.deleteProperty('FCR_ACTIVE_RUN');props.deleteProperty('FCR_STOP_RUN');fcrTriggers(true);
}

function fcrRunState(ss,id,state) {const s=ss.getSheetByName('RowComparisonRuns'),r=s.getDataRange().getValues().findIndex(v=>v[0]===id);if(r>0)s.getRange(r+1,2).setValue(state);}
function fcrContext(ss,id) {
 const r=ss.getSheetByName('RowComparisonRuns').getDataRange().getValues().slice(1).find(v=>v[0]===id);if(!r)throw new Error('Run missing');
 const folder=DriveApp.getFolderById(r[5]),files=folder.getFilesByName('source-rows.json');if(!files.hasNext())throw new Error('Source snapshot missing');
 return {folder,prompt:DriveApp.getFileById(r[6]).getBlob().getDataAsString(),settings:JSON.parse(r[7]),source:JSON.parse(files.next().getBlob().getDataAsString())};
}
function fcrPrepareJob(s,n,job,ctx) {
 const row=ctx.source.rows[job['Source row']-ctx.source.start],head=ctx.source.headers,read=h=>row[head.indexOf(h)];
 if(!read('Code File URL'))throw new Error('Saved old code URL missing; no Gemini request made.');
 const old=DriveApp.getFileById(fcFileId(read('Code File URL'))).getBlob().getDataAsString();if(!old.trim())throw new Error('Saved code is empty');
 const blob=DriveApp.getFileById(fcFileId(read('Upload your Flowchart'))).getBlob();
 if(!/^(application\/pdf|image\/(png|jpeg|webp|heic|heif))$/.test(blob.getContentType()))throw new Error('Unsupported flowchart MIME type '+blob.getContentType());
 const input=ctx.folder.createFile(blob).setName('row-'+job['Source row']+'-'+blob.getName()),code=ctx.folder.createFile('row-'+job['Source row']+'-old.py',old,'text/plain');
 const record=ctx.folder.createFile('row-'+job['Source row']+'-source.json',JSON.stringify({headers:head,values:row}),'application/json');
 fcrPut(s,n,{Status:'READY_TO_GENERATE','Old code URL':code.getUrl(),'Old code SHA256':fcHash(old),'Image snapshot URL':input.getUrl(),'Image SHA256':fcHash(blob.getBytes()),'Source snapshot URL':record.getUrl(),'Old API seconds (saved)':typeof read('Response Time (ms)')==='number'?read('Response Time (ms)')/1000:'','Old total tokens (saved)':fcrNumber(read('Total Tokens')),'Old input tokens (saved)':fcrNumber(read('Input Tokens')),'Old output tokens (saved)':fcrNumber(read('Output Tokens')),'Old thinking tokens (saved)':fcrNumber(read('Thoughts Token Count'))});
}
function fcrGenerateJob(p,s,n,job,ctx) {
 const blob=DriveApp.getFileById(fcFileId(job['Image snapshot URL'])).getBlob();
 if(fcHash(blob.getBytes())!==job['Image SHA256']||fcHash(ctx.prompt)!==job['Prompt SHA256'])throw new Error('Frozen input or prompt changed');
 fcrPut(s,n,{Status:'GENERATING','Started at':new Date()});
 const tick=Date.now(),resp=UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/'+encodeURIComponent(job['New model'])+':generateContent',{
  method:'post',contentType:'application/json',headers:{'x-goog-api-key':p.GEMINI_API_KEY},muteHttpExceptions:true,
  payload:JSON.stringify({contents:[{parts:[{inline_data:{mime_type:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())}},{text:ctx.prompt}]}],generationConfig:{temperature:0}})
 });
 const ms=Date.now()-tick,text=resp.getContentText(),raw=ctx.folder.createFile('row-'+job['Source row']+'-response.json',text,'application/json');
 fcrPut(s,n,{'Raw response URL':raw.getUrl(),'New API seconds':ms/1000});
 if(resp.getResponseCode()<200||resp.getResponseCode()>=300)throw new Error('Gemini HTTP '+resp.getResponseCode()+'; see Raw response URL. No automatic retry.');
 const json=JSON.parse(text),c=(json.candidates||[])[0]||{},u=json.usageMetadata||{};
 fcrPut(s,n,{'New input tokens':u.promptTokenCount??'','New output tokens':u.candidatesTokenCount??'','New thinking tokens':u.thoughtsTokenCount??'','New total tokens':u.totalTokenCount??'','Returned model':json.modelVersion||'','New estimated USD':fcCost(u,ctx.settings.rate)});
 if(c.finishReason!=='STOP')throw new Error('Incomplete output '+c.finishReason);
 const code=((c.content||{}).parts||[]).filter(v=>!v.thought&&typeof v.text==='string').map(v=>v.text).join('\n').replace(/^\s*```(?:python|py)?\s*\n/i,'').replace(/\n?```\s*$/,'').trim();
 if(!code)throw new Error('Empty generated code');
 const file=ctx.folder.createFile('row-'+job['Source row']+'-new.py',code,'text/plain');
 fcrPut(s,n,{'New code URL':file.getUrl(),'New code SHA256':fcHash(code),Status:'READY_TO_SUBMIT_OLD'});
}
function fcrSubmitJob(p,s,n,job,ctx,side) {
 const key=side==='Old'?'Old':'New';
 if(job[key+' dev submission'])throw new Error('Existing submission ID; duplicate submission blocked');
 const now=fcrJudge(p,ctx.settings,'/problems'),match=fcrMatch(job.Problem,now);
 if(!match||String(match.id)!==String(job['Judge problem ID'])||JSON.stringify(match)!==job['Judge metadata JSON'])throw new Error('Judge problem metadata changed; review before submitting');
 const code=DriveApp.getFileById(fcFileId(job[key+' code URL'])).getBlob().getDataAsString();
 if(fcHash(code)!==job[key+' code SHA256'])throw new Error(key+' code changed after snapshot');
 const zip=Utilities.zip([Utilities.newBlob(code,'text/x-python','solution.py')],'solution.zip');
 // Persist before POST. A timeout never leads to a second automatic POST.
 fcrPut(s,n,{Status:'SUBMITTING_'+key.toUpperCase(),'Started at':new Date()});
 const result=fcrJudge(p,ctx.settings,'/submissions',{problem_id:String(job['Judge problem ID']),language_id:ctx.settings.language,files:[{filename:'solution.zip',data:Utilities.base64Encode(zip.getBytes())}]});
 if(!result.id)throw new Error('Submission response has no ID; manual reconciliation required');
 const update={Status:key==='Old'?'READY_TO_SUBMIT_NEW':'POLLING','Started at':new Date()}; update[key+' dev submission']=String(result.id);fcrPut(s,n,update);
}
function fcrPollJob(p,s,n,job,ctx) {
 const update={};let ready=true;
 ['Old','New'].forEach(side=>{
  const id=job[side+' dev submission'];if(!id)throw new Error('Missing '+side+' submission ID');
  if(!job[side+' dev verdict']){const v=fcPickVerdict(fcrJudge(p,ctx.settings,'/judgements?submission_id='+encodeURIComponent(id)),id);if(v)update[side+' dev verdict']=v;else ready=false;}
 });
 if(ready){update.Status='DONE';update.Notes='Both exact saved/new code snapshots judged on the same dev problem. Historical generation prompt/tests not verified; one problem is not overall model accuracy. Cost is an estimate.';}
 else if(Date.now()-new Date(job['Started at']).getTime()>20*60000){update.Status='REVIEW_REQUIRED';update.Notes='Judge still pending after 20 minutes; inspect submission IDs. No resubmission made.';}
 fcrPut(s,n,update);
}
function fcrWorker() {
 const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
 try {
  const {ss,p}=fcrEnvironment(),props=PropertiesService.getScriptProperties(),id=props.getProperty('FCR_ACTIVE_RUN');
  if(!id){fcrTriggers(true);return;}
  const s=fcrTable(ss,'RowComparisons',FCR_COLS),ctx=fcrContext(ss,id),begin=Date.now();fcrRunState(ss,id,'RUNNING');
  for(let steps=0;steps<12&&Date.now()-begin<150000;steps++){
   if(props.getProperty('FCR_STOP_RUN')===id){fcrCancelQueued(ss,id);return;}
   const all=s.getDataRange().getValues(),index=all.findIndex((r,i)=>i>0&&r[0]===id&&!fcrDone(r[3]));
   if(index<0){fcrRunState(ss,id,'COMPLETE');props.deleteProperty('FCR_ACTIVE_RUN');fcrTriggers(true);return;}
   const job=fcrRecord(all[index]),n=index+1;
   try {
    if(/^(GENERATING|SUBMITTING_OLD|SUBMITTING_NEW)$/.test(job.Status)) {fcrPut(s,n,{Status:'REVIEW_REQUIRED',Notes:'Previous execution ended during an external request. Inspect raw response/submission history before any retry; duplicate calls blocked.'});continue;}
    if(job.Status==='QUEUED')fcrPrepareJob(s,n,job,ctx);
    else if(job.Status==='READY_TO_GENERATE'){fcrGenerateJob(p,s,n,job,ctx);return;}
    else if(job.Status==='READY_TO_SUBMIT_OLD')fcrSubmitJob(p,s,n,job,ctx,'Old');
    else if(job.Status==='READY_TO_SUBMIT_NEW')fcrSubmitJob(p,s,n,job,ctx,'New');
    else if(job.Status==='POLLING'){fcrPollJob(p,s,n,job,ctx);if(!fcrDone(fcrRecord(s.getRange(n,1,1,FCR_COLS.length).getValues()[0]).Status))return;}
    else throw new Error('Unknown state '+job.Status);
   }catch(e){const current=fcrRecord(s.getRange(n,1,1,FCR_COLS.length).getValues()[0]);fcrPut(s,n,{Status:/^SUBMITTING/.test(current.Status)?'REVIEW_REQUIRED':'ERROR',Notes:String(e.message||e)+' No automatic generation/submission retry.'});}
  }
 } finally {lock.releaseLock();}
}
