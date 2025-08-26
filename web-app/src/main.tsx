import { render } from 'solid-js/web';
import { createEffect, createMemo, createSignal, For, onCleanup, onMount, Show } from 'solid-js';
import './styles.css';

type CmdParam = { Name: string; Required: boolean; Type: string; Position: number };
type CmdItem = { Name: string; Category: string; Synopsis: string; Description: string; Examples: string[]; Parameters: CmdParam[]; Invocation: string; ScriptPath: string };
type Group = { Category: string; Commands: CmdItem[] };

const DEMO: Group[] = [
  { Category: 'Utilities', Commands: [
    { Name: 'Get-Thing', Category: 'Utilities', Synopsis: 'Gets a thing', Description: 'Retrieves a sample thing.', Examples: [
      'Get-Thing -Name alpha\nName : alpha\nId   : 1',
      'Get-Thing -Name beta\nName : beta\nId   : 2'
    ], Parameters: [
      { Name: 'Name', Required: true, Type: 'String', Position: 0 },
      { Name: 'Verbose', Required: false, Type: 'Switch', Position: 1 }
    ], Invocation: 'Get-Thing', ScriptPath: 'functions/Get-Thing.ps1' }
  ]},
  { Category: 'Networking', Commands: [
    { Name: 'Test-Ping', Category: 'Networking', Synopsis: 'Ping a host', Description: 'Pings a host and reports latency.', Examples: [
      'Test-Ping -Host example.com\nReply from 93.184.216.34: time=10ms'
    ], Parameters: [
      { Name: 'Host', Required: true, Type: 'String', Position: 0 },
      { Name: 'Count', Required: false, Type: 'Int32', Position: 1 },
    ], Invocation: 'Test-Ping', ScriptPath: 'functions/Test-Ping.ps1' }
  ]}
];

const SYM = { bullet: '\u2022' };
declare global { interface Window { __MODEL__?: Group[] } }
const initialModel: Group[] = window.__MODEL__ ?? [];

type RepoConfig = {
  owner: string; repo: string; branch: string; path: string;
  categoryDepth: number; ignore: string[]; token?: string;
  preferFolderCategory: boolean; // if true, use folder-derived category first
};

const DEFAULT_REPO: RepoConfig = {
  owner: 'roberto-ryan',
  repo: 'Public',
  branch: 'main',
  path: 'functions',
  categoryDepth: 2,
  ignore: ['functions','function','scripts','script','src','source','powershell','pwsh','ps','bin','build','.github','.vscode','lib','modules','module','samples','examples','test','tests','docs','documentation'],
  preferFolderCategory: true,
};

function getRepoConfigFromUrl(): RepoConfig {
  const p = new URLSearchParams(location.search);
  const repoParam = p.get('repo') || `${DEFAULT_REPO.owner}/${DEFAULT_REPO.repo}`;
  const [owner, repo] = repoParam.split('/') as [string,string];
  return {
    owner: owner || DEFAULT_REPO.owner,
    repo: repo || DEFAULT_REPO.repo,
    branch: p.get('branch') || DEFAULT_REPO.branch,
    path: (p.get('path') || DEFAULT_REPO.path).replace(/^\/+|\/+$|^\.\/?/g,'').replace(/\\/g,'/'),
    categoryDepth: Math.max(1, Math.min(5, parseInt(p.get('depth') || `${DEFAULT_REPO.categoryDepth}`) || DEFAULT_REPO.categoryDepth)),
  ignore: (p.get('ignore') || DEFAULT_REPO.ignore.join(',')).split(',').map(s=>s.trim()).filter(Boolean),
  token: p.get('token') || undefined,
  preferFolderCategory: (p.get('preferFolder') ?? '1') !== '0',
  };
}

type GithubContent = { type: 'file'|'dir'; path: string; name: string; download_url?: string; };

async function listRepoTree(cfg: RepoConfig, subPath: string): Promise<GithubContent[]> {
  const url = `https://api.github.com/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}/contents/${encodeURIComponent(subPath)}?ref=${encodeURIComponent(cfg.branch)}`;
  const res = await fetch(url, { headers: { 'Accept': 'application/vnd.github+json' } });
  if (!res.ok) throw new Error(`GitHub list failed ${res.status}: ${await res.text()}`);
  const data = await res.json();
  return Array.isArray(data) ? data as GithubContent[] : [];
}

async function collectPs1Files(cfg: RepoConfig): Promise<GithubContent[]> {
  const out: GithubContent[] = [];
  const queue: string[] = [cfg.path.replace(/^\/+|\/+$/g,'')];
  while (queue.length) {
    const cur = queue.shift()!;
    try {
  const url = `https://api.github.com/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}/contents/${encodeURIComponent(cur)}?ref=${encodeURIComponent(cfg.branch)}`;
  const headers: Record<string,string> = { 'Accept': 'application/vnd.github+json' };
  if (cfg.token) headers['Authorization'] = `Bearer ${cfg.token}`;
  const res = await fetch(url, { headers });
  if (!res.ok) throw new Error(`GitHub list failed ${res.status}`);
  const items = (await res.json()) as GithubContent[];
      for (const it of items) {
        if (it.type === 'dir') queue.push(it.path);
        else if (it.type === 'file' && it.name.toLowerCase().endsWith('.ps1')) out.push(it);
      }
    } catch (e) {
      // If a directory is missing or inaccessible, continue
      // eslint-disable-next-line no-console
      console.warn('List failed for', cur, e);
    }
  }
  return out;
}

// Tiny concurrency limiter
async function mapLimit<T, R>(arr: T[], limit: number, fn: (t: T, i: number) => Promise<R>): Promise<R[]> {
  const ret: R[] = new Array(arr.length);
  let next = 0; let active = 0; let resolve!: (v: R[]) => void; let reject!: (e: any) => void;
  const done = new Promise<R[]>((res, rej) => { resolve = res; reject = rej; });
  const kick = () => {
    while (active < limit && next < arr.length) {
      const i = next++; active++;
      Promise.resolve(fn(arr[i], i)).then(v => { ret[i] = v; active--; if (next>=arr.length && active===0) resolve(ret); else kick(); }, err => { reject(err); });
    }
  };
  kick();
  return done;
}

function parseHelpBlock(text: string) {
  const m = text.match(/<#([\s\S]*?)#>/m);
  if (!m) return {} as any;
  const body = m[1];

  // Extract .PARAMETER names directly from the header line to avoid grabbing the first word of the description (e.g., "The ...")
  const isJunkHelpParam = (n: string) => /^(none|no|parameters?|true|false|n\/a|na)$/i.test(n);
  const paramHelpNames: string[] = Array.from(body.matchAll(/^\s*\.PARAMETER\s+([^\r\n]+)/gim))
    .map(x => x[1].trim().split(/\s+/)[0])
    .filter(Boolean)
    .filter(n => !isJunkHelpParam(n));

  // Build simple section map for SYNOPSIS/DESCRIPTION/EXAMPLE rendering
  const lines = body.split(/\r?\n/);
  let current: string | null = null; const sections: Record<string,string[]> = {};
  for (let raw of lines) {
    const trimmed = raw.trim();
    const head = trimmed.match(/^\.(\w+)/);
    if (head) { current = head[1].toUpperCase(); if (!sections[current]) sections[current] = []; continue; }
    if (current) sections[current].push(raw.replace(/\s+$/,''));
  }
  const syn = (sections['SYNOPSIS']||[]).join('\n').trim();
  const desc = (sections['DESCRIPTION']||[]).join('\n').trim();
  const exs: string[] = [];
  const keys = Object.keys(sections).filter(k=>k==='EXAMPLE' || k.startsWith('EXAMPLE'));
  for (const k of keys) {
    const t = (sections[k]||[]).join('\n').trim(); if (t) exs.push(t);
  }
  return { synopsis: syn, description: desc, examples: exs, paramHelpNames };
}

// Robustly parse PowerShell param(...) block including attributes
function parseParamBlock(text: string) {
  // Find the first param( ... ) and extract the balanced contents
  const paramStart = text.search(/param\s*\(/i);
  const params: CmdParam[] = [];
  if (paramStart < 0) return params;
  const openIdx = text.indexOf('(', paramStart);
  if (openIdx < 0) return params;
  let i = openIdx + 1;
  let depth = 1;
  let inStr: '"' | "'" | null = null;
  for (; i < text.length; i++) {
    const ch = text[i];
    const prev = text[i - 1];
    if (inStr) {
      if (ch === inStr) {
        // Handle escaped quotes in PowerShell: doubled quotes inside same-quoted string
        if (inStr === '"' && prev === '`') { /* escaped double quote with backtick */ }
        else if (text[i + 1] === inStr) { i++; /* doubled quote escape */ }
        else inStr = null;
      }
      continue;
    } else {
      if (ch === '"' || ch === "'") { inStr = ch as any; continue; }
      if (ch === '(') depth++;
      else if (ch === ')') { depth--; if (depth === 0) { break; } }
    }
  }
  const block = text.slice(openIdx + 1, i);

  // Split by commas at top-level (not inside []/() or strings)
  const chunks: string[] = [];
  let cur = '';
  let pDepth = 0; // () depth
  let bDepth = 0; // [] depth
  inStr = null;
  for (let j = 0; j < block.length; j++) {
    const ch = block[j];
    const prev = block[j - 1];
    if (inStr) {
      cur += ch;
      if (ch === inStr) {
        if (inStr === '"' && prev === '`') { /* escaped */ }
        else if (block[j + 1] === inStr) { cur += block[j + 1]; j++; }
        else inStr = null;
      }
      continue;
    }
    if (ch === '"' || ch === "'") { inStr = ch as any; cur += ch; continue; }
    if (ch === '[') { bDepth++; cur += ch; continue; }
    if (ch === ']') { bDepth = Math.max(0, bDepth - 1); cur += ch; continue; }
    if (ch === '(') { pDepth++; cur += ch; continue; }
    if (ch === ')') { pDepth = Math.max(0, pDepth - 1); cur += ch; continue; }
    if (ch === ',' && pDepth === 0 && bDepth === 0) { if (cur.trim()) chunks.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) chunks.push(cur.trim());

  // Now parse each parameter chunk, capturing attributes before the variable
  for (const ch of chunks) {
    // Find the parameter variable name; skip special tokens like $true/$false/$null and $env drive
    const allVars = Array.from(ch.matchAll(/\$([A-Za-z_][\w]*)\b/g));
    if (allVars.length === 0) continue;
  const isSpecialVar = (n: string) => /^(true|false|null|env|_|psitem)$/i.test(n);
    const chosen = allVars.find(m => !isSpecialVar(m[1]));
    if (!chosen) continue;
    const name = chosen[1];
    const head = ch.slice(0, chosen.index!); // text before $name, includes [Parameter(...)] [Validate*] [type]
    const bracketMatches = Array.from(head.matchAll(/\[([^\]]+)\]/g)).map(m => m[1]);

    // Determine Required based on any [Parameter(Mandatory=$true)]
  let required = false;
  let position = -1; // default to non-positional unless explicitly set
    let detectedType = '';
    for (const b of bracketMatches) {
      const s = b.trim();
      const lower = s.toLowerCase();
      if (lower.startsWith('parameter(')) {
        const body = s.slice(s.indexOf('(') + 1, s.lastIndexOf(')'));
        // Support both explicit and shorthand Mandatory forms:
        // - Mandatory=$true / $false
        // - Mandatory (implies $true)
        const mandEq = body.match(/mandatory\s*=\s*\$?(true|false)/i);
        if (mandEq) {
          required = mandEq[1].toLowerCase() === 'true';
        } else if (/\bmandatory\b/i.test(body)) {
          required = true;
        }
        const posM = body.match(/position\s*=\s*(-?\d+)/i);
        if (posM) position = parseInt(posM[1]);
      } else if (!/^validate\w*\s*\(/i.test(s) && !/^alias\s*\(/i.test(s)) {
        // Treat the last non-Parameter/Validate/Alias block as the type
        detectedType = s;
      }
    }
    const tSimple = /switch/i.test(detectedType) ? 'Switch'
      : /bool/i.test(detectedType) ? 'Boolean'
      : /int(16|32|64)?/i.test(detectedType) ? 'Int32'
      : /string/i.test(detectedType) ? 'String'
      : detectedType || 'Object';
    params.push({ Name: name, Required: required, Type: tSimple as any, Position: position });
  }
  return params;
}

// Regex override map, modeled after CLI defaults
const CATEGORY_OVERRIDE_MAP: Array<[RegExp, string]> = [
  [/\b(M365|O365|Office365|ExchangeOnline|SharePointOnline|OneDrive|Teams|Outlook|Planner|PowerPlatform|PowerApps|PowerAutomate)\b/i, 'Microsoft 365'],
  [/\b(ActiveDirectory|ADFS|ADDS|DomainController|SAMAccount|Kerberos|LDAP|AAD|AzureAD|EntraID)\b/i, 'Identity'],
  [/\b(Security|SecPol|ACL|Permissions|Credential|Secrets|PKI|Certificate|TLS|SSL|Firewall|Malware|Virus|Defender|AppLocker|BitLocker|Encryption)\b/i, 'Security'],
  [/\b(Intune|SCCM|ConfigMgr|EndpointManager|MDM|GPO|GroupPolicy)\b/i, 'Endpoint Mgmt'],
  [/\b(Azure|ARM|ResourceGroup|VMSS|AKS|AppService|KeyVault|StorageAccount|CosmosDB|LogicApp|FunctionApp|VNet|NSG)\b/i, 'Azure'],
  [/\b(Jira|Confluence|DevOps|Agile|Scrum|Kanban|Trello|Asana)\b/i, 'Dev/Work Mgmt'],
  [/\b(Hyper-V|VMware|vSphere|ESXi|VirtualBox|VHDX?|Snapshot|Checkpoint)\b/i, 'Virtualization'],
  [/\b(Docker|Podman|Containerd|K8s|Kubernetes|Helm|Image|Container)\b/i, 'Containers'],
  [/\b(DevOps|Terraform|Ansible|Chef|Puppet|Bicep|ARMTemplate|CI\/CD|Pipeline|Jenkins|Octopus)\b/i, 'DevOps/IaC'],
  [/\b(SQLServer|MSSQL|Postgres|PostgreSQL|MySQL|MariaDB|OracleDB|MongoDB|Redis|Database|SQLite)\b/i, 'Databases'],
  [/\b(OAuth|OIDC|SAML|JWT|FIDO2|MFA|2FA|SSO|AuthN|AuthZ)\b/i, 'Auth Standards'],
  [/\b(AWS|AmazonWebServices|EC2|S3|GCP|GoogleCloud|BigQuery|CloudRun)\b/i, 'Cloud Vendors'],
  [/\b(Network|NetCfg|DNS|DHCP|IPConfig|Ping|Traceroute|Subnet|Routing|Switch|Router|WiFi|NAT|Port|TCP|UDP|SSLVPN|VPN)\b/i, 'Networking'],
  [/\b(PowerShell|Bash|Python|CSharp|C#|JavaScript|TypeScript|GoLang?|Rust|Perl|Ruby)\b/i, 'Languages'],
  [/\b(Backup|Restore|Recovery|Snapshot|Replication|Failover|DisasterRecovery|DR|Veeam)\b/i, 'Backup and DR'],
  [/\b(Exchange|SMTP|IMAP|POP3|Mailbox|MailFlow|SendMail)\b/i, 'Email'],
  [/\b(Help|Get-Help|About_|Info|Discover|WhatIf)\b/i, 'Help/Discovery'],
  [/\b(Process|Tasklist|Taskkill|Get-Process|ProcMon|Handle)\b/i, 'Processes'],
  [/\b(Service|Get-Service|Set-Service|Start-Service|Stop-Service|Restart-Service)\b/i, 'Services'],
  [/\b(File|Folder|Directory|Path|Copy-Item|Move-Item|Remove-Item|Rename-Item|New-Item|Get-ChildItem|Tree)\b/i, 'Files/Directories'],
  [/\b(Registry|RegKey|HKLM|HKCU|HKCR|HKU|HKCC|Get-ItemProperty|Set-ItemProperty)\b/i, 'Registry'],
  [/\b(EventLog|Get-EventLog|Get-WinEvent|LogName|ApplicationLog|SystemLog|Audit)\b/i, 'Events/Logs'],
  [/\b(Device|PnP|Driver|Hardware|Disk|Volume|Partition|USB|PrinterPort|Monitor|Battery|Adapter)\b/i, 'Hardware/Devices'],
  [/\b(Print|Printer|PrintJob|Spooler|PrintQueue)\b/i, 'Printing'],
  [/\b(Update|Patch|WUInstall|WindowsUpdate|KB\d+)\b/i, 'Updates'],
  [/\b(User|Group|LocalUser|LocalGroup|Account|SID|Profile|Credential|NTUser)\b/i, 'Users/Groups'],
  [/\b(PerfMon|Performance|Counter|ResourceMonitor|CPU|Memory|DiskIO|Latency|Benchmark)\b/i, 'Performance'],
  [/\b(PowerPlan|Battery|Sleep|Hibernate|Shutdown|Restart|Reboot|UPS)\b/i, 'Power'],
  [/\b(Display|Resolution|Monitor|Screen|Graphics|DPI|Color|Brightness)\b/i, 'Display'],
  [/\b(Keyboard|Mouse|Input|HID|Touchpad|Tablet|Pen)\b/i, 'Input'],
  [/\b(Audio|Sound|Speaker|Microphone|Mute|Volume)\b/i, 'Audio'],
  [/\b(Troubleshoot|Diag|Diagnosis|Fix|Repair|Checkup|Health|SFC|DISM)\b/i, 'Troubleshooting'],
  [/\b(Install|Setup|Deployment|Sysprep|ImageX|WIM|ISO|Provisioning)\b/i, 'Installation'],
  [/\b(Recovery|WinRE|Reset|RestorePoint|SystemRestore|BootRepair)\b/i, 'Recovery'],
  [/\b(Util|Utility|Tool|Script|Helper|AdminTool|Sysinternals)\b/i, 'Utilities'],
];

function deriveCategory(cfg: RepoConfig, filePath: string, content?: string): string {
  const normPath = filePath.replace(/\\/g,'/');
  const base = cfg.path.replace(/\\/g,'/').replace(/^\/+|\/+$/g,'');
  const idx = normPath.toLowerCase().indexOf(base.toLowerCase());
  const rel = idx >= 0 ? normPath.slice(idx + base.length).replace(/^\/+/, '') : normPath;
  const parts = rel.split('/').slice(0, -1).filter(Boolean).filter(p => !cfg.ignore.some(ig => ig.toLowerCase() === p.toLowerCase()));
  const taken = parts.slice(0, cfg.categoryDepth);
  const folderCat = taken.length ? taken.join(' / ') : '';
  // Regex override (prefer if match exists and preferFolderCategory=false)
  let regexCat = '';
  const probe = `${filePath}\n${content || ''}`;
  for (const [rx, label] of CATEGORY_OVERRIDE_MAP) { if (rx.test(probe)) { regexCat = label; break; } }
  if (cfg.preferFolderCategory) return folderCat || regexCat || 'General';
  return regexCat || folderCat || 'General';
}

function parsePowerShellFile(cfg: RepoConfig, text: string, ghPath: string): CmdItem | null {
  const nameMatch = text.match(/^[\s\S]{0,2000}?function\s+([A-Za-z0-9_-]+)\b/m); // search early area
  const name = nameMatch?.[1] || ghPath.split('/').pop()!.replace(/\.ps1$/i,'');
  const help = parseHelpBlock(text);
  const paramsFromBlock = parseParamBlock(text);
  // If help had parameter names but types unknown, merge
  const pmap = new Map(paramsFromBlock.map(p => [p.Name.toLowerCase(), p] as const));
  const hasVar = (nm: string) => new RegExp(`\\$${nm}\\b`, 'i').test(text);
  for (const h of help.paramHelpNames || []) {
    const key = h.toLowerCase();
    // Only add help-only params if there's a matching $variable in the script
    if (!pmap.has(key) && hasVar(h)) pmap.set(key, { Name: h, Required: false, Type: 'String', Position: 0 });
  }
  const params = Array.from(pmap.values()).sort((a,b) => a.Position - b.Position);
  return {
    Name: name,
  Category: deriveCategory(cfg, ghPath, text),
    Synopsis: help.synopsis || '',
    Description: help.description || '',
    Examples: help.examples || [],
    Parameters: params,
    Invocation: name,
    ScriptPath: ghPath
  };
}

async function getFileContent(cfg: RepoConfig, f: GithubContent): Promise<string | null> {
  try {
    if (cfg.token) {
      const url = `https://api.github.com/repos/${encodeURIComponent(cfg.owner)}/${encodeURIComponent(cfg.repo)}/contents/${encodeURIComponent(f.path)}?ref=${encodeURIComponent(cfg.branch)}`;
      const res = await fetch(url, { headers: { 'Accept': 'application/vnd.github+json', 'Authorization': `Bearer ${cfg.token}` } });
      if (!res.ok) return null;
      const data = await res.json();
      if (data && typeof data.content === 'string') {
        try { return atob(data.content.replace(/\n/g,'')); } catch { return null; }
      }
      return null;
    }
    if (!f.download_url) return null;
    const res = await fetch(f.download_url);
    if (!res.ok) return null;
    return await res.text();
  } catch { return null; }
}

async function loadModelFromGithub(cfg: RepoConfig): Promise<{ groups: Group[]; title: string; }> {
  const files = await collectPs1Files(cfg);
  // Fetch contents with limited concurrency to be gentle on API
  const items = await mapLimit(files, 6, async (f) => {
    const text = await getFileContent(cfg, f);
    if (!text) return null;
    return parsePowerShellFile(cfg, text, f.path);
  });
  const valid = (items.filter(Boolean) as CmdItem[]);
  // Group by category
  const byCat = new Map<string, CmdItem[]>();
  for (const it of valid) {
    const key = it.Category || 'General';
    if (!byCat.has(key)) byCat.set(key, []);
    byCat.get(key)!.push(it);
  }
  const groups: Group[] = Array.from(byCat.entries()).sort((a,b)=>a[0].localeCompare(b[0])).map(([Category, Commands]) => ({ Category, Commands: Commands.sort((a,b)=>a.Name.localeCompare(b.Name)) }));
  const title = `${cfg.owner}/${cfg.repo}@${cfg.branch} • ${cfg.path}`;
  return { groups, title };
}

function App() {
  const cfg = getRepoConfigFromUrl();
  const [groups, setGroups] = createSignal<Group[]>(initialModel);
  const [title, setTitle] = createSignal<string>('Loading…');
  const [error, setError] = createSignal<string>('');
  const [catIdx, setCatIdx] = createSignal(0);
  const [cmdIdx, setCmdIdx] = createSignal(0);
  const [focus, setFocus] = createSignal<'cat'|'cmd'|'detail'>('cat');
  const [footer, setFooter] = createSignal('Left/Right: change pane  |  Up/Down: select or scroll  |  Enter: build/copy command  |  Ctrl+C: exit');
  const [toast, setToast] = createSignal('');
  const [paramValues, setParamValues] = createSignal<Record<string, any>>({});
  const [paramIdx, setParamIdx] = createSignal(0);
  type PromptMode = 'off' | 'param' | 'summary';
  const [promptMode, setPromptMode] = createSignal<PromptMode>('off');
  const [promptIdx, setPromptIdx] = createSignal(0);
  const [promptBuffer, setPromptBuffer] = createSignal('');
  const baseFooter = ' Left/Right: change pane  |  Up/Down: select or scroll  |  Enter: build/copy  |  Ctrl+C: exit ';
  setFooter(baseFooter);

  const cmds = createMemo(() => groups()[catIdx()]?.Commands ?? []);
  const selected = createMemo(() => cmds()[cmdIdx()]);

  // Scroll containers per pane
  let catsListEl: HTMLDivElement | undefined;
  let cmdsListEl: HTMLDivElement | undefined;
  let detailListEl: HTMLDivElement | undefined;
  const scrollDetail = (dy: number) => {
    if (detailListEl) detailListEl.scrollBy({ top: dy, behavior: 'auto' });
  };
  const ensureSelectedVisible = (list: HTMLDivElement | undefined, index: number) => {
    // Defer until DOM updates apply the selected class
    requestAnimationFrame(() => {
      if (!list || index == null || index < 0) return;
      const rows = list.querySelectorAll<HTMLDivElement>('.row');
      const el = rows[index];
      if (el && typeof el.scrollIntoView === 'function') {
        el.scrollIntoView({ block: 'nearest' });
      }
    });
  };
  const onKey = (e: KeyboardEvent) => {
    const t = e.target as HTMLElement | null;
    const tag = (t?.tagName || '').toLowerCase();
  // Wizard has highest priority
    if (promptMode() !== 'off') {
      const s = selected();
      const params = s?.Parameters || [];
      const idx = promptIdx();
      if (promptMode() === 'param' && s && params[idx]) {
        const p = params[idx];
        const type = p.Type || 'String';
        const isSwitch = /switch|boolean/i.test(type);
        if (isSwitch) {
          if (e.key === ' ') {
            const cur = !!paramValues()[p.Name];
            setParamValues({ ...paramValues(), [p.Name]: !cur });
            e.preventDefault();
            return;
          }
          if (e.key === 'Enter') {
            // accept and next
            const next = idx + 1;
            if (next < params.length) { setPromptIdx(next); }
            else {
              const s2 = selected();
              if (s2) {
                const miss = getMissingRequiredIndices(s2, paramValues());
                if (miss.length) { setPromptIdx(miss[0]); setPromptMode('param'); }
                else { setPromptMode('summary'); }
              } else { setPromptMode('summary'); }
            }
            return;
          }
          if (e.key === 'Escape') { setPromptMode('off'); return; }
          // ignore others
          return;
        } else {
          if (e.key === 'Enter') {
            const v = promptBuffer();
            const t = v.trim();
            if (t) setParamValues({ ...paramValues(), [p.Name]: t }); else { const vals = { ...paramValues() }; delete vals[p.Name]; setParamValues(vals); }
            setPromptBuffer('');
            const next = idx + 1;
            if (next < params.length) { setPromptIdx(next); }
            else {
              const s2 = selected();
              if (s2) {
                const miss = getMissingRequiredIndices(s2, paramValues());
                if (miss.length) { setPromptIdx(miss[0]); setPromptMode('param'); }
                else { setPromptMode('summary'); }
              } else { setPromptMode('summary'); }
            }
            return;
          }
          if (e.key === 'Escape') { setPromptMode('off'); setPromptBuffer(''); return; }
          if (e.key === 'Backspace' || e.key === 'Delete') { setPromptBuffer(promptBuffer().slice(0,-1)); e.preventDefault(); return; }
          if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { setPromptBuffer(promptBuffer() + e.key); e.preventDefault(); return; }
          return;
        }
      }
      if (promptMode() === 'summary') {
          if (e.key === 'Enter') { const s = selected(); if (s) { copyToClipboard(buildPreviewCommand(s)); showToast('Command copied to clipboard'); } setPromptMode('off'); return; }
        if (e.key === 'Escape') { setPromptMode('off'); return; }
        return;
      }
    }

  const typing = false;
    switch (e.key) {
      case 'Enter': {
        const s = selected();
        if (s) {
          const params = s.Parameters || [];
          if (params.length === 0) {
            // Even if there are no parameters, open the wizard to the final summary screen
            setPromptMode('summary');
          } else {
            setPromptIdx(0);
            const first = params[0];
            if (/switch|boolean/i.test(first.Type || '')) setPromptBuffer('');
            else setPromptBuffer(String(paramValues()[first.Name] ?? ''));
            setPromptMode('param');
            setFooter(' Enter: accept  |  Space: toggle  |  Esc: cancel ');
          }
        }
        break;
      }
      case 'ArrowUp': {
        if (typing) { e.preventDefault(); return; }
        if (focus()==='cat') { if (catIdx()>0) setCatIdx(catIdx()-1); e.preventDefault(); }
        else if (focus()==='cmd') { if (cmdIdx()>0) setCmdIdx(cmdIdx()-1); e.preventDefault(); }
        else if (focus()==='detail') { scrollDetail(-24); e.preventDefault(); }
        break;
      }
      case 'ArrowDown': {
        if (typing) { e.preventDefault(); return; }
        if (focus()==='cat') { if (catIdx()<groups().length-1) setCatIdx(catIdx()+1); e.preventDefault(); }
        else if (focus()==='cmd') { if (cmdIdx()<cmds().length-1) setCmdIdx(cmdIdx()+1); e.preventDefault(); }
        else if (focus()==='detail') { scrollDetail(24); e.preventDefault(); }
        break;
      }
      case 'PageUp': { if (!typing && focus()==='detail') { scrollDetail(-(detailListEl?.clientHeight||200)*0.85); e.preventDefault(); } break; }
      case 'PageDown': { if (!typing && focus()==='detail') { scrollDetail((detailListEl?.clientHeight||200)*0.85); e.preventDefault(); } break; }
      case 'Home': { if (!typing && focus()==='detail' && detailListEl) { detailListEl.scrollTo({ top: 0, behavior: 'auto' }); e.preventDefault(); } break; }
      case 'End': { if (!typing && focus()==='detail' && detailListEl) { detailListEl.scrollTo({ top: detailListEl.scrollHeight, behavior: 'auto' }); e.preventDefault(); } break; }
  case 'ArrowLeft': { if (focus()==='cmd') setFocus('cat'); else if (focus()==='detail') setFocus('cmd'); e.preventDefault(); break; }
  case 'ArrowRight': { if (focus()==='cat') setFocus('cmd'); else if (focus()==='cmd') setFocus('detail'); e.preventDefault(); break; }
      // No other editing paths outside wizard
      default: {
        // no-op
      }
    }
  };
  onMount(() => {
    window.addEventListener('keydown', onKey);
    // Kick off GitHub load
    (async () => {
      try {
        const { groups, title } = await loadModelFromGithub(cfg);
        if (groups.length === 0) { setGroups(DEMO); setTitle('Demo Data'); }
        else { setGroups(groups); setTitle(title); }
      } catch (e: any) {
        setError(e?.message || String(e));
        setGroups(DEMO); setTitle('Demo Data (GitHub load failed)');
      }
  setCatIdx(0); setCmdIdx(0);
    })();
  });
  onCleanup(() => window.removeEventListener('keydown', onKey));

  createEffect(() => { cmds(); setCmdIdx(0); });

  // Keep selection visible as it changes
  createEffect(() => { groups(); ensureSelectedVisible(catsListEl, catIdx()); });
  createEffect(() => { cmds(); ensureSelectedVisible(cmdsListEl, cmdIdx()); });

  // When selected command changes, reset parameter values
  createEffect(() => {
    const s = selected();
    if (!s) return;
    const init: Record<string, any> = {};
    for (const p of s.Parameters || []) {
      const key = p.Name;
      if (/switch|boolean/i.test(p.Type)) init[key] = false;
      else init[key] = '';
    }
    setParamValues(init);
  setParamIdx(0);
  
  });

  function psQuote(val: string) {
    const s = String(val);
    // Single-quote and escape single quotes as ''
    return `'${s.replace(/'/g, "''")}'`;
  }
  function buildPreviewCommand(cmd: CmdItem) {
    const base = cmd.Invocation || cmd.Name;
    const provided = paramValues();
    const named: string[] = [];
    type PosEnt = { pos: number; text: string };
    const positional: PosEnt[] = [];
    const params = cmd.Parameters || [];
    for (const p of params) {
      const key = p.Name;
      let val = provided[key];
      const isSwitch = /switch/i.test(p.Type);
      const isBool = /boolean/i.test(p.Type);
      const hasVal = isSwitch ? !!val : (val !== undefined && val !== null && String(val).trim() !== '');
      if (!hasVal) continue;
      if (isSwitch) {
        named.push(`-${key}`);
        continue;
      }
      const t = /int|double|float|decimal/i.test(p.Type)
        ? String(val)
        : psQuote(String(val));
      if (typeof p.Position === 'number' && p.Position >= 0) positional.push({ pos: p.Position, text: t });
      else named.push(`-${key} ${t}`);
    }
    positional.sort((a,b) => a.pos - b.pos);
    const parts = [base, ...positional.map(x=>x.text), ...named];
    return parts.join(' ');
  }
  function getMissingRequiredIndices(cmd: CmdItem, values: Record<string, any>): number[] {
    const miss: number[] = [];
    const params = cmd.Parameters || [];
    for (let i=0;i<params.length;i++) {
      const p = params[i];
      if (!p.Required) continue;
      const type = p.Type || 'String';
      const val = values[p.Name];
      if (/switch|boolean/i.test(type)) { if (!val) miss.push(i); }
      else { if (val === undefined || val === null || String(val).trim() === '') miss.push(i); }
    }
    return miss;
  }
  function copyToClipboard(text: string) {
    navigator.clipboard?.writeText(text).catch(() => {
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta);
    });
  }
  function showToast(msg: string) { setToast(msg); setTimeout(() => setToast(''), 1500); }

  createEffect(() => {
  if (promptMode() === 'off') setFooter(baseFooter);
  });

  const formatExamples = (exs: string[] = []) => {
    const out: any[] = [];
    exs.slice(0,3).forEach((raw, i) => {
      if (i>0) out.push(<div class="code">{"\n"}</div>);
      const lines = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
      if (lines.length>0) {
        out.push(<div class="code"><span class="bullet">{SYM.bullet}</span>{lines[0]}</div>);
        out.push(<div class="code">{"\n"}</div>);
      }
      for (let j=1; j<lines.length; j++) { out.push(<div class="code">{'  ' + lines[j].replace(/\s+$/,'')}</div>); }
    });
    return out;
  };

  return (
    <div class="app">
  <div class="header">VTS Tools Browser</div>
      <div class="main">
        <div class={"pane pane-cats "+(focus()==='cat'?'active':'')}>
          <div class="pane-title">Categories</div>
          <div class="list" ref={(el) => (catsListEl = el as HTMLDivElement)}>
            <Show when={error()}>
              <div class="row warn">{error()}</div>
            </Show>
            <For each={groups()}>{(g, i) => (
              <div class={'row '+(i()==catIdx()?'selected':'')}>
                <span class="bullet">{SYM.bullet}</span>{g.Category} ({g.Commands?.length ?? 0})
              </div>
            )}</For>
          </div>
        </div>
  <div class={"pane pane-cmds "+(focus()==='cmd'?'active':'')}>
          <div class="pane-title">Commands</div>
          <div class="list" ref={(el) => (cmdsListEl = el as HTMLDivElement)}>
            <For each={cmds()}>{(c, i) => (
              <div class={'row '+(i()==cmdIdx()?'selected':'')}>
                <span class="bullet">{SYM.bullet}</span>{c.Name}
              </div>
            )}</For>
          </div>
        </div>
        <div class={"pane "+(focus()==='detail'?'active':'')}>
          <div class="pane-title">Details</div>
          <div class="list" ref={(el) => (detailListEl = el as HTMLDivElement)}>
            <Show when={selected()}>
              {(s) => (
                <div>
                  <div class="code"><span class="section">Name: </span>{s().Name}</div>
                  <div class="code"><span class="section">Category: </span>{s().Category}</div>
                  <div class="code"><span class="section">Path: </span>{s().ScriptPath}</div>
                  <Show when={s().Parameters?.length}>
                    <div class="section">Build Command:</div>
                    <div class="code preview-cmd">{buildPreviewCommand(s())}</div>
                  </Show>
                  <Show when={s().Synopsis}><div class="section">Synopsis:</div></Show>
                  <Show when={s().Synopsis}><div class="code">{s().Synopsis}</div></Show>
                  <Show when={s().Description}><div class="section">Description:</div></Show>
                  <Show when={s().Description}><div class="code">{s().Description}</div></Show>
                  <Show when={s().Parameters?.length}> 
                    <div class="section-subtle">Press Enter to start the parameter wizard</div>
                    <For each={s().Parameters}>{(p, i) => {
                      const type = p.Type || 'String';
                      const val = paramValues()[p.Name];
                      const isSwitch = /switch|boolean/i.test(type);
                      const isOn = isSwitch && !!val;
                      const displayVal = isSwitch ? (isOn ? '[On]' : '[Off]') : (val ? String(val) : (p.Required ? '<required>' : ''));
                      return (
                        <div class={'code param-item ' + (i()==paramIdx() ? 'selected' : '')}>
                          {SYM.bullet} {p.Name} &lt;{type}&gt; {p.Required ? '(required)' : '(optional)'} {displayVal ? '— '+displayVal : ''}
                        </div>
                      );
                    }}</For>
                  </Show>
                  
                  <Show when={s().Examples?.length}>
                    <div class="section">Examples:</div>
                    {formatExamples(s().Examples)}
                  </Show>
                </div>
              )}
            </Show>
          </div>
        </div>
      </div>
      <div class="footer">{footer()}</div>
      <Show when={promptMode() !== 'off'}>
        <div class="prompt-overlay">
          <div class="prompt">
            <Show when={promptMode()==='param'}>
              <div>
                <div class="prompt-title">{selected()?.Name}</div>
                <div class="prompt-sub">Parameter {promptIdx()+1} of {selected()?.Parameters?.length || 0}</div>
                <div class="prompt-body">
                  <div class="code">
                    <span>{(() => {
                      const s = selected();
                      const p = s?.Parameters?.[promptIdx()];
                      if (!p) return '';
                      const type = p.Type || 'String';
                      if (/switch|boolean/i.test(type)) {
                        const on = !!paramValues()[p.Name];
                        return `${p.Name} <${type}> ${p.Required?'(required)':'(optional)'} — ${on ? '[On]' : '[Off]'}\n\nSpace: toggle • Enter: next • Esc: cancel`;
                      }
                      return `${p.Name} <${type}> ${p.Required?'(required)':'(optional)'}\n\n${promptBuffer()}_\n\nType to enter • Enter: next • Esc: cancel`;
                    })()}</span>
                  </div>
                </div>
              </div>
            </Show>
            <Show when={promptMode()==='summary'}>
              <div>
                <div class="prompt-title">Command Ready</div>
                <div class="prompt-body">
                  <div class="code">{selected() ? buildPreviewCommand(selected()!) : ''}</div>
                  <div class="section-subtle">Enter: copy • Esc: back</div>
                </div>
              </div>
            </Show>
          </div>
        </div>
      </Show>
      <div class={"copy-toast "+(toast()? 'show':'')}>{toast()}</div>
    </div>
  );
}

render(() => <App />, document.getElementById('root')!);
