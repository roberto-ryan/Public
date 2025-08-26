import { render } from 'solid-js/web';
import { createEffect, createMemo, createSignal, For, onCleanup, onMount, Show } from 'solid-js';

// Types
type CmdParam = { Name: string; Required: boolean; Type: string; Position: number };
type CmdItem = { Name: string; Category: string; Synopsis: string; Description: string; Examples: string[]; Parameters: CmdParam[]; Invocation: string; ScriptPath: string };
type Group = { Category: string; Commands: CmdItem[] };

// Demo data (placeholder) — this is replaced at runtime via window.__MODEL__ if present
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

// Symbols
const SYM = {
  bullet: '\u2022',
};

// Read initial model if injected
declare global { interface Window { __MODEL__?: Group[] } }
const initialModel: Group[] = window.__MODEL__ ?? DEMO;

function App() {
  const [groups, setGroups] = createSignal<Group[]>(initialModel);
  const [catIdx, setCatIdx] = createSignal(0);
  const [cmdIdx, setCmdIdx] = createSignal(0);
  const [focus, setFocus] = createSignal<'cat'|'cmd'|'detail'>('cat');
  const [footer, setFooter] = createSignal("Left/Right: change pane  |  Up/Down: select or scroll  |  Enter: copy command  |  Ctrl+C: exit");
  const [toast, setToast] = createSignal('');

  const cmds = createMemo(() => groups()[catIdx()]?.Commands ?? []);
  const selected = createMemo(() => cmds()[cmdIdx()]);

  // Key handling
  const onKey = (e: KeyboardEvent) => {
    switch (e.key) {
      case 'Enter': {
        const s = selected();
        if (s) {
          const cmd = buildPreviewCommand(s);
          copyToClipboard(cmd);
          showToast('Command copied to clipboard');
        }
        break;
      }
      case 'ArrowUp': {
        if (focus()==='cat' && catIdx()>0) setCatIdx(catIdx()-1);
        else if (focus()==='cmd' && cmdIdx()>0) setCmdIdx(cmdIdx()-1);
        else if (focus()==='detail') window.scrollBy({ top: -24 });
        break;
      }
      case 'ArrowDown': {
        if (focus()==='cat' && catIdx()<groups().length-1) setCatIdx(catIdx()+1);
        else if (focus()==='cmd' && cmdIdx()<cmds().length-1) setCmdIdx(cmdIdx()+1);
        else if (focus()==='detail') window.scrollBy({ top: 24 });
        break;
      }
      case 'ArrowLeft': {
        if (focus()==='cmd') setFocus('cat');
        else if (focus()==='detail') setFocus('cmd');
        break;
      }
      case 'ArrowRight': {
        if (focus()==='cat') setFocus('cmd');
        else if (focus()==='cmd') setFocus('detail');
        break;
      }
    }
  };
  onMount(() => window.addEventListener('keydown', onKey));
  onCleanup(() => window.removeEventListener('keydown', onKey));

  // When category changes, reset command index
  createEffect(() => { cmds(); setCmdIdx(0); });

  // Build command preview similar to PowerShell rendering logic
  function buildPreviewCommand(cmd: CmdItem) {
    // For UI parity: just render a base invocation with placeholder parameters
    // Users can edit parameters in detail panel (future extension)
    return cmd.Invocation || cmd.Name;
  }

  function copyToClipboard(text: string) {
    navigator.clipboard?.writeText(text).catch(() => {
      const ta = document.createElement('textarea');
      ta.value = text; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta);
    });
  }

  function showToast(msg: string) {
    setToast(msg);
    setTimeout(() => setToast(''), 1500);
  }

  // Render helpers
  const formatExamples = (exs: string[] = []) => {
    const out: JSX.Element[] = [];
    exs.slice(0,3).forEach((raw, i) => {
      if (i>0) out.push(<div class="code">{"\n"}</div>); // single blank line between examples
      const lines = raw.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
      if (lines.length>0) {
        out.push(<div class="code"><span class="bullet">{SYM.bullet}</span>{lines[0]}</div>);
        out.push(<div class="code">{"\n"}</div>); // blank line between command and output
      }
      for (let j=1; j<lines.length; j++) {
        const t = lines[j].replace(/\s+$/,'');
        out.push(<div class="code">{'  ' + t}</div>);
      }
    });
    return out;
  };

  return (
    <div class="app">
      <div class="header">VTS Tools Browser</div>
      <div class="main">
        <div class={"pane "+(focus()==='cat'?'active':'')}>
          <div class="pane-title">Categories</div>
          <div class="list">
            <For each={groups()}>{(g, i) => (
              <div class={'row '+(i()==catIdx()?'selected':'')}>
                <span class="bullet">{SYM.bullet}</span>{g.Category} ({g.Commands?.length ?? 0})
              </div>
            )}</For>
          </div>
        </div>
        <div class={"pane "+(focus()==='cmd'?'active':'')}>
          <div class="pane-title">Commands</div>
          <div class="list">
            <For each={cmds()}>{(c, i) => (
              <div class={'row '+(i()==cmdIdx()?'selected':'')}>
                <span class="bullet">{SYM.bullet}</span>{c.Name}
              </div>
            )}</For>
          </div>
        </div>
        <div class={"pane "+(focus()==='detail'?'active':'')}>
          <div class="pane-title">Details</div>
          <div class="list">
            <Show when={selected()}>
              {(s) => (
                <div>
                  <div class="code"><span class="section">Name: </span>{s().Name}</div>
                  <div class="code"><span class="section">Category: </span>{s().Category}</div>
                  <Show when={s().Synopsis}><div class="section">Synopsis:</div></Show>
                  <Show when={s().Synopsis}><div class="code">{s().Synopsis}</div></Show>
                  <Show when={s().Description}><div class="section">Description:</div></Show>
                  <Show when={s().Description}><div class="code">{s().Description}</div></Show>
                  <Show when={s().Parameters?.length}> 
                    <div class="section">Parameters:</div>
                    <For each={s().Parameters}>{(p) => (
                      <div class="code">{' '}{SYM.bullet} {p.Name} &lt;{p.Type}&gt; ({p.Required?'required':'optional'})</div>
                    )}</For>
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
      <div class={"copy-toast "+(toast()? 'show':'')}>{toast()}</div>
    </div>
  );
}

render(() => <App />, document.getElementById('root')!);
