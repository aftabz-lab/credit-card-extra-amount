// Leader-table filters only. Source parsing, mapping and snapshot rules stay in core.js / sync.js.
export const AMOUNT_OPERATIONS=[
  ['eq','Equals...'],['ne','Does Not Equal...'],['gt','Greater Than...'],
  ['gte','Greater Than Or Equal To...'],['lt','Less Than...'],
  ['lte','Less Than Or Equal To...'],['between','Between...'],['top','Top 10...'],
  ['above','Above Average'],['below','Below Average'],['custom','Custom Filter...']
];
const comparisons=AMOUNT_OPERATIONS.slice(0,6);
const escapeHtml=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const format=value=>new Intl.NumberFormat('en-US',{maximumFractionDigits:10}).format(value);

export function parseAmount(value){
  const text=String(value??'').trim().replace(/,/g,'');
  if(!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(text))return null;
  const n=Number(text);return Number.isFinite(n)?n:null;
}
function compare(value,operator,limit){
  if(!Number.isFinite(value)||!Number.isFinite(limit))return false;
  // Ignore floating-point addition noise without rounding the source amounts.
  const equal=Math.abs(value-limit)<=Number.EPSILON*Math.max(1,Math.abs(value),Math.abs(limit))*8;
  switch(operator){
    case 'eq':return equal;
    case 'ne':return !equal;
    case 'gt':return value>limit&&!equal;
    case 'gte':return value>limit||equal;
    case 'lt':return value<limit&&!equal;
    case 'lte':return value<limit||equal;
    default:return false;
  }
}
export function applyAmountFilter(leaders,rule){
  if(!rule)return leaders;
  const values=leaders.map(g=>g.extra).filter(Number.isFinite);
  if(!values.length)return [];
  if(rule.operator==='above'||rule.operator==='below'){
    const average=values.reduce((sum,value)=>sum+value,0)/values.length;
    return leaders.filter(g=>compare(g.extra,rule.operator==='above'?'gt':'lt',average));
  }
  if(rule.operator==='top'){
    const count=rule.percent?Math.ceil(values.length*rule.count/100):rule.count;
    const sorted=[...values].sort((a,b)=>rule.bottom?a-b:b-a);
    const limit=sorted[Math.max(0,Math.min(sorted.length,Math.ceil(count))-1)];
    // As in a spreadsheet's Top 10 filter, include ties at the boundary.
    return leaders.filter(g=>compare(g.extra,rule.bottom?'lte':'gte',limit));
  }
  return leaders.filter(g=>{
    if(rule.operator==='between')return compare(g.extra,'gte',rule.value)&&compare(g.extra,'lte',rule.second);
    if(rule.operator==='custom'){
      const first=compare(g.extra,rule.firstOperator,rule.value);
      if(!rule.secondOperator)return first;
      const second=compare(g.extra,rule.secondOperator,rule.second);
      return rule.join==='or'?first||second:first&&second;
    }
    return compare(g.extra,rule.operator,rule.value);
  });
}
function describe(rule){
  if(!rule)return '';
  const label=operator=>comparisons.find(([key])=>key===operator)?.[1].replace(/\.\.\.$/,'')||operator;
  if(rule.operator==='above'||rule.operator==='below')return rule.operator==='above'?'Above Average':'Below Average';
  if(rule.operator==='top')return `${rule.bottom?'Bottom':'Top'} ${rule.count}${rule.percent?'%':''}`;
  if(rule.operator==='between')return `${format(rule.value)}–${format(rule.second)} BDT`;
  if(rule.operator==='custom')return `${label(rule.firstOperator)} ${format(rule.value)}${rule.secondOperator?` ${rule.join.toUpperCase()} ${label(rule.secondOperator)} ${format(rule.second)}`:''} BDT`;
  return `${label(rule.operator)} ${format(rule.value)} BDT`;
}

export function createLeaderFilters({getBanks,getSelectedBanks,setSelectedBanks,getAmount,setAmount}){
  const $=id=>document.getElementById(id);
  const bankButton=$('leader-bank-button'),amountButton=$('leader-amount-button');
  const bankPanel=$('leader-bank-panel'),amountPanel=$('leader-amount-panel');
  const panel=$('leader-summary-panel');
  bankPanel.innerHTML=`<div class="leader-filter-title">Bank/MFS</div><p class="leader-filter-note">All channels are included unless selected below.</p>
    <input id="leader-bank-search" type="search" placeholder="Search Bank/MFS" aria-label="Search Bank/MFS options">
    <div class="leader-bank-tools"><button type="button" class="text-button" id="leader-bank-select-all">Select all shown</button><button type="button" class="text-button" id="leader-bank-clear">Clear selection</button><span id="leader-bank-count"></span></div>
    <div id="leader-bank-options" class="leader-bank-options" role="listbox" aria-label="Bank/MFS options" aria-multiselectable="true"></div>`;
  const comparisonOptions=comparisons.map(([key,label])=>`<option value="${key}">${label.replace(/\.\.\.$/,'')}</option>`).join('');
  amountPanel.innerHTML=`<div class="leader-filter-title">Amount</div><p class="leader-filter-note">Filter leader Total extra amount in BDT.</p>
    <div class="leader-amount-layout"><div class="leader-amount-menu" aria-label="Amount filter options">${AMOUNT_OPERATIONS.map(([key,label])=>`<button type="button" data-amount-operation="${key}" aria-pressed="false">${label}</button>`).join('')}</div>
    <form id="leader-amount-form" class="leader-amount-form" hidden><strong id="leader-amount-operation-label"></strong>
      <div id="leader-amount-simple"><label>Amount (BDT)<input id="leader-amount-value" type="text" inputmode="decimal" autocomplete="off" placeholder="e.g. 10000"></label><label id="leader-amount-second-wrap" hidden>To amount (BDT)<input id="leader-amount-second" type="text" inputmode="decimal" autocomplete="off"></label></div>
      <div id="leader-amount-top" hidden><label>Show<select id="leader-amount-rank"><option value="top">Top</option><option value="bottom">Bottom</option></select></label><label>Number<input id="leader-amount-top-count" type="number" min="1" step="1" value="10"></label><label>Of leaders<select id="leader-amount-top-unit"><option value="items">Items</option><option value="percent">Percent</option></select></label></div>
      <div id="leader-amount-custom" hidden><label>First condition<select id="leader-amount-first-op">${comparisonOptions}</select></label><label>Amount (BDT)<input id="leader-amount-first-value" type="text" inputmode="decimal" autocomplete="off"></label><label>Combine<select id="leader-amount-join"><option value="and">AND</option><option value="or">OR</option></select></label><label>Second condition<select id="leader-amount-second-op"><option value="">None</option>${comparisonOptions}</select></label><label>Amount (BDT)<input id="leader-amount-second-value" type="text" inputmode="decimal" autocomplete="off" disabled></label></div>
      <p id="leader-amount-average" class="leader-filter-note" hidden>The average is calculated from the leaders in the current bank and dashboard filter selection.</p>
      <p id="leader-amount-error" class="leader-filter-error" role="alert" hidden></p><button class="button small" type="submit">Apply</button>
    </form></div><div class="leader-amount-footer"><button class="text-button" type="button" id="leader-amount-clear">Clear amount filter</button><button class="text-button" type="button" id="leader-amount-close">Close</button></div>`;
  let draftOperator=null;
  function close(returnFocus=false){
    const button=bankPanel.hidden?amountButton:bankButton;
    bankPanel.hidden=true;amountPanel.hidden=true;
    bankButton.setAttribute('aria-expanded','false');amountButton.setAttribute('aria-expanded','false');
    if(returnFocus)button.focus({preventScroll:true});
  }
  bankButton.onclick=()=>{const opening=bankPanel.hidden;close();bankPanel.hidden=!opening;bankButton.setAttribute('aria-expanded',String(opening));if(opening)$('leader-bank-search').focus({preventScroll:true});};
  amountButton.onclick=()=>{const opening=amountPanel.hidden;close();amountPanel.hidden=!opening;amountButton.setAttribute('aria-expanded',String(opening));};
  document.addEventListener('click',event=>{if(!bankPanel.contains(event.target)&&!amountPanel.contains(event.target)&&!bankButton.contains(event.target)&&!amountButton.contains(event.target))close();},true);
  panel.addEventListener('keydown',event=>{if(event.key==='Escape'&&(!bankPanel.hidden||!amountPanel.hidden)){event.preventDefault();close(true);}});
  function renderBanks(){
    const banks=getBanks(),selected=getSelectedBanks(),query=$('leader-bank-search').value.toLocaleLowerCase().trim();
    const shown=banks.filter(bank=>String(bank.label).toLocaleLowerCase().includes(query));
    const list=$('leader-bank-options'),scroll=list.scrollTop,focused=document.activeElement?.dataset.leaderBank;
    const html=shown.map(bank=>`<label role="option" aria-selected="${selected.includes(bank.key)}"><input type="checkbox" data-leader-bank="${escapeHtml(bank.key)}" ${selected.includes(bank.key)?'checked':''}><span>${escapeHtml(bank.label)}</span></label>`).join('')||'<div class="leader-filter-note">No matching Bank/MFS.</div>';
    if(list.innerHTML!==html){list.innerHTML=html;list.scrollTop=scroll;if(focused)[...list.querySelectorAll('input')].find(input=>input.dataset.leaderBank===focused)?.focus({preventScroll:true});}
    $('leader-bank-count').textContent=selected.length?`${selected.length} selected / ${banks.length} available`:`All ${banks.length} Bank/MFS`;
  }
  $('leader-bank-search').oninput=renderBanks;
  $('leader-bank-options').onchange=event=>{
    const key=event.target.dataset.leaderBank;if(!key)return;
    const selected=getSelectedBanks();setSelectedBanks(event.target.checked?[...new Set([...selected,key])]:selected.filter(value=>value!==key));
  };
  $('leader-bank-select-all').onclick=()=>{const shown=[...$('leader-bank-options').querySelectorAll('input')].map(input=>input.dataset.leaderBank);setSelectedBanks([...new Set([...getSelectedBanks(),...shown])]);};
  $('leader-bank-clear').onclick=()=>setSelectedBanks([]);
  $('leader-amount-second-op').onchange=()=>{$('leader-amount-second-value').disabled=!$('leader-amount-second-op').value;};
  function edit(operator){
    draftOperator=operator;const rule=getAmount()?.operator===operator?getAmount():{};
    $('leader-amount-form').hidden=false;$('leader-amount-error').hidden=true;
    $('leader-amount-operation-label').textContent=AMOUNT_OPERATIONS.find(([key])=>key===operator)[1].replace(/\.\.\.$/,'');
    const simple=comparisons.some(([key])=>key===operator)||operator==='between';
    $('leader-amount-simple').hidden=!simple;$('leader-amount-second-wrap').hidden=operator!=='between';
    $('leader-amount-top').hidden=operator!=='top';$('leader-amount-custom').hidden=operator!=='custom';$('leader-amount-average').hidden=!['above','below'].includes(operator);
    $('leader-amount-value').value=rule.value??'';$('leader-amount-second').value=rule.second??'';
    $('leader-amount-rank').value=rule.bottom?'bottom':'top';$('leader-amount-top-count').value=rule.count??10;$('leader-amount-top-unit').value=rule.percent?'percent':'items';
    $('leader-amount-first-op').value=rule.firstOperator||'gte';$('leader-amount-first-value').value=rule.value??'';$('leader-amount-join').value=rule.join||'and';$('leader-amount-second-op').value=rule.secondOperator||'';$('leader-amount-second-value').value=rule.second??'';$('leader-amount-second-value').disabled=!rule.secondOperator;
    amountPanel.querySelectorAll('[data-amount-operation]').forEach(button=>button.setAttribute('aria-pressed',String(button.dataset.amountOperation===operator)));
    const focus=simple?'leader-amount-value':operator==='top'?'leader-amount-top-count':operator==='custom'?'leader-amount-first-value':null;
    if(focus)$(focus).focus({preventScroll:true});
  }
  amountPanel.addEventListener('click',event=>{const button=event.target.closest('[data-amount-operation]');if(button)edit(button.dataset.amountOperation);});
  $('leader-amount-form').onsubmit=event=>{
    event.preventDefault();if(!draftOperator)return;
    const rule={operator:draftOperator};let error='';
    if(draftOperator==='top'){
      rule.count=parseAmount($('leader-amount-top-count').value);rule.bottom=$('leader-amount-rank').value==='bottom';rule.percent=$('leader-amount-top-unit').value==='percent';
      if(rule.count===null||!Number.isInteger(rule.count)||rule.count<1||(rule.percent&&rule.count>100))error=rule.percent?'Enter a whole percentage from 1 to 100.':'Enter a positive whole number of leaders.';
    }else if(draftOperator==='custom'){
      rule.firstOperator=$('leader-amount-first-op').value;rule.value=parseAmount($('leader-amount-first-value').value);rule.secondOperator=$('leader-amount-second-op').value;rule.join=$('leader-amount-join').value;
      if(rule.secondOperator)rule.second=parseAmount($('leader-amount-second-value').value);
      if(rule.value===null||(rule.secondOperator&&rule.second===null))error='Enter a valid BDT amount for each condition.';
    }else if(!['above','below'].includes(draftOperator)){
      rule.value=parseAmount($('leader-amount-value').value);
      if(draftOperator==='between')rule.second=parseAmount($('leader-amount-second').value);
      if(rule.value===null||(draftOperator==='between'&&rule.second===null))error='Enter a valid BDT amount.';
      else if(draftOperator==='between'&&rule.value>rule.second)error='The first amount must be less than or equal to the second amount.';
    }
    if(error){$('leader-amount-error').textContent=error;$('leader-amount-error').hidden=false;return;}
    setAmount(rule);close();amountButton.focus({preventScroll:true});
  };
  $('leader-amount-clear').onclick=()=>{setAmount(null);draftOperator=null;$('leader-amount-form').hidden=true;amountPanel.querySelectorAll('[data-amount-operation]').forEach(button=>button.setAttribute('aria-pressed','false'));close();amountButton.focus({preventScroll:true});};
  $('leader-amount-close').onclick=()=>close(true);
  return {sync({shown,total}){
    renderBanks();const selected=getSelectedBanks(),rule=getAmount();
    bankButton.classList.toggle('leader-filter-active',Boolean(selected.length));amountButton.classList.toggle('leader-filter-active',Boolean(rule));
    bankButton.title=selected.length?`Selected Bank/MFS: ${getBanks().filter(bank=>selected.includes(bank.key)).map(bank=>bank.label).join(', ')}`:'All Bank/MFS';
    amountButton.title=rule?describe(rule):'Filter leader amounts';
    const parts=[];if(selected.length)parts.push(`Bank/MFS: ${getBanks().filter(bank=>selected.includes(bank.key)).map(bank=>bank.label).join(', ')}`);if(rule)parts.push(`Amount: ${describe(rule)}`);
    $('leader-filter-summary').hidden=!parts.length;$('leader-filter-summary').textContent=parts.length?`${parts.join(' · ')} · ${shown} of ${total} leaders`:'';
  }};
}
