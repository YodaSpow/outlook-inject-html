# ✨ Outlook Web “Inject HTML” Bookmarklet

**outlook-inject-html** is a tiny, single-click **bookmarklet** (JavaScript that runs from a browser bookmark) that lets you insert a local `.html` file directly into the **Message body** of **Outlook Web**.

- **Who it’s for:** Marketers, CRM/Email teams, and non-technical users who receive ready-to-send HTML from designers.
- **What it solves:** Avoids copying/pasting messy HTML or digging through dev tools. Click the bookmark → choose your `filename.html` → content appears in the email body.
- **Why it’s safe/simple:** Uses the browser’s file picker (no install), targets the real compose editor, and shows a 1-second highlight of the area it’s modifying.

---

## 🧩 Summary

- Runs in **Chrome** (macOS/Windows).
- Works on **outlook.office.com** (Outlook Web).
- Inserts your local **`filename.html`** into the compose **Message body**.
- If you type `####` in the body first, the tool will replace that placeholder **in place** — most users will want to put `####` **before their signature** to keep signatures intact.

> Outlook removes `<script>` and may limit some CSS. Use inline CSS and table-based email markup for best results.

---

## ⚡ Quick Start

1) **Show your bookmarks bar**  
   - Mac: `Cmd + Shift + B`  
   - Windows: `Ctrl + Shift + B`

2) **Create a new bookmark**  
   - Right-click the bookmarks bar → **Add page…**  
   - **Name:** `Outlook • Inject HTML`  
   - **URL:** paste the bookmarklet code below

3) **Use it**  
   - Open **outlook.office.com** → click **New message**  
   - (Optional but recommended) type `####` **before your signature**  
   - Click the bookmark → choose your `filename.html`  
   - Done!

---

## 🔖 The Bookmarklet

> Paste this entire single line into the **URL** field when creating the bookmark.

```text
javascript:(()=>{ /* Outlook Inject HTML from Local File (target fix) */
function toast(m){const t=document.createElement('div');t.textContent=m;t.style.cssText='position:fixed;z-index:9999999;left:50%;top:16px;transform:translateX(-50%);background:#111;color:#fff;padding:10px 14px;border-radius:8px;font:13px/1.3 system-ui,Segoe UI,Arial;box-shadow:0 6px 16px rgba(0,0,0,.25)';document.body.appendChild(t);setTimeout(()=>t.remove(),3200)}
function highlight(el){if(!el) return;const r=el.getBoundingClientRect();const o=document.createElement('div');o.style.cssText=`position:fixed;left:${r.left+window.scrollX}px;top:${r.top+window.scrollY}px;width:${r.width}px;height:${r.height}px;outline:3px solid #4caf50;outline-offset:0;pointer-events:none;z-index:9999998;border-radius:6px;`;document.body.appendChild(o);setTimeout(()=>o.remove(),1200)}
function findBody(doc=document){/* 1) Prefer aria-label="Message body" in page */
let el=doc.querySelector('[contenteditable="true"][aria-label="Message body"],div[role="textbox"][aria-label="Message body"]');
if(el) return {el,ctx:doc,how:'direct'};
/* 2) Look inside iframes */
for(const f of Array.from(doc.querySelectorAll('iframe'))){try{const idoc=f.contentDocument; if(!idoc) continue; el=idoc.querySelector('[contenteditable="true"][aria-label="Message body"],div[role="textbox"][aria-label="Message body"]'); if(el) return {el,ctx:idoc,how:'iframe'};}catch(e){}}
/* 3) Fallback: pick contenteditable containing #### */
for(const c of Array.from(doc.querySelectorAll('[contenteditable="true"],div[role="textbox"]'))){if((c.textContent||'').includes('####')) return {el:c,ctx:doc,how:'placeholder'}}
/* 4) Last resort: largest contenteditable block (avoid To/Cc by size) */
let best=null,area=0;for(const c of Array.from(doc.querySelectorAll('[contenteditable="true"],div[role="textbox"]'))){const b=c.getBoundingClientRect();const a=b.width*b.height;if(a>area){area=a;best=c}}return {el:best,ctx:doc,how:'largest'}}
function injectInto(targetEl,html){const cur=targetEl.innerHTML||'';if(cur.includes('####')){targetEl.innerHTML=cur.replace('####',html)}else{const has=((targetEl.innerText||'').trim().length>0);if(has){if(!confirm('No "####" in Message body. Replace ENTIRE body with your HTML?')){toast('Cancelled.');return}}targetEl.innerHTML=html}targetEl.focus&&targetEl.focus();toast('HTML injected into Message body ✔')}
function pickFile(cb){const i=document.createElement('input');i.type='file';i.accept='.html,text/html,.htm';i.style.display='none';i.onchange=()=>{const f=i.files&&i.files[0];if(!f){toast('No file selected.');i.remove();return}const r=new FileReader();r.onload=()=>{cb(String(r.result||''));setTimeout(()=>i.remove(),0)};r.onerror=()=>{alert('Could not read file.');};r.readAsText(f)};document.body.appendChild(i);i.click()}
if(!/outlook\.office\.com|outlook\.live\.com/.test(location.hostname)){if(!confirm('This is intended for Outlook Web. Continue anyway?')) return}
const tgt=findBody(); if(!tgt.el){alert('Compose “Message body” not found. Click “New message” first, then run again.');return}
highlight(tgt.el); pickFile(html=>injectInto(tgt.el,html));
})();
