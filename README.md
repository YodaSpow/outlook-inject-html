# ✨ Outlook Web “Inject HTML” Bookmarklet

**outlook-inject-html** is a tiny, single-click **bookmarklet** (JavaScript that runs from a browser bookmark) that lets you insert a local `.html` file directly into the **Message body** of **Outlook Web**.

- **Who it’s for:** Marketers, CRM/Email teams, and non-technical users who receive ready-to-send HTML from designers/MarTech.
- **What it solves:** Avoids copying/pasting messy HTML or digging through dev tools. Click the bookmark → choose your `filename.html` → content appears in the email body.
- **Why it’s safe/simple:** Uses the browser’s file picker (no install), targets the real compose editor, and briefly highlights the area it’s modifying.

**Live page (with drag-to-bookmarks button):** https://yodaspow.github.io/outlook-inject-html/

---

## 🧩 Summary

- **Chrome-only tested** (macOS/Windows).
- Works on **Outlook Web**: `outlook.office.com`.
- Inserts your local **`filename.html`** into the compose **Message body**.
- If you type **`####`** in the body first, the tool will replace that placeholder **in place** — most users will want to put `####` **before their signature** to keep signatures intact.

> Outlook removes `<script>` and may limit some CSS. Use inline CSS and table-based email markup for best results.

---

## ⚡ Quick Start

### Option A — easiest (drag method)
1. **Show your Bookmarks Bar**  
   - Mac: `Cmd + Shift + B`  
   - Windows: `Ctrl + Shift + B`
2. Go to the **live page**: https://yodaspow.github.io/outlook-inject-html/  
3. **Drag** the **“Outlook • Inject HTML”** pill onto your Bookmarks Bar.

### Option B — manual (keyboard-friendly)
1. **Show your Bookmarks Bar** (`Cmd + Shift + B` / `Ctrl + Shift + B`)
2. **Open Bookmark Manager**  
   - Mac: `Option + Cmd + B`  
   - Windows: `Ctrl + Shift + O`
3. Click **New bookmark**  
   - **Name:** `Outlook • Inject HTML`  
   - **URL:** paste the **bookmarklet** from below (single line)  
   - Save it to your **Bookmarks Bar**

---

## ▶️ Use it
1. Open **[outlook.office.com](https://outlook.office.com/mail/inbox/)** → click **New message**  
2. *(Optional but recommended)* type **`####` _before your signature_**  
3. Click the bookmark → pick your **`filename.html`**  
4. Done 🎉

**Confirmation behavior (no `####` present):**  
If the message body already has content and no `####`, you’ll see a confirm dialog:

> **OK** replaces the entire body (including any signature).  
> **Cancel** keeps your content — add `####` before the signature and run again.

---

## 🔖 The Bookmarklet

> Paste this entire single line into the **URL** field when creating the bookmark.

```js
javascript:(()=>{ /* Outlook Inject HTML from Local File (target fix) */function toast(m){const t=document.createElement('div');t.textContent=m;t.style.cssText='position:fixed;z-index:9999999;left:50%;top:16px;transform:translateX(-50%);background:#111;color:#fff;padding:10px 14px;border-radius:8px;font:13px/1.3 system-ui,Segoe UI,Arial;box-shadow:0 6px 16px rgba(0,0,0,.25)';document.body.appendChild(t);setTimeout(()=>t.remove(),3200)}function highlight(el){if(!el) return;const r=el.getBoundingClientRect();const o=document.createElement('div');o.style.cssText=%60position:fixed;left:${r.left+window.scrollX}px;top:${r.top+window.scrollY}px;width:${r.width}px;height:${r.height}px;outline:3px solid #4caf50;outline-offset:0;pointer-events:none;z-index:9999998;border-radius:6px;%60;document.body.appendChild(o);setTimeout(()=>o.remove(),1200)}function findBody(doc=document){/* 1) Prefer aria-label="Message body" in page */let el=doc.querySelector('[contenteditable="true"][aria-label="Message body"],div[role="textbox"][aria-label="Message body"]');if(el) return {el,ctx:doc,how:'direct'};/* 2) Look inside iframes */for(const f of Array.from(doc.querySelectorAll('iframe'))){try{const idoc=f.contentDocument; if(!idoc) continue; el=idoc.querySelector('[contenteditable="true"][aria-label="Message body"],div[role="textbox"][aria-label="Message body"]'); if(el) return {el,ctx:idoc,how:'iframe'};}catch(e){}}/* 3) Fallback: pick contenteditable containing #### */for(const c of Array.from(doc.querySelectorAll('[contenteditable="true"],div[role="textbox"]'))){if((c.textContent||'').includes('####')) return {el:c,ctx:doc,how:'placeholder'}}/* 4) Last resort: largest contenteditable block (avoid To/Cc by size) */let best=null,area=0;for(const c of Array.from(doc.querySelectorAll('[contenteditable="true"],div[role="textbox"]'))){const b=c.getBoundingClientRect();const a=b.width*b.height;if(a>area){area=a;best=c}}return {el:best,ctx:doc,how:'largest'}}function injectInto(targetEl,html){const cur=targetEl.innerHTML||'';if(cur.includes('####')){targetEl.innerHTML=cur.replace('####',html)}else{const has=((targetEl.innerText||'').trim().length>0);if(has){if(!confirm('No "####" found.\n\nPress OK to replace the entire message body with your HTML (including any signature).\nPress Cancel to keep your content — add #### before your signature and run again.')){toast('Cancelled.');return}}targetEl.innerHTML=html}targetEl.focus&&targetEl.focus();toast('HTML injected into Message body ✔')}function pickFile(cb){const i=document.createElement('input');i.type='file';i.accept='.html,text/html,.htm';i.style.display='none';i.onchange=()=>{const f=i.files&&i.files[0];if(!f){toast('No file selected.');i.remove();return}const r=new FileReader();r.onload=()=>{cb(String(r.result||''));setTimeout(()=>i.remove(),0)};r.onerror=()=>{alert('Could not read file.');};r.readAsText(f)};document.body.appendChild(i);i.click()}if(!/outlook\.office\.com|outlook\.live\.com/.test(location.hostname)){if(!confirm('This is intended for Outlook Web. Continue anyway?')) return}const tgt=findBody(); if(!tgt.el){alert('Compose “Message body” not found. Click “New message” first, then run again.');return}highlight(tgt.el); pickFile(html=>injectInto(tgt.el,html));})();
