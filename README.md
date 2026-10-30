# Christmas Giving Tree

Single-page embed. Edit `CONFIG` at the top of the script in `index.html`.

## Sheet
Columns: `id, need, details, for, amount, status` (`open` / `claimed` / `filled`; `removed` or `hidden` takes a tag down).
Publish: File → Share → Publish to web → the "Needs" tab → CSV → paste the link into `sheetCsvUrl`.

## Embed
```html
<iframe id="giving-tree" src="https://givingtree.oceansidechurch.ca/" allow="clipboard-write" title="Christmas Giving Tree"
  style="width:100%;border:0;min-height:900px"></iframe>
<script>
(function(){var f=document.getElementById('giving-tree'),t=0,ly=0;
function send(){f.contentWindow.postMessage({type:'giving-tree:viewport',height:window.innerHeight,launch:true,steps:true},'*');view();}
function view(){f.contentWindow.postMessage({type:'giving-tree:view',top:-f.getBoundingClientRect().top,height:window.innerHeight},'*');}
addEventListener('scroll',function(){if(!t){t=1;requestAnimationFrame(function(){t=0;view();});}},{passive:true});
addEventListener('message',function(e){var d=e.data;if(!d||e.source!==f.contentWindow)return;
 if(d.type==='giving-tree:height'){f.style.height=d.height+'px';send();}
 if(d.type==='giving-tree:scroll'){window.scrollTo({top:f.getBoundingClientRect().top+window.scrollY+d.y-16,behavior:d.step?'instant':'smooth'});}
 if(d.type==='giving-tree:modal'){var de=document.documentElement;if(d.open){var r=f.getBoundingClientRect();if(r.bottom<innerHeight)window.scrollBy({top:r.bottom-innerHeight,behavior:'instant'});ly=window.scrollY;}de.style.overflow=d.open?'hidden':'';if(window.scrollY!==ly)window.scrollTo({top:ly,behavior:'instant'});if(d.open)view();}
 if(d.type==='giving-tree:launch'&&/^https:\/\/launcher\.nucleus\.church\//.test(d.href)){var a=document.createElement('a');a.href=d.href;a.style.display='none';document.body.appendChild(a);a.click();a.remove();}});
addEventListener('resize',send);f.addEventListener('load',send);})();
</script>
```

## Forms that write to the sheet (optional)
`apps-script/Code.gs` is a Google Apps Script that turns the Needs spreadsheet into a small backend:
- **Submit a need** adds a row to the Needs tab with status `new` (hidden from the tree) plus the submitter's name and contact in private columns.
- **I'll bring it** records the pledge on a Pledges tab and flips the tag to `claimed` immediately.
- The tree reads only public fields (id, need, details, amount, status) through the script, so names and contact details never reach the page.

Setup steps are at the top of `Code.gs`. Then put the web app URL in `CONFIG.apiUrl`.
Once `apiUrl` is live, stop publishing the sheet to the web (File → Share → Publish to web → Stop publishing): the Needs tab now holds contact details.

## Sunday TV display
`https://givingtree.oceansidechurch.ca/tv.html` (same as `/?tv`). A portrait, full-screen tree with a QR code
to `CONFIG.tvLink`, live counts, and a check of the sheet every `CONFIG.tvRefreshSeconds` (newly claimed tags light up).
Add `?rotate` (`tv.html?rotate`) when the screen is mounted on its side but the computer's display isn't rotated.

`pick.html` (same as `/?pick`) opens straight on the tree with no intro. The TV's QR code points there.
