# Christmas Giving Tree

Single-page embed. Edit `CONFIG` at the top of the script in `index.html`.

## Sheet
Columns: `id, need, details, for, amount, status` (`open` / `claimed` / `filled`; `removed` or `hidden` takes a tag down).
Publish: File → Share → Publish to web → the "Needs" tab → CSV → paste the link into `sheetCsvUrl`.

## Embed
```html
<iframe id="giving-tree" src="https://YOUR-USER.github.io/giving-tree/" allow="clipboard-write" title="Christmas Giving Tree"
  style="width:100%;border:0;min-height:900px"></iframe>
<script>
(function(){var f=document.getElementById('giving-tree'),t=0;
function send(){f.contentWindow.postMessage({type:'giving-tree:viewport',height:window.innerHeight,launch:true},'*');view();}
function view(){f.contentWindow.postMessage({type:'giving-tree:view',top:-f.getBoundingClientRect().top,height:window.innerHeight},'*');}
addEventListener('scroll',function(){if(!t){t=1;requestAnimationFrame(function(){t=0;view();});}},{passive:true});
addEventListener('message',function(e){var d=e.data;if(!d||e.source!==f.contentWindow)return;
 if(d.type==='giving-tree:height'){f.style.height=d.height+'px';send();}
 if(d.type==='giving-tree:scroll'){window.scrollTo({top:f.getBoundingClientRect().top+window.scrollY+d.y-16,behavior:'smooth'});}
 if(d.type==='giving-tree:modal'){document.documentElement.style.overflow=d.open?'hidden':'';if(d.open)view();}
 if(d.type==='giving-tree:launch'&&/^https:\/\/launcher\.nucleus\.church\//.test(d.href)){var a=document.createElement('a');a.href=d.href;a.style.display='none';document.body.appendChild(a);a.click();a.remove();}});
addEventListener('resize',send);f.addEventListener('load',send);})();
</script>
```
