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
(function(){var f=document.getElementById('giving-tree');
function send(){f.contentWindow.postMessage({type:'giving-tree:viewport',height:window.innerHeight},'*');}
addEventListener('message',function(e){var d=e.data;if(!d)return;
 if(d.type==='giving-tree:height'){f.style.height=d.height+'px';send();}
 if(d.type==='giving-tree:scroll'){window.scrollTo({top:f.getBoundingClientRect().top+window.scrollY+d.y-16,behavior:'smooth'});}});
addEventListener('resize',send);f.addEventListener('load',send);})();
</script>
```
