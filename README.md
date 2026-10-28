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
addEventListener('message',e=>{
  if(e.data&&e.data.type==='giving-tree:height')
    document.getElementById('giving-tree').style.height=e.data.height+'px';
});
</script>
```
