document.querySelectorAll('.prompt').forEach(function(box){
  var btn = box.querySelector('button.copy');
  var pre = box.querySelector('pre');
  if(!btn || !pre) return;
  btn.addEventListener('click', function(){
    var text = pre.innerText;
    function ok(){ btn.textContent='복사됨 ✓'; btn.classList.add('done'); setTimeout(function(){ btn.textContent='복사'; btn.classList.remove('done'); }, 1400); }
    if(navigator.clipboard && window.isSecureContext){
      navigator.clipboard.writeText(text).then(ok, fallback);
    } else { fallback(); }
    function fallback(){
      var ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
      document.body.appendChild(ta); ta.select();
      try{ document.execCommand('copy'); ok(); }catch(e){ btn.textContent='직접 복사'; }
      document.body.removeChild(ta);
    }
  });
});
function toggleAll(open){ document.querySelectorAll('main details').forEach(function(d){ d.open = open; }); }
