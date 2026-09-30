/* Scale the fixed 1920x1080 stage to the window (letterboxed). */
(function(){
  var deck=document.querySelector('.deck');
  function fit(){
    var s=Math.min(window.innerWidth/1920,window.innerHeight/1080);
    var x=(window.innerWidth-1920*s)/2, y=(window.innerHeight-1080*s)/2;
    deck.style.transform='translate('+x+'px,'+y+'px) scale('+s+')';
  }
  window.addEventListener('resize',fit); fit();
})();

/* Show the theme name briefly when T switches themes (runtime sets data-theme). */
(function(){
  var t=document.createElement('div'); t.className='theme-toast'; document.body.appendChild(t);
  var h; new MutationObserver(function(){
    var n=document.documentElement.getAttribute('data-theme'); if(!n) return;
    t.textContent='Theme · '+n; t.classList.add('show'); clearTimeout(h);
    h=setTimeout(function(){t.classList.remove('show')},1400);
  }).observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});
})();

/* Mark dark themes so deck.css can switch chart colours (html[data-bg="dark"]). */
(function(){
  var root=document.documentElement, probe=document.createElement('i');
  probe.style.cssText='position:absolute;width:0;height:0;color:var(--bg)'; document.body.appendChild(probe);
  function check(){
    var m=getComputedStyle(probe).color.match(/[\d.]+/g); if(!m) return;
    var l=(0.2126*m[0]+0.7152*m[1]+0.0722*m[2])/255;
    root.setAttribute('data-bg', l<0.45?'dark':'light');
  }
  function later(){ var link=document.getElementById('theme-link');
    if(link) link.addEventListener('load',check,{once:true}); setTimeout(check,250); setTimeout(check,900); }
  new MutationObserver(later).observe(root,{attributes:true,attributeFilter:['data-theme']});
  check(); window.addEventListener('load',check);
})();
