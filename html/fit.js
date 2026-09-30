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
