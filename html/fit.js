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
