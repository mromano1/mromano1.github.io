
const slides = window.SLIDES || [];
let current = 0;
const img = document.getElementById('slideImage');
const title = document.getElementById('slideTitle');
const progress = document.getElementById('progress');
const slideText = document.getElementById('slideText');
const notesText = document.getElementById('notesText');
function list(items){return items && items.length ? '<ul>'+items.map(x=>`<li>${x}</li>`).join('')+'</ul>' : '<p>No text extracted for this slide.</p>'}
function show(i){current=Math.max(0,Math.min(slides.length-1,i)); const s=slides[current]; img.src=s.image; img.alt='Slide '+s.n+': '+s.title; title.textContent=s.title; progress.textContent=`Slide ${s.n} of ${slides.length}`; slideText.innerHTML=list(s.text); notesText.innerHTML=list(s.notes); history.replaceState(null,'',`#slide-${s.n}`);}
document.getElementById('prev').addEventListener('click',()=>show(current-1));
document.getElementById('next').addEventListener('click',()=>show(current+1));
document.addEventListener('keydown',e=>{if(e.key==='ArrowLeft')show(current-1); if(e.key==='ArrowRight')show(current+1)});
const m=location.hash.match(/slide-(\d+)/); show(m?parseInt(m[1])-1:0);
