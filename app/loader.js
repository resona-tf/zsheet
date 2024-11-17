loadJs('common.js');
loadJs('zohoapi.js');
loadJs('varEngine.js');
loadJs('generator.js');
loadJs('main.js');

function loadJs(src) {
	let s = document.createElement("script");
	s.type = "text/javascript";
	s.defer = true;
	s.src = `${src}?r=${Date.now()}`;
	document.body.appendChild(s);
}