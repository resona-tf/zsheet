// Version 1.0.1 - Synchronous loading implementation
async function loadJs(src) {
    return new Promise((resolve, reject) => {
        const script = document.createElement("script");
        script.type = "text/javascript";
        script.src = `${src}?r=${Date.now()}`;
        
        script.onload = () => resolve();
        script.onerror = () => reject(new Error(`Failed to load ${src}`));
        
        document.body.appendChild(script);
    });
}

async function loadAllScripts() {
    try {
        await loadJs('./js/common.js');
        await loadJs('./js/zohoapi.js');
        await loadJs('./js/varEngine.js');
        await loadJs('./js/generator.js');
        await loadJs('./js/main.js');
        console.log('All scripts loaded successfully');
    } catch (error) {
        console.error('Error loading scripts:', error);
    }
}

loadAllScripts();
