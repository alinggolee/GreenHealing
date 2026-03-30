export function getLang() {
    return localStorage.getItem('alinggo-lang') || 'en';
}

export function setLang(lang) {
    localStorage.setItem('alinggo-lang', lang);
}

export function toggleLang() {
    const current = getLang();
    const next = current === 'bi' ? 'en' : 'bi';
    setLang(next);
    return next;
}
