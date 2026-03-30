/**
 * Top navigation bar component
 */

import { Router } from '../router.js';
import { toggleTheme, isDarkTheme } from '../theme.js';
import { getLang, toggleLang } from '../lang.js';

const sections = [
    { key: 'lesson', label: 'Course Page', icon: '📋' },
    { key: 'content', label: 'Course Content', icon: '📖' },
    { key: 'esp', label: 'ESP', icon: '💡' },
    { key: 'teemi', label: 'TEEMI', icon: '👥', external: true, url: 'https://teemi.tw/' }
];

/**
 * Render top navigation bar.
 * @param {string} lessonId - e.g. 'W3'
 * @param {string} activeSection - 'lesson' | 'objective' | 'video' | 'activity' | 'teami'
 */
export function renderNavbar(lessonId, activeSection, lessonData = null) {
    const nav = document.getElementById('top-nav');
    if (!nav) return;

    // Render Theme Toggle into Global Header (once)
    // We do this first so it appears even on directory page (though app.js handles directory separately)
    // But since renderNavbar is called, we ensure it's there.
    const globalHeader = document.getElementById('global-header');
    if (globalHeader && !globalHeader.querySelector('#nav-theme-toggle')) {
        // Theme Toggle
        const themeIcon = isDarkTheme() ? '☀️' : '🌙';
        const modalBtn = document.createElement('button');
        modalBtn.id = 'nav-theme-toggle';
        modalBtn.className = 'nav-btn theme-toggle-btn';
        modalBtn.setAttribute('aria-label', 'Toggle theme');
        modalBtn.innerHTML = themeIcon;
        modalBtn.style.marginLeft = 'auto'; // Right align

        globalHeader.appendChild(modalBtn);

        modalBtn.addEventListener('click', () => {
            const isDark = toggleTheme();
            modalBtn.textContent = isDark ? '☀️' : '🌙';
        });

        // Language Toggle
        const langBtn = document.createElement('button');
        langBtn.id = 'nav-lang-toggle';
        langBtn.className = 'nav-btn lang-toggle-btn';
        langBtn.setAttribute('aria-label', 'Toggle language');
        langBtn.textContent = getLang() === 'en' ? 'Switch to Bilingual (EN/ZH)' : 'Switch to EN';
        langBtn.style.marginLeft = '8px';

        globalHeader.appendChild(langBtn);

        langBtn.addEventListener('click', () => {
            const next = toggleLang();
            langBtn.textContent = next === 'en' ? 'Switch to Bilingual (EN/ZH)' : 'Switch to EN';
            location.reload(); // Reload page to apply language
        });
    }

    // Check if we are on directory page. If so, hide the local nav and return.
    if (activeSection === 'directory') {
        nav.classList.add('hidden');
        nav.innerHTML = '';
        return;
    }

    nav.classList.remove('hidden');

    const navButtons = sections.map(s => {
        let label = s.label;
        // If we have lesson data, use the lesson title for the "lesson" button
        if (s.key === 'lesson' && lessonData && lessonData.title) {
            label = lessonData.title;
        }

        if (s.external) {
            return `<a href="${s.url}" target="_blank" rel="noopener" class="nav-btn">
                ${label}
            </a>`;
        }
        const hash = s.key === 'lesson'
            ? `#/lesson/${lessonId}`
            : `#/lesson/${lessonId}/${s.key}`;
        const isActive = s.key === activeSection;
        return `<button class="nav-btn ${isActive ? 'active' : ''}" data-nav="${hash}">
      ${label}
    </button>`;
    }).join('');

    nav.innerHTML = navButtons;

    nav.addEventListener('click', (e) => {
        const btn = e.target.closest('[data-nav]');
        if (!btn) return;
        Router.navigate(btn.dataset.nav);
    });
}

/** Hide the navbar (for directory page) */
export function hideNavbar() {
    const nav = document.getElementById('top-nav');
    if (nav) {
        nav.classList.add('hidden');
        nav.innerHTML = '';
    }
}
