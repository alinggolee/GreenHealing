/**
 * App entry — loads config, initializes router, theme, and components
 */

import { Router } from './router.js';
import { renderDirectory } from './pages/directory.js';
import { renderLesson } from './pages/lesson.js';
import { renderDetail } from './pages/detail.js';
import { renderSidebar, hideSidebar } from './components/sidebar.js';
import { renderNavbar, hideNavbar } from './components/navbar.js';
import { applyLessonColor, initThemeToggle } from './theme.js';

let config = null;

async function loadConfig() {
    const res = await fetch('data/config.json');
    config = await res.json();
}

function getLesson(id) {
    return config.lessons.find(l => l.id === id);
}

function getContent() {
    return document.getElementById('content');
}

async function init() {
    // Theme setup (before content render to avoid flash)
    initThemeToggle();

    await loadConfig();

    const router = new Router();

    const toggleBtn = document.getElementById('sidebar-toggle');
    const sidebar = document.getElementById('sidebar');
    const mainContent = document.getElementById('main');
    const overlay = document.getElementById('sidebar-overlay');

    // Directory page
    router.on('/', () => {
        document.body.classList.add('is-directory');
        applyLessonColor(null);
        // Do NOT render sidebar on directory page
        hideSidebar();
        // Show directory title in navbar (and init global theme toggle)
        renderNavbar(null, 'directory');
        // Render directory grid content
        renderDirectory(getContent(), config.lessons);
        // Hide toggle button on directory page
        if (toggleBtn) toggleBtn.classList.add('hidden');
    });

    // Lesson overview page
    router.on('/lesson/:id', ({ id }) => {
        document.body.classList.remove('is-directory');
        const lesson = getLesson(id);
        if (!lesson) return Router.navigate('#/');
        applyLessonColor(id);
        renderSidebar(config.lessons, id);
        renderNavbar(id, 'lesson', lesson);
        renderLesson(getContent(), lesson);
        if (toggleBtn) toggleBtn.classList.remove('hidden');
    });

    // Detail page
    router.on('/lesson/:id/:section', ({ id, section }) => {
        document.body.classList.remove('is-directory');
        const lesson = getLesson(id);
        if (!lesson) return Router.navigate('#/');
        const validSections = ['content', 'esp'];
        if (!validSections.includes(section)) return Router.navigate(`#/lesson/${id}`);
        applyLessonColor(id);
        renderSidebar(config.lessons, id);
        renderNavbar(id, section, lesson);
        renderDetail(getContent(), lesson, section);
        if (toggleBtn) toggleBtn.classList.remove('hidden');
    });

    router.start();

    // Sidebar toggle logic
    // Sidebar toggle logic - delegated since button is re-rendered
    document.addEventListener('click', (e) => {
        const toggleBtn = e.target.closest('#sidebar-toggle-btn');
        if (!toggleBtn) return;

        const isMobile = window.innerWidth <= 768;
        if (isMobile) {
            sidebar?.classList.toggle('open');
            overlay?.classList.toggle('active');
        } else {
            // Desktop: toggle collapse
            sidebar?.classList.toggle('collapsed');
            // mainContent?.classList.toggle('collapsed-sidebar'); // No margin change needed now
        }
    });

    overlay?.addEventListener('click', () => {
        sidebar?.classList.remove('open');
        overlay?.classList.remove('active');
    });
}

init();
