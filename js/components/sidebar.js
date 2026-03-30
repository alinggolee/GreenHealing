/**
 * Sidebar component — renders lesson list + back-to-directory button
 */

import { Router } from '../router.js';

const sectionLabels = {
  content: 'Course Content',
  esp: 'ESP',
  teami: 'TEAMI'
};

/**
 * Render sidebar into #sidebar element.
 * @param {Array} lessons - lessons array from config
 * @param {string|null} currentLessonId - e.g. 'W3'
 */
export function renderSidebar(lessons, currentLessonId) {
  const el = document.getElementById('sidebar');
  if (!el) return;

  el.classList.remove('hidden');
  document.getElementById('main').classList.add('with-sidebar');

  el.innerHTML = `
    <div class="sidebar-header-container" style="border-bottom: 1px solid var(--color-border);">
      <div class="sidebar-header">
        <button id="sidebar-toggle-btn" class="sidebar-toggle" title="展開/收合目錄">
          ☰
        </button>
      </div>
    </div>

    <div class="sidebar-lessons">
      ${lessons.map(l => `
        <button class="sidebar-lesson-btn ${l.id === currentLessonId ? 'active' : ''}"
                data-nav="#/lesson/${l.id}">
          <span class="lesson-id">${l.id}</span>
          <span class="lesson-title"> — ${l.title}</span>
        </button>
      `).join('')}
    </div>
  `;

  // Event delegation
  el.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-nav]');
    if (!btn) return;
    Router.navigate(btn.dataset.nav);
    // Close mobile sidebar
    el.classList.remove('open');
    document.getElementById('sidebar-overlay')?.classList.remove('active');
  });
}

/** Hide sidebar (for directory page) */
export function hideSidebar() {
  const el = document.getElementById('sidebar');
  if (el) {
    el.classList.add('hidden');
    el.innerHTML = '';
  }
  document.getElementById('main')?.classList.remove('with-sidebar');
}
