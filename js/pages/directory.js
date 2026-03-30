/**
 * Directory page — grid of lesson items with image + line + title
 */

import { Router } from '../router.js';
import { getLessonColor } from '../theme.js';

export function renderDirectory(container, lessons) {
  const wrapper = document.createElement('div');
  wrapper.innerHTML = `
    <div class="directory-container">
    <div class="card-grid">
      ${lessons.map(lesson => `
        <div class="card" data-nav="#/lesson/${lesson.id}" data-color="${getLessonColor(lesson.id)}">
          <img class="card-image" src="${lesson.cover}" alt="${lesson.title}"
               onerror="this.src='assets/placeholder.svg'">
          <div class="card-body">
            <div class="card-week">${lesson.id}</div>
            <div class="card-title">${lesson.title}</div>
          </div>
        </div>
      `).join('')}
    </div>
    </div>
  `;

  wrapper.addEventListener('click', (e) => {
    const card = e.target.closest('.card[data-nav]');
    if (!card) return;
    Router.navigate(card.dataset.nav);
  });

  container.replaceChildren(wrapper);
}
