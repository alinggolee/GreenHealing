/**
 * Lesson page — 目標 on top (with course title), 影片/活動/TEAMI in a row
 */

import { Router } from '../router.js';
import { parseMarkdown } from '../markdown.js';
import { getLang } from '../lang.js';

/**
 * Parse summary.md into sections by ## headings.
 */
function parseSummary(md) {
  const result = {};
  const parts = md.split(/^## /m);
  for (const part of parts) {
    if (!part.trim()) continue;
    const firstLine = part.split('\n')[0].trim();
    const body = part.slice(firstLine.length).trim();
    if (/目標|Objective/i.test(firstLine)) result.objective = body;
    else if (/課程內容|Course Content/i.test(firstLine)) result.content = body;
    else if (/ESP/i.test(firstLine)) result.esp = body;
  }
  return result;
}

export async function renderLesson(container, lesson) {
  const wrapper = document.createElement('div');

  const hasContent = lesson.sections.includes('content');
  const hasEsp = lesson.sections.includes('esp');
  const hasTeami = lesson.sections.includes('teami');

  // Build bottom row items
  let bottomItems = '';
  if (hasContent) {
    bottomItems += `
          <div class="section-block" data-section="content">
            <div class="section-block-header">
              <span class="section-block-title">Course Content</span>
              <span class="section-block-link" data-nav="#/lesson/${lesson.id}/content">Full →</span>
            </div>
            <div class="section-block-body" id="summary-content">
              <div class="loading"></div>
            </div>
          </div>`;
  }
  if (hasEsp) {
    bottomItems += `
          <div class="section-block" data-section="esp">
            <div class="section-block-header">
              <span class="section-block-title">ESP</span>
              <span class="section-block-link" data-nav="#/lesson/${lesson.id}/esp">Full →</span>
            </div>
            <div class="section-block-body" id="summary-esp">
              <div class="loading"></div>
            </div>
          </div>`;
  }
  // Always render TEEMI
  bottomItems += `
          <div class="teami-block">
            <div>
              <a href="https://teemi.tw/" target="_blank" rel="noopener" class="teami-btn">TEEMI</a>
            </div>
          </div>`;

  wrapper.innerHTML = `
    <div class="lesson-sections">
      <div class="section-block objective-block" data-section="objective">
        ${(lesson.objectiveImage || lesson.cover) ? `<img src="${lesson.objectiveImage || lesson.cover}" class="objective-bg-img" alt="">` : ''}
        <div class="objective-content-wrapper">
          <div class="section-block-header">
            <span class="section-block-title is-main">${lesson.id} — ${lesson.title}</span>
          </div>
          <div class="section-block-body" id="summary-objective">
            <div class="loading"></div>
          </div>
        </div>
      </div>
      <div class="lesson-bottom-row">
        ${bottomItems}
      </div>
    </div>
  `;

  // Navigation
  wrapper.addEventListener('click', (e) => {
    const link = e.target.closest('.section-block-link[data-nav]');
    if (!link) return;
    e.preventDefault();
    Router.navigate(link.dataset.nav);
  });

  container.replaceChildren(wrapper);

  // Fetch summary
  try {
    const suffix = getLang() === 'en' ? '-en' : '';
    let res = await fetch(`data/${lesson.id}/summary${suffix}.md`);
    if (!res.ok && suffix !== '') {
      res = await fetch(`data/${lesson.id}/summary.md`);
    }
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    let md = await res.text();
    md = md.replace(
      /!\[([^\]]*)\]\((?!https?:\/\/)([^)]+)\)/g,
      (_, alt, src) => `![${alt}](data/${lesson.id}/${src})`
    );

    const parsed = parseSummary(md);
    for (const [key, content] of Object.entries(parsed)) {
      const el = wrapper.querySelector(`#summary-${key}`);
      if (el) {
        el.innerHTML = `<div class="md-content">${parseMarkdown(content)}</div>`;
      }
    }
  } catch {
    for (const key of ['objective', 'content', 'esp']) {
      const el = wrapper.querySelector(`#summary-${key}`);
      if (el) el.innerHTML = '<p style="color:var(--color-text-muted)">No summary provided.</p>';
    }
  }
}
