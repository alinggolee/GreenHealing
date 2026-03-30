/**
 * Detail page — loads and renders markdown content for a section
 */

import { parseMarkdown } from '../markdown.js';
import { getLang } from '../lang.js';

const sectionNames = {
  content: 'Course Content',
  esp: 'ESP'
};

/**
 * Render a detail page for a specific lesson section.
 * @param {HTMLElement} container
 * @param {object} lesson - lesson object from config
 * @param {string} section - 'objective' | 'video' | 'activity' | 'teami'
 */
export async function renderDetail(container, lesson, section) {
  const sectionLabel = sectionNames[section] || section;

  container.innerHTML = `
    <div class="detail-content">
      <div class="md-content loading"></div>
    </div>
  `;

  const mdContainer = container.querySelector('.md-content');

  try {
    const suffix = getLang() === 'en' ? '-en' : '';
    let res = await fetch(`data/${lesson.id}/${section}${suffix}.md`);

    // Fallback to default if -en file doesn't exist
    if (!res.ok && suffix !== '') {
      res = await fetch(`data/${lesson.id}/${section}.md`);
    }

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const mdText = await res.text();

    // Resolve relative image paths
    const resolvedMd = mdText.replace(
      /!\[([^\]]*)\]\((?!https?:\/\/)([^)]+)\)/g,
      (_, alt, src) => `![${alt}](data/${lesson.id}/${src})`
    );

    mdContainer.classList.remove('loading');
    mdContainer.innerHTML = parseMarkdown(resolvedMd);
  } catch (err) {
    mdContainer.classList.remove('loading');
    mdContainer.innerHTML = `<p>無法載入內容。（${err.message}）</p>`;
  }
}
