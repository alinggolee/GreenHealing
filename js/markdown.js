/**
 * Lightweight Markdown → HTML parser
 * Supports: h1-h3, bold, italic, code, links, images, lists, blockquote,
 *           custom ::video[URL] directive for iframe embeds.
 */

export function parseMarkdown(md) {
    if (!md) return '';

    const lines = md.split('\n');
    let html = '';
    let inList = false;
    let listType = '';
    let inBlockquote = false;
    let inDetails = false;
    let inParagraph = false;

    const closeParagraph = () => {
        if (inParagraph) { html += '</p>'; inParagraph = false; }
    };

    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];

        // Blank line — close open blocks
        if (line.trim() === '') {
            closeParagraph();
            if (inList) {
                html += listType === 'ul' ? '</ul>' : '</ol>';
                inList = false;
            }
            if (inBlockquote) {
                html += '</blockquote>';
                inBlockquote = false;
            }
            continue;
        }

        // Video embed: ::video[URL]
        const videoMatch = line.trim().match(/^::video\[(.+)\]$/);
        if (videoMatch) {
            closeParagraph();
            html += `<div class="video-embed"><iframe src="${escapeHtml(videoMatch[1])}" allowfullscreen loading="lazy"></iframe></div>`;
            continue;
        }

        // ESP Details block start: :::esp[Summary]
        const espMatch = line.trim().match(/^:::esp\[(.+)\]$/);
        if (espMatch) {
            closeParagraph();
            if (inList) { html += listType === 'ul' ? '</ul>' : '</ol>'; inList = false; }
            if (inBlockquote) { html += '</blockquote>'; inBlockquote = false; }

            html += `<details class="esp-details"><summary>${inline(espMatch[1])}</summary><div class="esp-content">`;
            inDetails = true;
            continue;
        }

        // ESP Details block end: :::
        if (inDetails && line.trim() === ':::') {
            closeParagraph();
            if (inList) { html += listType === 'ul' ? '</ul>' : '</ol>'; inList = false; }
            if (inBlockquote) { html += '</blockquote>'; inBlockquote = false; }
            html += `</div></details>`;
            inDetails = false;
            continue;
        }

        // Headings
        const headingMatch = line.trim().match(/^(#{1,3})\s+(.+)$/);
        if (headingMatch) {
            closeParagraph();
            const level = headingMatch[1].length;
            html += `<h${level}>${inline(headingMatch[2])}</h${level}>`;
            continue;
        }

        // Blockquote
        if (line.startsWith('> ')) {
            closeParagraph();
            if (!inBlockquote) {
                html += '<blockquote>';
                inBlockquote = true;
            }
            html += `<p>${inline(line.slice(2))}</p>`;
            continue;
        } else if (inBlockquote) {
            html += '</blockquote>';
            inBlockquote = false;
        }

        // Unordered list
        const ulMatch = line.match(/^[-*]\s+(.+)$/);
        if (ulMatch) {
            closeParagraph();
            if (!inList || listType !== 'ul') {
                if (inList) html += listType === 'ul' ? '</ul>' : '</ol>';
                html += '<ul>';
                inList = true;
                listType = 'ul';
            }
            html += `<li>${inline(ulMatch[1])}</li>`;
            continue;
        }

        // Ordered list
        const olMatch = line.match(/^\d+\.\s+(.+)$/);
        if (olMatch) {
            closeParagraph();
            if (!inList || listType !== 'ol') {
                if (inList) html += listType === 'ul' ? '</ul>' : '</ol>';
                html += '<ol>';
                inList = true;
                listType = 'ol';
            }
            html += `<li>${inline(olMatch[1])}</li>`;
            continue;
        }

        // Close list if non-list line
        if (inList) {
            html += listType === 'ul' ? '</ul>' : '</ol>';
            inList = false;
        }

        // Paragraph
        if (!inParagraph) {
            html += `<p>${inline(line)}`;
            inParagraph = true;
        } else {
            // Continuation of paragraph (with a br tag)
            html += `<br>${inline(line)}`;
        }
    }

    // Close trailing blocks
    closeParagraph();
    if (inList) html += listType === 'ul' ? '</ul>' : '</ol>';
    if (inBlockquote) html += '</blockquote>';
    if (inDetails) html += '</div></details>';

    return html;
}

/** Process inline markdown: bold, italic, code, images, links */
function inline(text) {
    return text
        // Images: ![alt](src)
        .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img src="$2" alt="$1" loading="lazy">')
        // Links: [text](url)
        .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>')
        // Bold: **text**
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        // Italic: *text*
        .replace(/\*(.+?)\*/g, '<em>$1</em>')
        // Inline code: `code`
        .replace(/`(.+?)`/g, '<code>$1</code>');
}

function escapeHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
