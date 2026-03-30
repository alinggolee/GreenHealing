/**
 * Theme utilities — dark/light toggle + per-lesson color mapping
 */

/** Lesson ID → color name mapping */
const lessonColorMap = {
    W1: 'green', W2: 'green', W3: 'green',
    W4: 'blue', W5: 'blue', W6: 'blue',
    W7: 'purple', W8: 'purple', W9: 'purple',
    W10: 'red', W11: 'red', W12: 'red', W13: 'red',
    W14: 'yellow', W15: 'yellow'
};

/**
 * Get the color theme name for a lesson ID.
 * @param {string} lessonId - e.g. 'W3'
 * @returns {string} color name
 */
export function getLessonColor(lessonId) {
    return lessonColorMap[lessonId] || 'purple';
}

/**
 * Apply lesson color to <html> element.
 * @param {string|null} lessonId
 */
export function applyLessonColor(lessonId) {
    if (lessonId) {
        document.documentElement.setAttribute('data-color', getLessonColor(lessonId));
    } else {
        document.documentElement.removeAttribute('data-color');
    }
}

/**
 * Initialize dark/light theme toggle.
 * Reads from localStorage, defaults to light.
 */
export function initThemeToggle() {
    const saved = localStorage.getItem('theme');
    if (saved === 'dark') {
        document.documentElement.setAttribute('data-theme', 'dark');
    }
}

/**
 * Toggle theme and return new state
 * @returns {boolean} true if dark mode is now active
 */
export function toggleTheme() {
    const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    if (isDark) {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('theme', 'light');
        return false;
    } else {
        document.documentElement.setAttribute('data-theme', 'dark');
        localStorage.setItem('theme', 'dark');
        return true;
    }
}

export function isDarkTheme() {
    return document.documentElement.getAttribute('data-theme') === 'dark';
}
