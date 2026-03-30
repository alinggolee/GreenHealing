/**
 * Hash-based Router
 * Routes: #/ | #/lesson/:id | #/lesson/:id/:section
 */

export class Router {
  constructor() {
    this.routes = [];
    this._onHashChange = this._onHashChange.bind(this);
  }

  /**
   * Register a route pattern.
   * @param {string} pattern - e.g. '/lesson/:id/:section'
   * @param {Function} handler - callback(params)
   */
  on(pattern, handler) {
    const parts = pattern.split('/').filter(Boolean);
    this.routes.push({ parts, handler });
  }

  /** Start listening to hash changes */
  start() {
    window.addEventListener('hashchange', this._onHashChange);
    this._onHashChange();
  }

  /** Stop listening */
  stop() {
    window.removeEventListener('hashchange', this._onHashChange);
  }

  /** Navigate to a hash */
  static navigate(hash) {
    window.location.hash = hash;
  }

  /** @private */
  _onHashChange() {
    const hash = window.location.hash.slice(1) || '/';
    const hashParts = hash.split('/').filter(Boolean);

    for (const route of this.routes) {
      const params = this._match(route.parts, hashParts);
      if (params !== null) {
        route.handler(params);
        return;
      }
    }

    // Fallback to home
    Router.navigate('#/');
  }

  /**
   * @private
   * Match route parts against hash parts, return params or null.
   */
  _match(routeParts, hashParts) {
    if (routeParts.length !== hashParts.length) return null;

    const params = {};
    for (let i = 0; i < routeParts.length; i++) {
      if (routeParts[i].startsWith(':')) {
        params[routeParts[i].slice(1)] = decodeURIComponent(hashParts[i]);
      } else if (routeParts[i] !== hashParts[i]) {
        return null;
      }
    }
    return params;
  }
}
