// Polyfills window.storage (Claude's artifact storage API) using localStorage,
// so the app's persistence code works unmodified outside Claude.
// Note: this is per-browser, per-device storage — it won't sync across devices.
// The `shared` parameter from the original API is accepted but ignored, since
// there's no multi-user backend in this standalone build.

const PREFIX = "suturelog:";

function fullKey(key) {
  return PREFIX + key;
}

window.storage = {
  async get(key) {
    const raw = localStorage.getItem(fullKey(key));
    if (raw === null) return null;
    return { key, value: raw };
  },

  async set(key, value) {
    localStorage.setItem(fullKey(key), value);
    return { key, value };
  },

  async delete(key) {
    const existed = localStorage.getItem(fullKey(key)) !== null;
    localStorage.removeItem(fullKey(key));
    return { key, deleted: existed };
  },

  async list(prefix) {
    const keys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith(PREFIX)) {
        const bare = k.slice(PREFIX.length);
        if (!prefix || bare.startsWith(prefix)) keys.push(bare);
      }
    }
    return { keys, prefix };
  }
};
