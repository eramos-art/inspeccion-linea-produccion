// Service worker mínimo: permite instalar la app en el celular.
// No cachea datos (las inspecciones siempre se leen en vivo desde Google Sheets).
self.addEventListener('install', e => self.skipWaiting());
self.addEventListener('activate', e => self.clients.claim());
self.addEventListener('fetch', () => {}); // deja pasar todas las peticiones normal
