self.addEventListener('push', function(event) {
  const data = event.data ? event.data.json() : {};
  const title = data.title || 'Расписание ААСК';
  const options = {
    body: data.body || 'Проверьте расписание!',
    icon: '/icons/apple-touch-icon.png',
    badge: '/icons/apple-touch-icon.png',
    vibrate: [200, 100, 200]
  };
  event.waitUntil(self.registration.showNotification(title, options));
});