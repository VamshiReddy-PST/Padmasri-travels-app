// Service worker for driver.html push notifications only (no offline
// caching - keeping this deliberately minimal so it can't ever serve stale
// app code to a driver). Registered from driver.html once a driver logs in
// and grants notification permission - see subscribeToPush() there.

self.addEventListener("install", () => {
  self.skipWaiting();
});
self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

// Fired when the backend sends a push (e.g. a route was just dispatched to
// this driver - see sendPushToDriver() in server.js). Shows a native
// notification even if driver.html isn't open.
self.addEventListener("push", (event) => {
  let data = { title: "Padmasri Travels", body: "You have a new update.", url: "/driver.html" };
  try {
    if (event.data) data = Object.assign(data, event.data.json());
  } catch (e) {
    // non-JSON payload - fall back to the defaults above
  }
  event.waitUntil(
    self.registration.showNotification(data.title, {
      body: data.body,
      icon: "/icon.svg",
      badge: "/icon.svg",
      data: { url: data.url || "/driver.html" },
      tag: "padmasri-driver-trip",
      renotify: true,
    })
  );
});

// Tapping the notification focuses an already-open driver.html tab if one
// exists, otherwise opens a new one.
self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || "/driver.html";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientsList) => {
      for (const client of clientsList) {
        if (client.url.includes("driver.html") && "focus" in client) return client.focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow(targetUrl);
    })
  );
});
