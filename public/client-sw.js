// Service worker for client.html (employee) push notifications only - no
// offline caching, deliberately minimal so it can never serve stale app
// code to an employee. Registered from client.html once an employee logs in
// and grants notification permission - see subscribeEmployeeToPush() there.
// Mirrors driver-sw.js exactly, just pointed at client.html instead.

self.addEventListener("install", () => {
  self.skipWaiting();
});
self.addEventListener("activate", (event) => {
  event.waitUntil(self.clients.claim());
});

// Fired when the backend sends a push (e.g. the driver just arrived at this
// employee's pickup point - see sendPushToEmployee() in server.js). Shows a
// native notification even if client.html isn't open.
self.addEventListener("push", (event) => {
  let data = { title: "Padmasri Travels", body: "You have a new update.", url: "/client.html" };
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
      data: { url: data.url || "/client.html" },
      tag: "padmasri-employee-ride",
      renotify: true,
    })
  );
});

// Tapping the notification focuses an already-open client.html tab if one
// exists, otherwise opens a new one.
self.addEventListener("notificationclick", (event) => {
  event.notification.close();
  const targetUrl = (event.notification.data && event.notification.data.url) || "/client.html";
  event.waitUntil(
    self.clients.matchAll({ type: "window", includeUncontrolled: true }).then((clientsList) => {
      for (const client of clientsList) {
        if (client.url.includes("client.html") && "focus" in client) return client.focus();
      }
      if (self.clients.openWindow) return self.clients.openWindow(targetUrl);
    })
  );
});
