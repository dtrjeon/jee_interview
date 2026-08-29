const CACHE_NAME = "ee-interview-checklist-v2";
const ASSETS = [
  "./index.html",
  "./ee_interview_checklist.html",
  "./manifest.json"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  const isHtml = event.request.headers.get("accept")?.includes("text/html")
    || event.request.url.endsWith(".html");

  if (isHtml) {
    // HTML은 네트워크(최신 파일) 우선 — 파일을 고칠 때마다 버전을 안 올려도 바로 반영됨.
    // 오프라인일 때만 캐시된 예전 버전을 보여줌.
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  // 그 외 정적 파일(아이콘, manifest 등)은 기존처럼 캐시 우선
  event.respondWith(
    caches.match(event.request).then((cached) => {
      return (
        cached ||
        fetch(event.request).then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
      );
    })
  );
});
