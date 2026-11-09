// 앱 화면 파일을 캐시해서 홈 화면 앱으로 빠르게 열리게 함 (일정·메모 데이터는 캐시하지 않음)
// + 다른 앱(클로바노트·녹음 앱)에서 "공유"로 보낸 녹취를 받아 앱에 넘김
const C = "ndrnote-v34", SHARE = "ndr-share";
const FILES = ["./", "index.html", "style.css", "app.js", "parser.js", "stt-worker.js", "manifest.webmanifest", "icon-192.png", "icon-512.png"];
self.addEventListener("install", e => e.waitUntil(caches.open(C).then(c => c.addAll(FILES)).then(() => self.skipWaiting())));
self.addEventListener("activate", e => e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith("ndrnote-") && k !== C).map(k => caches.delete(k)))).then(() => self.clients.claim())));
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method === "POST" && url.pathname.endsWith("/share")) {
    e.respondWith((async () => {
      try {
        const f = await e.request.formData();
        const text = String(f.get("text") || "").trim();
        const title = String(f.get("title") || "").trim();
        const items = [];
        for (const file of f.getAll("file")) {
          if (!file || !file.size) continue;
          if (/^(audio|video)\//.test(file.type) || /\.(m4a|mp3|wav|aac|amr|3gp|ogg|opus|flac)$/i.test(file.name || "")) {
            const key = "audio-" + items.length + "-" + Date.now();   // 녹음 파일은 잠깐 보관했다가 앱이 텍스트로 바꾸면 지움
            await (await caches.open(SHARE)).put(key, new Response(file, { headers: { "content-type": file.type || "audio/mp4" } }));
            items.push({ name: String(file.name || "녹음"), audioKey: key, type: file.type || "audio/mp4", lastModified: file.lastModified || Date.now() });
          } else if (!/^image\//.test(file.type)) items.push({ name: String(file.name || "").replace(/\.(txt|text)$/i, ""), text: (await file.text()).trim() });
        }
        const c = await caches.open(SHARE);
        await c.put("shared", new Response(JSON.stringify({ title, text, items, at: Date.now() }), { headers: { "content-type": "application/json" } }));
      } catch (err) { }
      return Response.redirect(new URL("./?shared=1", self.registration.scope).href, 303);
    })());
    return;
  }
  if (e.request.method !== "GET" || !e.request.url.startsWith(self.location.origin)) return;
  e.respondWith(fetch(e.request, { cache: "no-cache" }).then(r => { const copy = r.clone(); caches.open(C).then(c => c.put(e.request, copy)); return r; }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
