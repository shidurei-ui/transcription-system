function getPageVideoInfo() {
  const info = { pageUrl: window.location.href, title: document.title, directUrls: [], hasVideo: false };
  document.querySelectorAll("video").forEach((v) => {
    info.hasVideo = true;
    const src = v.src || v.currentSrc;
    if (src && !src.startsWith("blob:") && !src.startsWith("data:")) info.directUrls.push(src);
    v.querySelectorAll("source").forEach((s) => { if (s.src) info.directUrls.push(s.src); });
  });
  document.querySelectorAll("iframe").forEach((f) => {
    const src = f.src || "";
    if (src.match(/youtube\.com\/embed|youtu\.be/)) {
      info.hasVideo = true;
      const videoId = src.match(/embed\/([^?&]+)/)?.[1];
      if (videoId) info.directUrls.push(`https://www.youtube.com/watch?v=${videoId}`);
    }
  });
  info.directUrls = [...new Set(info.directUrls)];
  return info;
}

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.type === "GET_VIDEO_INFO") { sendResponse(getPageVideoInfo()); }
  return true;
});
