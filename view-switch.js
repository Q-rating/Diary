(function () {
  try {
    var p = location.pathname;
    if (!/(\/|index\.html|mobile\.html)$/.test(p)) return;
    var q = new URLSearchParams(location.search).get('view');
    if (q === 'auto') localStorage.removeItem('diary.viewPref');
    else if (q === 'mobile' || q === 'desktop') localStorage.setItem('diary.viewPref', q);
    var pref = localStorage.getItem('diary.viewPref');
    var ua = navigator.userAgent;
    var mob = /Android|iPhone|iPad|iPod|Mobile/i.test(ua) || (navigator.maxTouchPoints > 1 && /Macintosh/.test(ua));
    var want = pref || (mob ? 'mobile' : 'desktop');
    var onM = /mobile\.html$/.test(p);
    if (want === 'mobile' && !onM) location.replace(p.replace(/(index\.html)?$/, 'mobile.html'));
    else if (want === 'desktop' && onM) location.replace(p.replace(/mobile\.html$/, ''));
  } catch (e) {}
})();
