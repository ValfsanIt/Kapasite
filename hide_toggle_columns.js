/**
 * Kapasite sayfasında DataTable'ın "Toggle Columns" UI öğesini gizler.
 * Kolonları Gizle butonuna tıklanınca DataTable gizli kolonlar için bu arayüzü gösteriyor;
 * bu script onu bulup gizliyor.
 */
(function () {
    var timer = null;
  
    function hideEl(el) {
      el.style.setProperty('display', 'none', 'important');
      el.style.setProperty('visibility', 'hidden', 'important');
      el.style.setProperty('width', '0', 'important');
      el.style.setProperty('height', '0', 'important');
      el.style.setProperty('overflow', 'hidden', 'important');
      el.style.setProperty('position', 'absolute', 'important');
      el.style.setProperty('left', '-9999px', 'important');
      el.setAttribute('aria-hidden', 'true');
    }
  
    function hideToggleColumnsElements() {
      var containers = document.querySelectorAll('.kap-page-container .dash-table-container, .kap-fullscreen-panel .dash-table-container');
      for (var c = 0; c < containers.length; c++) {
        var all = containers[c].querySelectorAll('*');
        for (var i = 0; i < all.length; i++) {
          var el = all[i];
          var text = (el.textContent || el.innerText || '').trim();
          if (text === 'Toggle Columns') {
            hideEl(el);
            var parent = el.parentElement;
            if (parent && parent !== containers[c] && (parent.textContent || '').trim() === 'Toggle Columns') {
              hideEl(parent);
            }
          }
        }
      }
    }
  
    function run() {
      hideToggleColumnsElements();
    }
  
    function debouncedRun() {
      if (timer) clearTimeout(timer);
      timer = setTimeout(run, 100);
    }
  
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', function () {
        run();
        setTimeout(run, 300);
        setTimeout(run, 1000);
      });
    } else {
      run();
      setTimeout(run, 300);
      setTimeout(run, 1000);
    }
  
    var observer = new MutationObserver(debouncedRun);
    observer.observe(document.body, { childList: true, subtree: true });
  })();
  