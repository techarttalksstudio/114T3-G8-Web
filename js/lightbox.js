(function () {
  function init() {
    const targets = document.querySelectorAll('.slide-images--gallery img');
    if (!targets.length) return;

    const overlay = document.createElement('div');
    overlay.className = 'img-lightbox';
    overlay.innerHTML = '<button type="button" class="img-lightbox__close" aria-label="關閉">×</button><img alt="" />';
    document.body.appendChild(overlay);

    const bigImg = overlay.querySelector('img');
    const closeBtn = overlay.querySelector('.img-lightbox__close');

    function open(src, alt) {
      bigImg.src = src;
      bigImg.alt = alt || '';
      overlay.classList.add('is-open');
    }
    function close() {
      overlay.classList.remove('is-open');
      bigImg.src = '';
    }

    targets.forEach((el) => {
      el.addEventListener('click', () => open(el.src, el.alt));
    });

    overlay.addEventListener('click', (e) => {
      if (e.target === overlay) close();
    });
    closeBtn.addEventListener('click', close);
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') close();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
