document.addEventListener('DOMContentLoaded', () => {
  // ==========================================
  // 1. Header Scroll Effect
  // ==========================================
  const header = document.querySelector('header');
  
  const handleScroll = () => {
    if (window.scrollY > 50) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
  };

  window.addEventListener('scroll', handleScroll);
  // Initial check in case page is refreshed while scrolled
  handleScroll();

  // ==========================================
  // 2. Mobile Menu Toggle
  // ==========================================
  const burger = document.querySelector('.burger');
  const navLinks = document.querySelector('.nav-links');
  const navLinksItems = document.querySelectorAll('.nav-links a');

  burger.addEventListener('click', () => {
    navLinks.classList.toggle('open');
    burger.classList.toggle('toggle');
  });

  // Close mobile menu when clicking any nav link
  navLinksItems.forEach(link => {
    link.addEventListener('click', () => {
      navLinks.classList.remove('open');
      burger.classList.remove('toggle');
    });
  });

  // ==========================================
  // 3. Scroll Reveal Animation
  // ==========================================
  const revealElements = document.querySelectorAll('.reveal');

  const revealObserver = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('active');
        // Once revealed, we can unobserve if we only want it to animate once
        observer.unobserve(entry.target);
      }
    });
  }, {
    threshold: 0.15, // Trigger when 15% of the element is visible
    rootMargin: '0px 0px -50px 0px' // Adjust triggers to fire slightly early
  });

  revealElements.forEach(element => {
    revealObserver.observe(element);
  });

  // ==========================================
  // 4. FAQ Accordion Logic
  // ==========================================
  const faqItems = document.querySelectorAll('.faq-item');

  faqItems.forEach(item => {
    const trigger = item.querySelector('.faq-trigger');
    const panel = item.querySelector('.faq-panel');

    trigger.addEventListener('click', () => {
      const isActive = item.classList.contains('active');

      // Close all other items first (accordion style)
      faqItems.forEach(otherItem => {
        otherItem.classList.remove('active');
        otherItem.querySelector('.faq-panel').style.maxHeight = null;
      });

      if (!isActive) {
        item.classList.add('active');
        // Dynamically set height to scrollHeight for smooth transition
        panel.style.maxHeight = panel.scrollHeight + 'px';
      } else {
        item.classList.remove('active');
        panel.style.maxHeight = null;
      }
    });
  });

  // ==========================================
  // 5. Product Sizing Selector & Checkout Logic (Both Pink)
  // ==========================================
  const sizeTabs = document.querySelectorAll('.size-tab');
  const displayImage = document.getElementById('product-image');
  const buyBtn = document.getElementById('btn-buy-whatsapp');
  
  const specCapacity = document.getElementById('spec-capacity');
  const specDiameter = document.getElementById('spec-diameter');
  const specLength = document.getElementById('spec-length');
  const specIdealFor = document.getElementById('spec-ideal-for');

  const WHATSAPP_PHONE = '919558642815';

  const productData = {
    small: {
      image: 'assets/box_small.jpg',
      capacity: '25 ml',
      diameter: '42 mm',
      length: '65 mm',
      idealFor: 'Light Flow / Pre-Childbirth',
      message: "Hi Gallops! I'm interested in ordering the Gallops Menstrual Cup in Size Small. Please guide me on payment and delivery."
    },
    medium: {
      image: 'assets/box_medium.jpg',
      capacity: '28 ml',
      diameter: '43 mm',
      length: '68 mm',
      idealFor: 'Medium Flow / Universal Fit',
      message: "Hi Gallops! I'm interested in ordering the Gallops Menstrual Cup in Size Medium. Please guide me on payment and delivery."
    },
    large: {
      image: 'assets/box_large.jpg',
      capacity: '30 ml',
      diameter: '45 mm',
      length: '70 mm',
      idealFor: 'Heavy Flow / Post-Childbirth',
      message: "Hi Gallops! I'm interested in ordering the Gallops Menstrual Cup in Size Large. Please guide me on payment and delivery."
    }
  };

  const thumbBoxImg = document.getElementById('thumb-box-img');
  const thumbBtns = document.querySelectorAll('.thumb-btn');

  if (sizeTabs.length > 0) {
    sizeTabs.forEach(tab => {
      tab.addEventListener('click', () => {
        const selectedSize = tab.dataset.size;
        
        sizeTabs.forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        const data = productData[selectedSize];
        if (!data) return;

        // Reset active thumbnail to box packaging
        if (thumbBtns.length > 0) {
          thumbBtns.forEach(b => b.classList.remove('active'));
          const boxBtn = document.querySelector('.thumb-btn[data-src="box"]');
          if (boxBtn) boxBtn.classList.add('active');
        }

        // Update box thumbnail image source to match the active size packaging
        if (thumbBoxImg) {
          thumbBoxImg.src = data.image;
        }

        displayImage.classList.add('fade-out');
        
        setTimeout(() => {
          displayImage.src = data.image;
          displayImage.alt = `Gallops Menstrual Cup - Size ${selectedSize.toUpperCase()} (Pink)`;
          
          specCapacity.textContent = data.capacity;
          specDiameter.textContent = data.diameter;
          specLength.textContent = data.length;
          specIdealFor.textContent = data.idealFor;

          const encodedText = encodeURIComponent(data.message);
          buyBtn.href = `https://wa.me/${WHATSAPP_PHONE}?text=${encodedText}`;

          displayImage.classList.remove('fade-out');
        }, 300);
      });
    });
  }

  // Thumbnails toggle logic
  if (thumbBtns.length > 0) {
    thumbBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        thumbBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        const srcType = btn.dataset.src;
        let newSrc = srcType;

        if (srcType === 'box') {
          const activeTab = document.querySelector('.size-tab.active');
          const activeSize = activeTab ? activeTab.dataset.size : 'small';
          newSrc = productData[activeSize].image;
        }

        displayImage.classList.add('fade-out');
        setTimeout(() => {
          displayImage.src = newSrc;
          displayImage.classList.remove('fade-out');
        }, 300);
      });
    });
  }

  // Initialize checkout link for default size (Small)
  if (buyBtn) {
    const defaultData = productData.small;
    const encodedText = encodeURIComponent(defaultData.message);
    buyBtn.href = `https://wa.me/${WHATSAPP_PHONE}?text=${encodedText}`;
  }

  // ==========================================
  // 5.5. DTDC Order Tracking Form Submission
  // ==========================================
  const trackingForm = document.getElementById('tracking-form');
  const trackingInput = document.getElementById('tracking-id-input');
  const trackingStatus = document.getElementById('tracking-status-msg');

  if (trackingForm) {
    trackingForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const trackingId = trackingInput.value.trim();
      if (trackingId) {
        // Copy tracking ID to clipboard automatically
        navigator.clipboard.writeText(trackingId).then(() => {
          if (trackingStatus) {
            trackingStatus.textContent = `📋 Consignment ID "${trackingId}" copied to clipboard! Opening DTDC portal... please paste it in the search box.`;
            trackingStatus.style.display = 'block';
          }
        }).catch(err => {
          if (trackingStatus) {
            trackingStatus.textContent = `Opening DTDC portal for tracking ID: ${trackingId}`;
            trackingStatus.style.display = 'block';
          }
        });

        // Open DTDC portal immediately in a new tab to bypass popup blockers
        window.open('https://www.dtdc.in/', '_blank');
      }
    });
  }

  // ==========================================
  // 6. Navigation Active State on Scroll
  // ==========================================
  const sections = document.querySelectorAll('section[id]');
  const navA = document.querySelectorAll('.nav-links a');

  const handleActiveNav = () => {
    let current = '';
    const scrollPos = window.scrollY + 120; // offset header

    sections.forEach(section => {
      const sectionTop = section.offsetTop;
      const sectionHeight = section.clientHeight;
      if (scrollPos >= sectionTop && scrollPos < sectionTop + sectionHeight) {
        current = section.getAttribute('id');
      }
    });

    navA.forEach(a => {
      a.classList.remove('active');
      if (a.getAttribute('href') === `#${current}`) {
        a.classList.add('active');
      }
    });
  };

  window.addEventListener('scroll', handleActiveNav);
  handleActiveNav();
});

