/* ============================================================
   DATA VISTA 2.0 — Main Script
   Countdown timer, scroll animations, navbar effects
   ============================================================ */

'use strict';

// ── DOM Ready ──
document.addEventListener('DOMContentLoaded', () => {
    initNavbar();
    initMobileMenu();
    initCountdown();
    initScrollAnimations();
    initSmoothScroll();
    initEventCardEffects();
});

// ── Navbar Scroll Effect ──
function initNavbar() {
    const navbar = document.getElementById('navbar');
    if (!navbar) return;

    const handleScroll = () => {
        if (window.scrollY > 60) {
            navbar.classList.add('scrolled');
        } else {
            navbar.classList.remove('scrolled');
        }
    };

    window.addEventListener('scroll', handleScroll, { passive: true });
    handleScroll();
}

// ── Mobile Menu Toggle ──
function initMobileMenu() {
    const toggle = document.getElementById('mobile-toggle');
    const menu   = document.getElementById('mobile-menu');
    if (!toggle || !menu) return;

    toggle.addEventListener('click', () => {
        const isOpen = menu.classList.toggle('open');
        toggle.innerHTML = isOpen
            ? '<i class="fas fa-times"></i>'
            : '<i class="fas fa-bars"></i>';
        toggle.setAttribute('aria-expanded', isOpen);
    });

    // Close menu when a link is clicked
    menu.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', () => {
            menu.classList.remove('open');
            toggle.innerHTML = '<i class="fas fa-bars"></i>';
            toggle.setAttribute('aria-expanded', false);
        });
    });
}

// ── Countdown Timer ──
function initCountdown() {
    // Event: 31st August 2026, 9:00 AM IST (DATA VISTA '26)
    const eventDate = new Date('2026-08-31T09:00:00+05:30').getTime();

    const daysEl    = document.getElementById('days');
    const hoursEl   = document.getElementById('hours');
    const minutesEl = document.getElementById('minutes');
    const secondsEl = document.getElementById('seconds');

    if (!daysEl || !hoursEl || !minutesEl || !secondsEl) return;

    function pad(n) {
        return String(Math.max(0, n)).padStart(2, '0');
    }

    function animateNumber(el, newVal) {
        if (el.textContent !== newVal) {
            el.style.transform = 'translateY(-8px)';
            el.style.opacity   = '0';
            el.style.transition = 'all 0.2s ease';
            setTimeout(() => {
                el.textContent = newVal;
                el.style.transform = 'translateY(0)';
                el.style.opacity   = '1';
            }, 200);
        }
    }

    function tick() {
        const now  = Date.now();
        const diff = eventDate - now;

        if (diff <= 0) {
            // Event has passed
            ['days','hours','minutes','seconds'].forEach(id => {
                const el = document.getElementById(id);
                if (el) el.textContent = '00';
            });
            const titleEl = document.querySelector('.countdown-title');
            if (titleEl) {
                titleEl.innerHTML = '<i class="fas fa-flag-checkered"></i> The Event Has Concluded!';
            }
            return;
        }

        const days    = Math.floor(diff / (1000 * 60 * 60 * 24));
        const hours   = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
        const seconds = Math.floor((diff % (1000 * 60)) / 1000);

        animateNumber(daysEl,    pad(days));
        animateNumber(hoursEl,   pad(hours));
        animateNumber(minutesEl, pad(minutes));
        animateNumber(secondsEl, pad(seconds));
    }

    tick();
    setInterval(tick, 1000);
}

// ── Scroll-triggered Reveal Animations ──
function initScrollAnimations() {
    // Add 'reveal' class to animatable elements
    const targets = document.querySelectorAll(
        '.event-card, .contact-item, .about-text p, .about-stats, ' +
        '.section-header, .countdown-container, .map-placeholder, .stat-item'
    );

    targets.forEach((el, i) => {
        el.classList.add('reveal');
        el.style.transitionDelay = `${(i % 5) * 0.08}s`;
    });

    const observer = new IntersectionObserver(
        (entries) => {
            entries.forEach(entry => {
                if (entry.isIntersecting) {
                    entry.target.classList.add('visible');
                    observer.unobserve(entry.target);
                }
            });
        },
        { threshold: 0.1, rootMargin: '0px 0px -50px 0px' }
    );

    targets.forEach(el => observer.observe(el));
}

// ── Smooth Scroll for Nav Links ──
function initSmoothScroll() {
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            const target = document.querySelector(this.getAttribute('href'));
            if (!target) return;
            e.preventDefault();
            const navHeight = document.querySelector('.navbar')?.offsetHeight || 80;
            const top = target.getBoundingClientRect().top + window.scrollY - navHeight;
            window.scrollTo({ top, behavior: 'smooth' });
        });
    });
}

// ── Event Card Hover Glow ──
function initEventCardEffects() {
    document.querySelectorAll('.event-card').forEach(card => {
        card.addEventListener('mousemove', (e) => {
            const rect  = card.getBoundingClientRect();
            const x     = ((e.clientX - rect.left) / rect.width)  * 100;
            const y     = ((e.clientY - rect.top)  / rect.height) * 100;
            const glow  = card.querySelector('.card-glow');
            if (glow) {
                glow.style.background = `radial-gradient(circle at ${x}% ${y}%, rgba(0,212,255,0.08) 0%, transparent 60%)`;
            }
        });
    });
}

// ── Active Nav Highlight on Scroll ──
(function () {
    const sections = document.querySelectorAll('section[id]');
    const navLinks = document.querySelectorAll('.nav-link, .mobile-menu .nav-link');

    function highlight() {
        const scrollY = window.scrollY + 120;
        sections.forEach(section => {
            const top    = section.offsetTop;
            const height = section.offsetHeight;
            const id     = section.getAttribute('id');

            if (scrollY >= top && scrollY < top + height) {
                navLinks.forEach(link => {
                    link.classList.remove('active');
                    if (link.getAttribute('href') === `#${id}`) {
                        link.classList.add('active');
                    }
                });
            }
        });
    }

    window.addEventListener('scroll', highlight, { passive: true });
})();

// ── Add active nav link style ──
const navStyle = document.createElement('style');
navStyle.textContent = `
    .nav-link.active {
        color: var(--gold-light) !important;
    }
    .nav-link.active::after {
        width: 100% !important;
    }
`;
document.head.appendChild(navStyle);
