// Animation functionality
document.addEventListener('DOMContentLoaded', () => {
    
    // Animate elements on scroll
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
            }
        });
    }, observerOptions);

    // Observe all cards
    document.querySelectorAll('.card').forEach(el => {
        observer.observe(el);
    });

    // Counter animation for stats
    const animateCounter = (element, target, suffix) => {
        let current = 0;
        const duration = 1500; // ms
        const steps = 60;
        const increment = target / steps;
        const stepTime = Math.abs(Math.floor(duration / steps));
        
        const timer = setInterval(() => {
            current += increment;
            if (current >= target) {
                element.textContent = target + suffix;
                clearInterval(timer);
            } else {
                element.textContent = Math.floor(current) + suffix;
            }
        }, stepTime);
    };

    const statsObserver = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting && !entry.target.dataset.animated) {
                const numberEl = entry.target.querySelector('.stat-number');
                const targetValue = parseInt(numberEl.dataset.target);
                const suffix = numberEl.dataset.suffix || '';
                
                entry.target.dataset.animated = 'true';
                animateCounter(numberEl, targetValue, suffix);
            }
        });
    }, { threshold: 0.5 });

    document.querySelectorAll('.stat-item').forEach(stat => {
        statsObserver.observe(stat);
    });

});