document.addEventListener("DOMContentLoaded", () => {
    const root = document.documentElement;
    const storageKey = "axonasp-theme";
    const themeSelect = document.querySelector("[data-theme-select]");
    const menuToggle = document.querySelector("[data-menu-toggle]");
    const menuShell = document.querySelector("[data-menu-shell]");

    const setTheme = (theme) => {
        if (!theme) {
            root.removeAttribute("data-theme");
            localStorage.removeItem(storageKey);
            return;
        }
        root.setAttribute("data-theme", theme);
        localStorage.setItem(storageKey, theme);
    };

    const savedTheme = localStorage.getItem(storageKey);
    if (savedTheme === "light" || savedTheme === "dark") {
        setTheme(savedTheme);
    }

    if (themeSelect) {
        themeSelect.value = savedTheme === "light" || savedTheme === "dark" ? savedTheme : "auto";
    }

    if (themeSelect) {
        themeSelect.addEventListener("change", (event) => {
            const selected = event.target.value;
            if (selected === "auto") {
                setTheme(null);
            } else {
                setTheme(selected);
            }
        });
    }

    if (menuToggle && menuShell) {
        menuToggle.addEventListener("click", () => {
            const isOpen = menuShell.classList.toggle("open");
            menuToggle.setAttribute("aria-expanded", isOpen ? "true" : "false");
        });
    }

    const currentPage = window.location.pathname.split("/").pop() || "index.htm";
    document.querySelectorAll(".nav-menu a").forEach((link) => {
        const href = link.getAttribute("href");
        if (!href || href.startsWith("http")) {
            return;
        }
        if (href === currentPage) {
            link.classList.add("active");
            link.setAttribute("aria-current", "page");
        }
    });

    const revealElements = document.querySelectorAll(".reveal");
    if ("IntersectionObserver" in window) {
        const revealObserver = new IntersectionObserver((entries) => {
            entries.forEach((entry) => {
                if (entry.isIntersecting) {
                    entry.target.classList.add("is-visible");
                    revealObserver.unobserve(entry.target);
                }
            });
        }, { threshold: 0.2 });

        revealElements.forEach((el) => revealObserver.observe(el));
    } else {
        revealElements.forEach((el) => el.classList.add("is-visible"));
    }

    const animateCounter = (element, target, suffix = "") => {
        const duration = 1400;
        const start = performance.now();
        const step = (timestamp) => {
            const progress = Math.min((timestamp - start) / duration, 1);
            const value = Math.floor(progress * target);
            element.textContent = `${value}${suffix}`;
            if (progress < 1) {
                requestAnimationFrame(step);
            } else {
                element.textContent = `${target}${suffix}`;
            }
        };
        requestAnimationFrame(step);
    };

    const statCards = document.querySelectorAll(".stat");
    if ("IntersectionObserver" in window) {
        const statsObserver = new IntersectionObserver((entries) => {
            entries.forEach((entry) => {
                if (!entry.isIntersecting || entry.target.dataset.counted === "true") {
                    return;
                }

                const number = entry.target.querySelector(".stat-number");
                if (!number) {
                    return;
                }

                const target = parseInt(number.dataset.target || "0", 10);
                const suffix = number.dataset.suffix || "";
                entry.target.dataset.counted = "true";
                animateCounter(number, target, suffix);
            });
        }, { threshold: 0.5 });

        statCards.forEach((card) => statsObserver.observe(card));
    }
});