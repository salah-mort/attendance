// Smooth scroll to features
function scrollToFeatures() {
  document.getElementById("features").scrollIntoView({ behavior: "smooth" });
}

// Intersection Observer للتأثيرات عند التمرير
const observerOptions = {
  threshold: 0.1,
  rootMargin: "0px 0px -100px 0px",
};

const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) {
      entry.target.style.animation = `slideInUp 0.6s ease-out forwards`;
      observer.unobserve(entry.target);
    }
  });
}, observerOptions);

// إضافة التأثير لعناصر معينة
document.addEventListener("DOMContentLoaded", () => {
  const cards = document.querySelectorAll(".feature-card, .step, .stat-box");
  cards.forEach((card, index) => {
    card.style.opacity = "0";
    card.style.animation = `slideInUp 0.6s ease-out ${index * 0.1}s forwards`;
  });
});

// إضافة animation keyframe
const style = document.createElement("style");
style.textContent = `
    @keyframes slideInUp {
        from {
            opacity: 0;
            transform: translateY(30px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }
`;
document.head.appendChild(style);

// Parallax effect
window.addEventListener("scroll", () => {
  const scrolled = window.pageYOffset;
  const stars = document.querySelector(".stars");
  if (stars) {
    stars.style.transform = `translateY(${scrolled * 0.5}px)`;
  }
});

// عد تنازلي للميزات
let countAnimationStarted = false;
window.addEventListener("scroll", () => {
  const statsSection = document.querySelector(".stats-section");
  if (!statsSection) return;

  const rect = statsSection.getBoundingClientRect();
  if (rect.top < window.innerHeight && !countAnimationStarted) {
    countAnimationStarted = true;
    animateStats();
  }
});

function animateStats() {
  const statNumbers = document.querySelectorAll(".stat-number");
  statNumbers.forEach((stat) => {
    const finalValue = stat.textContent;
    const numValue = parseInt(finalValue);

    if (!isNaN(numValue)) {
      let currentValue = 0;
      const increment = Math.ceil(numValue / 50);
      const interval = setInterval(() => {
        currentValue += increment;
        if (currentValue >= numValue) {
          stat.textContent = finalValue;
          clearInterval(interval);
        } else {
          stat.textContent = currentValue + "+";
        }
      }, 20);
    }
  });
}
