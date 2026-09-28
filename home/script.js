document.addEventListener("DOMContentLoaded", () => {
  const toolCards = document.querySelectorAll(".tool-card");

  // 1. Add visual feedback on click
  toolCards.forEach((card) => {
    card.addEventListener("click", (e) => {
      // Add a brief pulse animation effect on launch
      card.style.transform = "scale(0.98)";
      card.style.opacity = "0.8";
    });
  });

  // 2. Keyboard shortcuts to launch tools quickly
  // Pressing '1' launches PPTX Viewer, '2' launches PDF Reader, etc.
  document.addEventListener("keydown", (e) => {
    // Ignore keypresses if user is typing into an input field
    if (["INPUT", "TEXTAREA"].includes(document.activeElement.tagName)) return;

    const keyNumber = parseInt(e.key, 10);
    if (!isNaN(keyNumber) && keyNumber > 0 && keyNumber <= toolCards.length) {
      const targetCard = toolCards[keyNumber - 1];
      if (targetCard && targetCard.getAttribute("href")) {
        targetCard.click(); // Trigger click & navigation
      }
    }
  });

  // 3. Optional: Add active keyboard accessibility indicator
  toolCards.forEach((card, index) => {
    const shortcutBadge = document.createElement("span");
    shortcutBadge.className = "shortcut-badge";
    shortcutBadge.textContent = `[${index + 1}]`;
    shortcutBadge.style.cssText = `
      position: absolute;
      top: 12px;
      right: 12px;
      font-size: 0.75rem;
      color: #64748b;
      font-family: monospace;
    `;
    card.style.position = "relative";
    card.appendChild(shortcutBadge);
  });
});
