document.addEventListener("DOMContentLoaded", () => {
  const slideStage = document.getElementById("slide-stage");
  const prevBtn = document.getElementById("prev-btn");
  const nextBtn = document.getElementById("next-btn");
  const currentSlideEl = document.getElementById("current-slide");
  const totalSlidesEl = document.getElementById("total-slides");
  const fileBtns = document.querySelectorAll(".file-btn");
  const fileInput = document.getElementById("file-input");

  let currentSlideIndex = 0;
  let totalSlides = 0;
  let slides = [];

  // Initialize viewer with a default sample PPTX if clicked
  fileBtns.forEach((btn) => {
    btn.addEventListener("click", (e) => {
      fileBtns.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const fileName = btn.getAttribute("data-file");
      loadRemotePPTX(fileName);
    });
  });

  // Handle local file uploads
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) {
      fileBtns.forEach((b) => b.classList.remove("active"));
      loadLocalPPTX(file);
    }
  });

  // Navigation Event Listeners
  prevBtn.addEventListener("click", () => {
    if (currentSlideIndex > 0) {
      showSlide(currentSlideIndex - 1);
    }
  });

  nextBtn.addEventListener("click", () => {
    if (currentSlideIndex < totalSlides - 1) {
      showSlide(currentSlideIndex + 1);
    }
  });

  // Keyboard navigation
  document.addEventListener("keydown", (e) => {
    if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
      prevBtn.click();
    } else if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === " ") {
      nextBtn.click();
    }
  });

  // Load PPTX from URL/Path
  function loadRemotePPTX(url) {
    showLoading();
    fetch(url)
      .then((res) => {
        if (!res.ok) throw new Error("File not found");
        return res.arrayBuffer();
      })
      .then((buffer) => renderPPTX(buffer))
      .catch((err) => {
        showError("Could not load PPTX file. Ensure files exist on your server or upload a file.");
      });
  }

  // Load local uploaded PPTX file
  function loadLocalPPTX(file) {
    showLoading();
    const reader = new FileReader();
    reader.onload = (e) => renderPPTX(e.target.result);
    reader.readAsArrayBuffer(file);
  }

  // Render PPTX ArrayBuffer into HTML slides
  function renderPPTX(buffer) {
    slideStage.innerHTML = ""; // Clear current view

    $(`#${slideStage.id}`).pptx2html({
      pptx: buffer,
      slideMode: false,
      keyBoard: false,
      error: function () {
        showError("Failed to parse PPTX file format.");
      },
      success: function () {
        // Find all rendered slides
        slides = Array.from(slideStage.querySelectorAll("section"));
        totalSlides = slides.length;

        if (totalSlides === 0) {
          showError("No slides found in presentation.");
          return;
        }

        currentSlideIndex = 0;
        showSlide(0);
      },
    });
  }

  // Display specific slide by index
  function showSlide(index) {
    slides.forEach((slide, idx) => {
      if (idx === index) {
        slide.classList.add("active-slide");
      } else {
        slide.classList.remove("active-slide");
      }
    });

    currentSlideIndex = index;
    updateUI();
  }

  // Update Footer UI Controls
  function updateUI() {
    currentSlideEl.textContent = totalSlides > 0 ? currentSlideIndex + 1 : 0;
    totalSlidesEl.textContent = totalSlides;

    prevBtn.disabled = currentSlideIndex <= 0;
    nextBtn.disabled = currentSlideIndex >= totalSlides - 1 || totalSlides === 0;
  }

  function showLoading() {
    slideStage.innerHTML = `<div class="loading-spinner">Loading Presentation...</div>`;
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    currentSlideEl.textContent = 0;
    totalSlidesEl.textContent = 0;
  }

  function showError(msg) {
    slideStage.innerHTML = `<div style="color: #ff5555; text-align: center; padding: 20px;">${msg}</div>`;
  }
});
