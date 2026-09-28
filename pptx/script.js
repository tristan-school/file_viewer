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

  // Switch preset files
  fileBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      fileBtns.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const fileName = btn.getAttribute("data-file");
      loadRemotePPTX(fileName);
    });
  });

  // Handle local uploaded files
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) {
      // Remove active status from preset buttons
      fileBtns.forEach((b) => b.classList.remove("active"));
      loadLocalPPTX(file);
    }
  });

  // Navigation handlers
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

  function loadRemotePPTX(url) {
    showLoading("Loading preset presentation...");
    fetch(url)
      .then((res) => {
        if (!res.ok) throw new Error("File not found");
        return res.arrayBuffer();
      })
      .then((buffer) => renderPPTX(buffer))
      .catch(() => {
        showError("Remote file not found. Try uploading a local .pptx file.");
      });
  }

  function loadLocalPPTX(file) {
    showLoading(`Loading ${file.name}...`);
    const reader = new FileReader();

    reader.onload = (e) => {
      renderPPTX(e.target.result);
    };

    reader.onerror = () => {
      showError("Error reading the selected file.");
    };

    reader.readAsArrayBuffer(file);
  }

  function renderPPTX(buffer) {
    slideStage.innerHTML = "";

    // Parse presentation using PPTX2HTML library
    if (window.$ && $.fn.pptx2html) {
      $(`#${slideStage.id}`).pptx2html({
        pptx: buffer,
        slideMode: false,
        keyBoard: false,
        error: function () {
          showError("Failed to parse .pptx file structure.");
        },
        success: function () {
          slides = Array.from(slideStage.querySelectorAll("section"));
          totalSlides = slides.length;

          if (totalSlides === 0) {
            showError("No slides found in this file.");
            return;
          }

          showSlide(0);
        },
      });
    } else {
      showError("PPTX parsing library missing or failed to load.");
    }
  }

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

  function updateUI() {
    currentSlideEl.textContent = totalSlides > 0 ? currentSlideIndex + 1 : 0;
    totalSlidesEl.textContent = totalSlides;

    prevBtn.disabled = currentSlideIndex <= 0;
    nextBtn.disabled = currentSlideIndex >= totalSlides - 1 || totalSlides === 0;
  }

  function showLoading(msg) {
    slideStage.innerHTML = `<div class="loading-spinner">${msg}</div>`;
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    currentSlideEl.textContent = 0;
    totalSlidesEl.textContent = 0;
  }

  function showError(msg) {
    slideStage.innerHTML = `<div style="color: #ef4444; text-align: center; padding: 20px;">${msg}</div>`;
    prevBtn.disabled = true;
    nextBtn.disabled = true;
  }
});
