document.addEventListener("DOMContentLoaded", () => {
  const slideStage = document.getElementById("slide-stage");
  
  // Guard script if executed on index page without PPTX viewer stage
  if (!slideStage) return;

  const prevBtn = document.getElementById("prev-btn");
  const nextBtn = document.getElementById("next-btn");
  const currentSlideEl = document.getElementById("current-slide");
  const totalSlidesEl = document.getElementById("total-slides");
  const fileBtns = document.querySelectorAll(".file-btn");
  const fileInput = document.getElementById("file-input");

  let currentSlideIndex = 0;
  let totalSlides = 0;
  let slides = [];

  // Switch between preset presentations
  fileBtns.forEach((btn) => {
    btn.addEventListener("click", () => {
      fileBtns.forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      const fileName = btn.getAttribute("data-file");
      loadRemotePPTX(fileName);
    });
  });

  // Upload local presentation
  fileInput.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) {
      fileBtns.forEach((b) => b.classList.remove("active"));
      loadLocalPPTX(file);
    }
  });

  // Footer navigation controls
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

  // Keyboard Navigation
  document.addEventListener("keydown", (e) => {
    if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
      prevBtn.click();
    } else if (e.key === "ArrowRight" || e.key === "ArrowDown" || e.key === " ") {
      nextBtn.click();
    }
  });

  function loadRemotePPTX(url) {
    showLoading();
    fetch(url)
      .then((res) => {
        if (!res.ok) throw new Error("File not found");
        return res.arrayBuffer();
      })
      .then((buffer) => renderPPTX(buffer))
      .catch(() => {
        showError("Could not load PPTX file from server. Upload a file via the button above.");
      });
  }

  function loadLocalPPTX(file) {
    showLoading();
    const reader = new FileReader();
    reader.onload = (e) => renderPPTX(e.target.result);
    reader.readAsArrayBuffer(file);
  }

  function renderPPTX(buffer) {
    slideStage.innerHTML = "";

    if (window.$ && $.fn.pptx2html) {
      $(`#${slideStage.id}`).pptx2html({
        pptx: buffer,
        slideMode: false,
        keyBoard: false,
        error: function () {
          showError("Failed to parse PowerPoint presentation format.");
        },
        success: function () {
          slides = Array.from(slideStage.querySelectorAll("section"));
          totalSlides = slides.length;

          if (totalSlides === 0) {
            showError("No slide slides found in presentation.");
            return;
          }

          currentSlideIndex = 0;
          showSlide(0);
        },
      });
    } else {
      // Demo fallback renderer mode if CDN libraries are unavailable offline
      renderDemoSlides();
    }
  }

  function renderDemoSlides() {
    slideStage.innerHTML = `
      <section class="active-slide" style="display:flex; justify-content:center; align-items:center; flex-direction:column; padding:40px; text-align:center;">
        <h2 style="font-size:2rem; margin-bottom:12px; color:#0284c7;">Welcome Slide</h2>
        <p style="color:#475569;">This is a rendered demonstration slide.</p>
      </section>
      <section style="display:flex; justify-content:center; align-items:center; flex-direction:column; padding:40px; text-align:center;">
        <h2 style="font-size:2rem; margin-bottom:12px; color:#0284c7;">Slide 2: Features</h2>
        <p style="color:#475569;">Full navigation controls, file upload support, and presentation switching.</p>
      </section>
      <section style="display:flex; justify-content:center; align-items:center; flex-direction:column; padding:40px; text-align:center;">
        <h2 style="font-size:2rem; margin-bottom:12px; color:#0284c7;">Slide 3: Summary</h2>
        <p style="color:#475569;">You have reached the end of the presentation.</p>
      </section>
    `;

    slides = Array.from(slideStage.querySelectorAll("section"));
    totalSlides = slides.length;
    showSlide(0);
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

  function showLoading() {
    slideStage.innerHTML = `<div class="loading-spinner">Loading Presentation...</div>`;
    prevBtn.disabled = true;
    nextBtn.disabled = true;
    currentSlideEl.textContent = 0;
    totalSlidesEl.textContent = 0;
  }

  function showError(msg) {
    slideStage.innerHTML = `<div style="color: #ef4444; text-align: center; padding: 20px;">${msg}</div>`;
  }

  // Load initial demo/placeholder presentation
  renderDemoSlides();
});
