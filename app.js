// Core PPTX Parser & Slide Renderer Engine
let slidesData = [];
let currentSlideIndex = 0;

const uploadInput = document.getElementById('pptx-upload');
const demoBtn = document.getElementById('demo-btn');
const emptyState = document.getElementById('empty-state');
const workspace = document.getElementById('slide-workspace');
const bottomNav = document.getElementById('bottom-nav');
const slideCanvas = document.getElementById('slide-canvas');
const slideViewport = document.getElementById('slide-viewport');
const fileNameLabel = document.getElementById('file-name-label');
const slideCounterBadge = document.getElementById('slide-counter-badge');
const thumbnailStrip = document.getElementById('thumbnail-strip');
const prevBtn = document.getElementById('prev-btn');
const nextBtn = document.getElementById('next-btn');
const fullscreenBtn = document.getElementById('fullscreen-btn');

// 1. Handle Real .PPTX File Upload
uploadInput.addEventListener('change', async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  fileNameLabel.textContent = file.name;
  await parsePptxFile(file);
});

// 2. Load Built-in Demo Presentation
demoBtn.addEventListener('click', () => {
  fileNameLabel.textContent = 'Demo_Engineering_Pitch.pptx';
  slidesData = [
    {
      title: 'Welcome to Pocket PPTX Viewer 📊',
      bullets: [
        '100% Client-Side PowerPoint (.pptx) Unpacker & Renderer',
        'Built with HTML5, CSS3, JavaScript & JSZip',
        'Packaged into a native Android APK via Capacitor & GitHub Actions'
      ],
      images: []
    },
    {
      title: 'How This App Opens .PPTX Files ⚙️',
      bullets: [
        'Every .pptx file is secretly a ZIP archive containing XML files',
        'JSZip extracts ppt/slides/slide1.xml, slide2.xml, etc. in milliseconds',
        'Our XML parser reads all <a:t> text nodes and embedded media images',
        'Zero server uploads — works 100% offline on your Android phone!'
      ],
      images: []
    },
    {
      title: 'Touch Controls & Gesture Support 👆',
      bullets: [
        'Swipe Left or Right on the slide stage to flip through slides',
        'Tap any slide thumbnail in the bottom filmstrip to jump directly',
        'Tap Present for distraction-free Fullscreen mode'
      ],
      images: []
    }
  ];
  currentSlideIndex = 0;
  launchViewerUI();
});

// 3. Unpack and Parse Real .PPTX Archive
async function parsePptxFile(file) {
  try {
    const zip = await JSZip.loadAsync(file);
    const slideFiles = Object.keys(zip.files)
      .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml/)[1], 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml/)[1], 10);
        return numA - numB;
      });

    if (slideFiles.length === 0) {
      alert('No valid slides found inside this .pptx file.');
      return;
    }

    const parsedSlides = [];

    for (let i = 0; i < slideFiles.length; i++) {
      const slidePath = slideFiles[i];
      const slideNum = slidePath.match(/slide(\d+)\.xml/)[1];
      const xmlText = await zip.file(slidePath).async('string');

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

      // Extract Paragraphs (<a:p>) and Text Runs (<a:t>)
      const paragraphs = Array.from(xmlDoc.getElementsByTagName('a:p'));
      const lines = [];

      paragraphs.forEach((p) => {
        const textNodes = Array.from(p.getElementsByTagName('a:t'));
        const lineText = textNodes.map((node) => node.textContent).join('').trim();
        if (lineText) lines.push(lineText);
      });

      // Check slide relationships for embedded images
      const relsPath = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
      const slideImages = [];
      if (zip.file(relsPath)) {
        const relsXml = await zip.file(relsPath).async('string');
        const relsDoc = parser.parseFromString(relsXml, 'application/xml');
        const relNodes = Array.from(relsDoc.getElementsByTagName('Relationship'));

        for (const rel of relNodes) {
          const target = rel.getAttribute('Target') || '';
          if (target.includes('../media/')) {
            const mediaPath = 'ppt/' + target.replace('../', '');
            const mediaFile = zip.file(mediaPath);
            if (mediaFile) {
              const blob = await mediaFile.async('blob');
              const imgUrl = URL.createObjectURL(blob);
              slideImages.push(imgUrl);
            }
          }
        }
      }

      parsedSlides.push({
        title: lines[0] || `Slide ${i + 1}`,
        bullets: lines.slice(1),
        images: slideImages
      });
    }

    slidesData = parsedSlides;
    currentSlideIndex = 0;
    launchViewerUI();
  } catch (err) {
    console.error(err);
    alert('Could not read PPTX file. Make sure it is a valid .pptx presentation.');
  }
}

// 4. Render Slide and Thumbnails
function launchViewerUI() {
  emptyState.classList.add('hidden');
  workspace.classList.remove('hidden');
  bottomNav.classList.remove('hidden');
  renderThumbnails();
  renderCurrentSlide();
}

function renderCurrentSlide() {
  const slide = slidesData[currentSlideIndex];
  slideCounterBadge.textContent = `Slide ${currentSlideIndex + 1} / ${slidesData.length}`;

  let html = `<h2>${escapeHtml(slide.title)}</h2>`;
  if (slide.bullets.length > 0) {
    html += '<ul>' + slide.bullets.map((b) => `<li>${escapeHtml(b)}</li>`).join('') + '</ul>';
  } else if (slide.images.length === 0) {
    html += '<p style="color:#71717a;">(Visual or blank slide)</p>';
  }

  if (slide.images.length > 0) {
    html += '<div class="slide-images-grid">' +
      slide.images.map((src) => `<img src="${src}" alt="Slide Visual" />`).join('') +
      '</div>';
  }

  slideCanvas.innerHTML = html;

  prevBtn.disabled = currentSlideIndex === 0;
  nextBtn.disabled = currentSlideIndex === slidesData.length - 1;

  // Highlight active thumbnail
  document.querySelectorAll('.thumb-item').forEach((el, idx) => {
    el.classList.toggle('active', idx === currentSlideIndex);
  });
}

function renderThumbnails() {
  thumbnailStrip.innerHTML = '';
  slidesData.forEach((slide, idx) => {
    const thumb = document.createElement('div');
    thumb.className = 'thumb-item' + (idx === currentSlideIndex ? ' active' : '');
    thumb.innerHTML = `<strong>#${idx + 1}</strong><span>${escapeHtml(slide.title.slice(0, 18))}</span>`;
    thumb.addEventListener('click', () => {
      currentSlideIndex = idx;
      renderCurrentSlide();
    });
    thumbnailStrip.appendChild(thumb);
  });
}

// 5. Navigation & Touch Swipe Gestures
prevBtn.addEventListener('click', () => {
  if (currentSlideIndex > 0) {
    currentSlideIndex--;
    renderCurrentSlide();
  }
});

nextBtn.addEventListener('click', () => {
  if (currentSlideIndex < slidesData.length - 1) {
    currentSlideIndex++;
    renderCurrentSlide();
  }
});

fullscreenBtn.addEventListener('click', () => {
  if (!document.fullscreenElement) {
    slideViewport.requestFullscreen().catch(() => {});
  } else {
    document.exitFullscreen();
  }
});

let touchStartX = 0;
slideViewport.addEventListener('touchstart', (e) => {
  touchStartX = e.changedTouches[0].screenX;
});

slideViewport.addEventListener('touchend', (e) => {
  const diff = e.changedTouches[0].screenX - touchStartX;
  if (Math.abs(diff) > 50) {
    if (diff < 0 && currentSlideIndex < slidesData.length - 1) {
      currentSlideIndex++;
      renderCurrentSlide();
    } else if (diff > 0 && currentSlideIndex > 0) {
      currentSlideIndex--;
      renderCurrentSlide();
    }
  }
});

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[char]));
}
