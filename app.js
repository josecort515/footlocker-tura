// ===== CONFIG =====
const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/1SWYxGtoO-ja3wNmKAlwlUvicoGGY4EqTGyw-2TuW3b4/export?format=csv';
// ===== PRODUCTS DATA =====
let products = [];

// ===== SIZE MAP BY CATEGORY/TYPE =====
const SIZE_MAP = {
  // Zapatos / calzado
  'zapatillas': ['38','39','40','41','42','43','44','45'],
  'zapatos':    ['38','39','40','41','42','43','44','45'],
  'botas':      ['38','39','40','41','42','43','44','45'],
  'sandalias':  ['38','39','40','41','42','43','44','45'],
  // Camisas
  'camisas':    ['S','M','L','XL','XXL'],
  'camisa':     ['S','M','L','XL','XXL'],
  // Camisetas
  'camisetas':  ['S','M','L','XL','XXL'],
  'camiseta':   ['S','M','L','XL','XXL'],
  // Pantalones
  'pantalones': ['30','32','34','36','38','40'],
  'pantalon':   ['30','32','34','36','38','40'],
  'jeans':      ['30','32','34','36','38','40'],
  // Sudaderas
  'sudaderas':  ['S','M','L','XL','XXL'],
  'sudadera':   ['S','M','L','XL','XXL'],
};

const DEFAULT_SIZES = ['S','M','L','XL','XXL'];

function getSizesForProduct(product) {
  const cat = (product.category || '').toLowerCase();
  const typ = (product.type || '').toLowerCase();
  return SIZE_MAP[cat] || SIZE_MAP[typ] || DEFAULT_SIZES;
}


// ===== CSV PARSER =====
function parseCSV(text) {
  const lines = text.trim().split('\n');
  if (lines.length < 2) return [];

  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());

  return lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.trim());
    const obj = {};
    headers.forEach((h, i) => { obj[h] = values[i] || ''; });

    const colores = (obj.colores || '').split('|').filter(c => c.trim());
    const fotosRaw = (obj.fotos || obj.img || '').split(';').filter(f => f.trim());
    const fotos = fotosRaw.map(group => group.split('|').filter(u => u.trim()));

    // Fallback: si no hay colores pero hay fotos, crear un solo color
    if (colores.length === 0 && fotos.length > 0) {
      colores.push('Único');
    }

    return {
      id: parseInt(obj.id) || 0,
      name: obj.nombre || obj.name || '',
      category: obj.categoria || obj.category || '',
      type: obj.tipo || obj.type || '',
      price: parseFloat(obj.precio || obj.price) || 0,
      oldPrice: parseFloat(obj.precio_old || obj.old_price) || null,
      badge: obj.badge || null,
      colores,
      fotos,
    };
  }).filter(p => p.id && p.name);
}

// ===== PARSE EXCEL (SheetJS) =====
function parseExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

  return rows.map(row => {
    const colores = String(row.colores || row.Colors || '').split('|').filter(c => c.trim());
    const fotosRaw = String(row.fotos || row.img || row.Fotos || row.Images || '').split(';').filter(f => f.trim());
    const fotos = fotosRaw.map(group => group.split('|').filter(u => u.trim()));

    if (colores.length === 0 && fotos.length > 0) colores.push('Único');

    return {
      id: parseInt(row.id) || 0,
      name: row.nombre || row.name || row.Name || row.Nombre || '',
      category: row.categoria || row.category || row.Category || row.Categoria || '',
      type: row.tipo || row.type || row.Type || row.Tipo || '',
      price: parseFloat(row.precio || row.price || row.Price || row.Precio) || 0,
      oldPrice: parseFloat(row.precio_old || row.old_price || row.OldPrice) || null,
      badge: row.badge || row.Badge || null,
      colores,
      fotos,
    };
  }).filter(p => p.id && p.name);
}

// ===== FETCH PRODUCTS =====
async function fetchProducts() {
  // 1. Intentar Excel local
  try {
    const res = await fetch('productos.xlsx');
    if (res.ok) {
      const buffer = await res.arrayBuffer();
      const parsed = parseExcel(buffer);
      if (parsed.length > 0) {
        console.log(`Productos cargados desde productos.xlsx: ${parsed.length}`);
        products = parsed;
        return;
      }
    }
  } catch (e) {
    console.warn('No se pudo cargar productos.xlsx:', e.message);
  }

  // 2. Intentar JSON local
  try {
    const res = await fetch('productos.json');
    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data) && data.length > 0) {
        console.log(`Productos cargados desde productos.json: ${data.length}`);
        products = data;
        return;
      }
    }
  } catch (e) {
    console.warn('No se pudo cargar productos.json:', e.message);
  }

  // 3. Intentar Google Sheets
  if (SHEET_CSV_URL) {
    try {
      console.log('Intentando cargar desde Google Sheets...');
      const res = await fetch(SHEET_CSV_URL);
      if (res.ok) {
        const text = await res.text();
        const parsed = parseCSV(text);
        if (parsed.length > 0) {
          console.log(`Productos cargados desde Sheet: ${parsed.length}`);
          products = parsed;
          return;
        }
      }
    } catch (e) {
      console.warn('No se pudo cargar desde Sheet:', e.message);
    }
  }

  // 4. Fallback: datos hardcodeados
  console.log('Usando datos de ejemplo');
  products = FALLBACK_PRODUCTS;
}

// ===== CART =====
function getCart() {
  return JSON.parse(localStorage.getItem('nike_cart') || '[]');
}

function saveCart(cart) {
  localStorage.setItem('nike_cart', JSON.stringify(cart));
  updateCartCount();
}

function addToCart(productId, size) {
  const cart = getCart();
  const product = products.find(p => p.id === productId);
  if (!product) return;

  const existing = cart.find(item => item.id === productId && item.size === size);
  if (existing) {
    existing.qty += 1;
  } else {
    cart.push({ id: productId, size, qty: 1, ...product });
  }

  saveCart(cart);
  showToast('Añadido al carrito');
}

function removeFromCart(productId, size) {
  let cart = getCart();
  cart = cart.filter(item => !(item.id === productId && item.size === size));
  saveCart(cart);
}

function updateQty(productId, size, delta) {
  const cart = getCart();
  const item = cart.find(i => i.id === productId && i.size === size);
  if (!item) return;

  item.qty += delta;
  if (item.qty <= 0) {
    removeFromCart(productId, size);
  } else {
    saveCart(cart);
  }
}

function getCartTotal() {
  return getCart().reduce((sum, item) => sum + item.price * item.qty, 0);
}

function getCartCount() {
  return getCart().reduce((sum, item) => sum + item.qty, 0);
}

function updateCartCount() {
  const el = document.querySelector('.header__cart-count');
  if (el) {
    const count = getCartCount();
    el.textContent = count;
    el.style.display = count > 0 ? 'flex' : 'none';
  }
}

// ===== NAVIGATION DRAWER =====
function initNav() {
  const hamburger = document.querySelector('.hamburger');
  const drawer = document.querySelector('.nav-drawer');
  const overlay = document.querySelector('.nav-overlay');
  const closeBtn = document.querySelector('.nav-drawer__close');

  if (!hamburger || !drawer) return;

  function openNav() {
    drawer.classList.add('active');
    overlay.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeNav() {
    drawer.classList.remove('active');
    overlay.classList.remove('active');
    document.body.style.overflow = '';
  }

  hamburger.addEventListener('click', openNav);
  overlay.addEventListener('click', closeNav);
  closeBtn.addEventListener('click', closeNav);
}

// ===== FILTER DRAWER =====
function initFilters() {
  const toggleBtn = document.querySelector('.filter-toggle__btn');
  const drawer = document.querySelector('.filter-drawer');
  const overlay = document.querySelector('.filter-overlay');
  const closeBtn = document.querySelector('.filter-drawer__close');

  if (!toggleBtn || !drawer) return;

  function openFilters() {
    drawer.classList.add('active');
    overlay.classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeFilters() {
    drawer.classList.remove('active');
    overlay.classList.remove('active');
    document.body.style.overflow = '';
  }

  toggleBtn.addEventListener('click', openFilters);
  overlay.addEventListener('click', closeFilters);
  closeBtn.addEventListener('click', closeFilters);
}

// ===== TABS =====
function initTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const wasOpen = tab.classList.contains('open');
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('open'));
      if (!wasOpen) tab.classList.add('open');
    });
  });
}

// ===== SIZE SELECTOR =====
function initSizeSelector() {
  const sizes = document.querySelectorAll('.size-btn:not(:disabled)');
  const addBtn = document.querySelector('.add-to-cart');

  sizes.forEach(btn => {
    btn.addEventListener('click', () => {
      sizes.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      if (addBtn) {
        addBtn.disabled = false;
        addBtn.textContent = 'Añadir al carrito';
      }
    });
  });
}

// ===== TOAST =====
function showToast(message) {
  let toast = document.querySelector('.toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.classList.add('show');
  setTimeout(() => toast.classList.remove('show'), 2500);
}

// ===== GALLERY =====
let currentGalleryIndex = 0;

function initGallery(total) {
  currentGalleryIndex = 0;
  const thumbs = document.querySelectorAll('.gallery__thumb');
  const mainImg = document.querySelector('.gallery__main-img');
  const counter = document.querySelector('.gallery__counter');
  const prevBtn = document.querySelector('.gallery__nav--prev');
  const nextBtn = document.querySelector('.gallery__nav--next');

  function updateGallery(index) {
    currentGalleryIndex = index;
    if (mainImg) mainImg.src = thumbs[index]?.dataset?.src || '';
    thumbs.forEach((t, i) => t.classList.toggle('active', i === index));
    if (counter) counter.textContent = `${index + 1} / ${total}`;
    if (prevBtn) prevBtn.style.display = index === 0 ? 'none' : 'flex';
    if (nextBtn) nextBtn.style.display = index === total - 1 ? 'none' : 'flex';
  }

  thumbs.forEach((thumb, i) => {
    thumb.addEventListener('click', () => updateGallery(i));
  });

  if (prevBtn) prevBtn.addEventListener('click', () => {
    if (currentGalleryIndex > 0) updateGallery(currentGalleryIndex - 1);
  });

  if (nextBtn) nextBtn.addEventListener('click', () => {
    if (currentGalleryIndex < total - 1) updateGallery(currentGalleryIndex + 1);
  });

  // Touch swipe
  let startX = 0;
  const gallery = document.querySelector('.gallery');
  if (gallery) {
    gallery.addEventListener('touchstart', e => { startX = e.touches[0].clientX; }, { passive: true });
    gallery.addEventListener('touchend', e => {
      const diff = startX - e.changedTouches[0].clientX;
      if (Math.abs(diff) > 50) {
        if (diff > 0 && currentGalleryIndex < total - 1) updateGallery(currentGalleryIndex + 1);
        if (diff < 0 && currentGalleryIndex > 0) updateGallery(currentGalleryIndex - 1);
      }
    }, { passive: true });
  }

  updateGallery(0);
}

// ===== RENDER PRODUCT CARDS =====
function renderProductGrid(containerId, items) {
  const container = document.getElementById(containerId);
  if (!container) return;

  container.innerHTML = items.map(p => `
    <a href="product.html?id=${p.id}" class="product-card">
      <div class="product-card__img-wrapper">
        <img src="${p.fotos[0][0]}" alt="${p.name}" loading="lazy">
        ${p.badge ? `<span class="product-card__badge">${p.badge}</span>` : ''}
      </div>
      <div class="product-card__name">${p.name}</div>
      <div class="product-card__category">${p.type}</div>
      <div class="product-card__price">
        ${p.price.toFixed(2)} €
        ${p.oldPrice ? `<span class="product-card__price--old">${p.oldPrice.toFixed(2)} €</span>` : ''}
      </div>
    </a>
  `).join('');
}

// ===== RENDER CHECKOUT =====
function renderCheckout() {
  const container = document.getElementById('checkout-items');
  const summaryEl = document.getElementById('checkout-summary');
  if (!container) return;

  const cart = getCart();

  if (cart.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state__icon">🛒</div>
        <div class="empty-state__title">Tu carrito está vacío</div>
        <div class="empty-state__text">Añade algunos productos para continuar</div>
        <a href="index.html" class="empty-state__btn">Ver productos</a>
      </div>
    `;
    if (summaryEl) summaryEl.style.display = 'none';
    return;
  }

  container.innerHTML = cart.map(item => `
    <div class="checkout-item">
      <div class="checkout-item__img">
        <img src="${item.fotos[0][0]}" alt="${item.name}">
      </div>
      <div class="checkout-item__details">
        <div class="checkout-item__name">${item.name}</div>
        <div class="checkout-item__meta">Talla: ${item.size}</div>
        <div class="checkout-item__qty">
          <button class="qty-btn" onclick="changeQty(${item.id}, '${item.size}', -1)">−</button>
          <span class="qty-value">${item.qty}</span>
          <button class="qty-btn" onclick="changeQty(${item.id}, '${item.size}', 1)">+</button>
        </div>
      </div>
      <div class="checkout-item__price">${(item.price * item.qty).toFixed(2)} €</div>
    </div>
  `).join('');

  if (summaryEl) {
    const subtotal = getCartTotal();
    const total = subtotal;

    summaryEl.innerHTML = `
      <div class="summary-row">
        <span>Subtotal</span>
        <span>${subtotal.toFixed(2)} €</span>
      </div>
      <div class="summary-row">
        <span>Envío</span>
        <span>Gratis</span>
      </div>
      <div class="summary-row summary-row--total">
        <span>Total</span>
        <span>${total.toFixed(2)} €</span>
      </div>
    `;
    summaryEl.style.display = 'block';
  }
}

function changeQty(id, size, delta) {
  updateQty(id, size, delta);
  renderCheckout();
}

// ===== WHATSAPP CHECKOUT =====
const WHATSAPP_NUMBER = '34601904823';

function sendToWhatsApp() {
  const cart = getCart();
  if (cart.length === 0) {
    showToast('Tu carrito está vacío');
    return;
  }

  const email = document.getElementById('email').value.trim();
  const name = document.getElementById('name').value.trim();
  const address = document.getElementById('address').value.trim();
  const city = document.getElementById('city').value.trim();
  const zip = document.getElementById('zip').value.trim();
  const province = document.getElementById('province').value;

  if (!name || !address || !city || !zip || !province || !email) {
    showToast('Por favor, completa todos los campos de envío');
    return;
  }

  let msg = '*Nuevo pedido - Nike Store*\n\n';
  msg += '*Productos:*\n';
  cart.forEach(item => {
    msg += `• ${item.name} (Ref: ${item.id}, Color: ${item.colores[0]}, Talla ${item.size}) x${item.qty} — ${(item.price * item.qty).toFixed(2)} €\n`;
  });

  msg += `\n*Total: ${getCartTotal().toFixed(2)} €*\n`;
  msg += '\n*Datos de envío:*\n';
  msg += `• Nombre: ${name}\n`;
  msg += `• Email: ${email}\n`;
  msg += `• Dirección: ${address}\n`;
  msg += `• Ciudad: ${city}\n`;
  msg += `• CP: ${zip}\n`;
  msg += `• Provincia: ${province}\n`;

  const url = `https://wa.me/${WHATSAPP_NUMBER}?text=${encodeURIComponent(msg)}`;
  window.open(url, '_blank');

  localStorage.removeItem('nike_cart');
  updateCartCount();
  renderCheckout();
}

// ===== RENDER PRODUCT DETAIL =====
function renderProductDetail() {
  const container = document.getElementById('product-detail');
  if (!container) return;

  const params = new URLSearchParams(window.location.search);
  const id = parseInt(params.get('id'));
  const product = products.find(p => p.id === id);

  if (!product) {
    container.innerHTML = `
      <div class="empty-state">
        <div class="empty-state__icon">😕</div>
        <div class="empty-state__title">Producto no encontrado</div>
        <a href="index.html" class="empty-state__btn">Volver al inicio</a>
      </div>
    `;
    return;
  }

  document.title = `${product.name} | NIKE`;

  const colores = product.colores;
  const fotos = product.fotos;
  const firstColorFotos = fotos[0] || [];
  const hasMultiplePhotos = firstColorFotos.length > 1;
  const hasMultipleColors = colores.length > 1;

  container.innerHTML = `
    <div class="gallery">
      <div class="gallery__main">
        <img class="gallery__main-img" src="${firstColorFotos[0]}" alt="${product.name}">
        ${hasMultiplePhotos ? `
          <button class="gallery__nav gallery__nav--prev" style="display:none">‹</button>
          <button class="gallery__nav gallery__nav--next">›</button>
          <span class="gallery__counter">1 / ${firstColorFotos.length}</span>
        ` : ''}
      </div>
      ${hasMultiplePhotos ? `
        <div class="gallery__thumbs">
          ${firstColorFotos.map((url, i) => `
            <button class="gallery__thumb ${i === 0 ? 'active' : ''}" data-src="${url}">
              <img src="${url}" alt="Foto ${i + 1}">
            </button>
          `).join('')}
        </div>
      ` : ''}
    </div>
    <div class="product-detail__info">
      <div class="product-detail__brand">Nike</div>
      <div class="product-detail__name">${product.name}</div>
      <div class="product-detail__type">${product.type}</div>
      <div class="product-detail__price">
        ${product.price.toFixed(2)} €
        ${product.oldPrice ? `<span class="product-card__price--old">${product.oldPrice.toFixed(2)} €</span>` : ''}
      </div>
      ${hasMultipleColors ? `
        <div class="color-selector">
          <div class="color-selector__label">Color: <span id="color-name">${colores[0]}</span></div>
          <div class="color-selector__options">
            ${colores.map((c, i) => `
              <button class="color-btn ${i === 0 ? 'active' : ''}" data-color-index="${i}" title="${c}">
                <span class="color-btn__name">${c}</span>
              </button>
            `).join('')}
          </div>
        </div>
      ` : `
        <div class="product-detail__color">Color: <span>${colores[0]}</span></div>
      `}

      <div class="size-selector">
        <div class="size-selector__label">
          Selecciona tu talla
          <a href="#">Guía de tallas</a>
        </div>
        <div class="size-grid">
          ${getSizesForProduct(product).map(s => `
            <button class="size-btn" data-size="${s}">${s}</button>
          `).join('')}
        </div>
      </div>

      <div class="add-to-cart-wrapper">
        <button class="add-to-cart" disabled onclick="handleAddToCart(${product.id})">Selecciona una talla</button>
      </div>
      <button class="wishlist-btn">
        <span>♡</span> Favoritos
      </button>

      <div class="tabs">
        <button class="tab">
          Descripción
          <span class="tab__icon">+</span>
        </button>
        <div class="tab__content">
          <div class="tab__content-inner">
            La ${product.name} combina un diseño icónico con comodidad moderna. 
            Perfecta para el uso diario, ofrece un ajuste cómodo y un estilo que nunca pasa de moda. 
            Fabricada con materiales de alta calidad para mayor durabilidad.
          </div>
        </div>
        <button class="tab">
          Detalles del producto
          <span class="tab__icon">+</span>
        </button>
        <div class="tab__content">
          <div class="tab__content-inner">
            <ul>
              <li>• Color seleccionado: <span id="detail-color">${colores[0]}</span></li>
              <li>• Estilo: ${product.type}</li>
              <li>• Envío y devoluciones gratuitas</li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  `;

  initSizeSelector();
  initTabs();
  if (hasMultiplePhotos) initGallery(firstColorFotos.length);
  if (hasMultipleColors) initColorSelector(product);
}

// ===== COLOR SELECTOR =====
function initColorSelector(product) {
  const colorBtns = document.querySelectorAll('.color-btn');
  const colorNameEl = document.getElementById('color-name');
  const detailColorEl = document.getElementById('detail-color');

  colorBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const index = parseInt(btn.dataset.colorIndex);

      // Update active state
      colorBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      // Update color name
      if (colorNameEl) colorNameEl.textContent = product.colores[index];
      if (detailColorEl) detailColorEl.textContent = product.colores[index];

      // Update gallery with new color's photos
      const newFotos = product.fotos[index] || [];
      const gallery = document.querySelector('.gallery');
      if (!gallery || newFotos.length === 0) return;

      const mainImg = gallery.querySelector('.gallery__main-img');
      const thumbsContainer = gallery.querySelector('.gallery__thumbs');
      const counter = gallery.querySelector('.gallery__counter');
      const prevBtn = gallery.querySelector('.gallery__nav--prev');
      const nextBtn = gallery.querySelector('.gallery__nav--next');

      if (mainImg) mainImg.src = newFotos[0];

      if (newFotos.length > 1) {
        // Rebuild thumbnails
        if (thumbsContainer) {
          thumbsContainer.innerHTML = newFotos.map((url, i) => `
            <button class="gallery__thumb ${i === 0 ? 'active' : ''}" data-src="${url}">
              <img src="${url}" alt="Foto ${i + 1}">
            </button>
          `).join('');
        }
        if (counter) counter.textContent = `1 / ${newFotos.length}`;
        if (prevBtn) prevBtn.style.display = 'none';
        if (nextBtn) nextBtn.style.display = 'flex';

        // Re-init gallery
        initGallery(newFotos.length);
      } else {
        // Single photo - hide gallery controls
        if (thumbsContainer) thumbsContainer.innerHTML = '';
        if (counter) counter.textContent = '';
        if (prevBtn) prevBtn.style.display = 'none';
        if (nextBtn) nextBtn.style.display = 'none';
      }
    });
  });
}

function handleAddToCart(productId) {
  const activeSize = document.querySelector('.size-btn.active');
  if (!activeSize) return;
  addToCart(productId, activeSize.dataset.size);
}

// ===== SEARCH & FILTER =====
function initSearch() {
  const input = document.getElementById('search-input');
  if (!input) return;

  const params = new URLSearchParams(window.location.search);
  const q = params.get('q');
  if (q) input.value = q;

  let debounceTimer;
  input.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      applyFilters();
    }, 300);
  });

  if (q) applyFilters();
}

function applyFilters() {
  const searchVal = (document.getElementById('search-input')?.value || '').toLowerCase();
  const activeChips = document.querySelectorAll('.filter-chip.active');
  const filters = Array.from(activeChips).map(c => c.dataset.value);

  let filtered = products;

  if (searchVal) {
    filtered = filtered.filter(p =>
      p.name.toLowerCase().includes(searchVal) ||
      p.type.toLowerCase().includes(searchVal) ||
      p.category.toLowerCase().includes(searchVal)
    );
  }

  if (filters.length > 0) {
    filtered = filtered.filter(p =>
      filters.includes(p.category) || filters.includes(p.type)
    );
  }

  const countEl = document.querySelector('.results-count');
  if (countEl) countEl.textContent = `${filtered.length} resultados`;

  renderProductGrid('product-grid', filtered);
}

function initFilterChips() {
  document.querySelectorAll('.filter-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      chip.classList.toggle('active');
      applyFilters();
    });
  });
}

// ===== STICKY SEARCH (Index) =====
function initStickySearch() {
  const stickySearch = document.getElementById('sticky-search');
  if (!stickySearch) return;

  const hero = document.querySelector('.hero');
  if (!hero) return;

  const threshold = hero.offsetHeight - 100;

  window.addEventListener('scroll', () => {
    if (window.scrollY > threshold) {
      stickySearch.classList.add('visible');
    } else {
      stickySearch.classList.remove('visible');
    }
  }, { passive: true });
}

function handleStickySearch(e) {
  e.preventDefault();
  const input = document.getElementById('sticky-search-input');
  const query = (input?.value || '').trim();
  if (query) {
    window.location.href = `search.html?q=${encodeURIComponent(query)}`;
  } else {
    window.location.href = 'search.html';
  }
  return false;
}

// ===== INIT =====
document.addEventListener('DOMContentLoaded', async () => {
  await fetchProducts();
  updateCartCount();
  initNav();
  initStickySearch();

  // Page-specific
  if (document.getElementById('product-grid')) {
    renderProductGrid('product-grid', products);
    initFilters();
    initFilterChips();
    initSearch();
  }

  if (document.getElementById('product-detail')) {
    renderProductDetail();
  }

  if (document.getElementById('checkout-items')) {
    renderCheckout();
  }
});
