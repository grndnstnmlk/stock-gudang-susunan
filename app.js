// ==============================================================================
// 3D TOBACCO WAREHOUSE VISUALIZER (STOCK GUDANG SUSUNAN BAL)
// Built with Three.js & Modern Web Standards
// ==============================================================================

let scene, camera, renderer, controls;
let balesData = []; // flattened bale entities
let baleMeshes = []; // Three.js meshes
let blockGroups = {}; // grouped by block id
let activeFilter = {
  block: 'all',
  grade: 'all',
  layers: { 1: true, 2: true, 3: true, 4: true },
  searchQuery: '',
  colorMode: 'grade', // 'grade' | 'weight' | 'layer' | 'realistic'
  explodeFactor: 0
};

let hoveredBale = null;
let selectedBale = null;
let raycaster = new THREE.Raycaster();
let mouse = new THREE.Vector2();

// Layout configuration
const BALE_WIDTH = 1.35;  // Lebar bal (Barat - Timur, X)
const BALE_HEIGHT = 0.85; // Tinggi per Tingkat (Y)
const BALE_DEPTH = 1.45;  // Panjang bal (Utara - Selatan per Saf, Z)
const BALE_GAP_X = 0.12;  // Celah antar blok
const BALE_GAP_Z = 0.08;  // Celah antar saf
const BALE_GAP_Y = 0.04;  // Celah antar tingkat

// Color maps
const GRADE_COLORS = {
  'SAM': 0xec4899, // Pink
  '55': 0x3b82f6,  // Blue
  '50': 0x06b6d4,  // Cyan
  '45': 0x10b981,  // Emerald
  '40': 0x84cc16,  // Lime
  '35': 0xeab308,  // Yellow
  '30': 0xf97316,  // Orange
  '25': 0xef4444,  // Red
  '20': 0xa855f7,  // Purple
  '15': 0x6366f1,  // Indigo
  'DEFAULT': 0x64748b // Slate
};

const LAYER_COLORS = {
  1: 0x3b82f6, // T1 Dasar - Blue
  2: 0x10b981, // T2 - Emerald
  3: 0xf59e0b, // T3 - Amber
  4: 0xec4899, // T4 Atas - Rose
  5: 0x8b5cf6,
  6: 0x06b6d4,
  7: 0x64748b
};

// Initialize App
window.addEventListener('DOMContentLoaded', () => {
  initThree();
  loadInitialData();
  setupUIEventListeners();
  animate();
});

// ==============================================================================
// 1. THREE.JS SCENE SETUP & WAREHOUSE RECTANGULAR BUILDING
// ==============================================================================
function initThree() {
  const container = document.getElementById('canvas-container');
  
  scene = new THREE.Scene();
  scene.background = new THREE.Color(0x0b0f19);
  scene.fog = new THREE.FogExp2(0x0b0f19, 0.002);

  camera = new THREE.PerspectiveCamera(45, window.innerWidth / window.innerHeight, 0.5, 1000);
  camera.position.set(0, 20, -18);

  renderer = new THREE.WebGLRenderer({ antialias: true, powerPreference: "high-performance" });
  renderer.setSize(window.innerWidth, window.innerHeight);
  renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  renderer.shadowMap.enabled = true;
  renderer.shadowMap.type = THREE.PCFSoftShadowMap;
  container.appendChild(renderer.domElement);

  controls = new THREE.OrbitControls(camera, renderer.domElement);
  controls.enableDamping = true;
  controls.dampingFactor = 0.06;
  controls.maxPolarAngle = Math.PI / 2 - 0.01;
  controls.minDistance = 3;
  controls.maxDistance = 180;
  controls.target.set(0, 2, 4.8);

  // Lighting
  const ambientLight = new THREE.AmbientLight(0xffffff, 0.7);
  scene.add(ambientLight);

  const dirLight = new THREE.DirectionalLight(0xffffff, 0.85);
  dirLight.position.set(30, 45, 25);
  dirLight.castShadow = true;
  dirLight.shadow.mapSize.width = 2048;
  dirLight.shadow.mapSize.height = 2048;
  dirLight.shadow.camera.near = 5;
  dirLight.shadow.camera.far = 150;
  const d = 35;
  dirLight.shadow.camera.left = -d;
  dirLight.shadow.camera.right = d;
  dirLight.shadow.camera.top = d;
  dirLight.shadow.camera.bottom = -d;
  scene.add(dirLight);

  const fillLight = new THREE.DirectionalLight(0x38bdf8, 0.35);
  fillLight.position.set(-30, 25, -25);
  scene.add(fillLight);

  // Warehouse Floor, Walls & Pillars
  createWarehouseEnvironment();

  // Resize handler
  window.addEventListener('resize', onWindowResize);
  
  // Interaction handlers (Support both desktop mouse & mobile touch tap)
  renderer.domElement.addEventListener('mousemove', onMouseMove);
  renderer.domElement.addEventListener('pointerdown', onPointerDown);
  renderer.domElement.addEventListener('pointerup', onPointerUp);
}

function createWarehouseEnvironment() {
  // Expansive Open Floor (Unobstructed 3D View)
  const whWidth = 50;   // Barat - Timur
  const whDepth = 26;   // Utara - Selatan

  // 1. Concrete Dark Tech Floor
  const floorGeo = new THREE.PlaneGeometry(whWidth + 16, whDepth + 16);
  const floorMat = new THREE.MeshStandardMaterial({
    color: 0x0a0f1d,
    roughness: 0.8,
    metalness: 0.2
  });
  const floor = new THREE.Mesh(floorGeo, floorMat);
  floor.rotation.x = -Math.PI / 2;
  floor.receiveShadow = true;
  scene.add(floor);

  // Clean Grid Floor
  const gridHelper = new THREE.GridHelper(whWidth + 16, 48, 0x38bdf8, 0x1e293b);
  gridHelper.position.y = 0.01;
  scene.add(gridHelper);

  // 2. Subtle Floor Zone Markings (Area Selatan vs Area Utara)
  createZoneMarkings(whWidth, whDepth);

  // 3. Compass Markers (UTARA, SELATAN, BARAT, TIMUR)
  createCompassMarkers(whWidth, whDepth);
}

function createZoneMarkings(w, d) {
  // Garis Marka Lantai Pembatas Zona
  const lineGeo = new THREE.PlaneGeometry(w - 4, 0.12);
  const lineMat = new THREE.MeshBasicMaterial({ color: 0xf59e0b, side: THREE.DoubleSide });
  const line = new THREE.Mesh(lineGeo, lineMat);
  line.rotation.x = -Math.PI / 2;
  line.position.set(0, 0.02, -0.6);
  scene.add(line);

  // Label Lantai Flat Area Lorong Utara
  const aislePlane = createFloorTextPlane("◄ AREA LORONG UTAMA (AKSES JALAN GUDANG) ►", 16, 1.0, {
    bgColor: 'rgba(15, 23, 42, 0.4)',
    borderColor: 'rgba(255, 255, 255, 0.1)',
    textColor: '#64748b',
    fontSize: 26
  });
  aislePlane.position.set(0, 0.03, -5.5);
  scene.add(aislePlane);

  // Label Lantai Flat Area Tumpukan Bal Selatan
  const stockPlane = createFloorTextPlane("◄ ZONA TUMPUKAN BAL STOCK (16 BLOK) ►", 16, 1.0, {
    bgColor: 'rgba(15, 23, 42, 0.4)',
    borderColor: 'rgba(56, 189, 248, 0.2)',
    textColor: '#38bdf8',
    fontSize: 26
  });
  stockPlane.position.set(0, 0.03, 0.4);
  scene.add(stockPlane);
}

function createCompassMarkers(w, d) {
  // UTARA (North, Z negatif)
  const northBadge = createFloorTextPlane("▲ UTARA (NORTH)", 7, 0.9, {
    bgColor: 'rgba(15, 23, 42, 0.5)',
    borderColor: '#38bdf8',
    textColor: '#38bdf8',
    fontSize: 28
  });
  northBadge.position.set(0, 0.03, -(d / 2) + 2);
  scene.add(northBadge);

  // SELATAN (South, Z positif)
  const southBadge = createFloorTextPlane("▼ SELATAN (SOUTH)", 7, 0.9, {
    bgColor: 'rgba(15, 23, 42, 0.5)',
    borderColor: '#f59e0b',
    textColor: '#f59e0b',
    fontSize: 28
  });
  southBadge.position.set(0, 0.03, (d / 2) - 1);
  scene.add(southBadge);

  // BARAT (West, X paling negatif - Blok 1)
  const westBadge = createFloorTextPlane("◄ BARAT (BLOK 01)", 6.5, 0.9, {
    bgColor: 'rgba(15, 23, 42, 0.5)',
    borderColor: '#10b981',
    textColor: '#10b981',
    fontSize: 26
  });
  westBadge.position.set(-(w / 2) + 6, 0.03, 4.8);
  scene.add(westBadge);

  // TIMUR (East, X paling positif - Blok 16)
  const eastBadge = createFloorTextPlane("(BLOK 16) TIMUR ►", 6.5, 0.9, {
    bgColor: 'rgba(15, 23, 42, 0.5)',
    borderColor: '#ec4899',
    textColor: '#ec4899',
    fontSize: 26
  });
  eastBadge.position.set((w / 2) - 6, 0.03, 4.8);
  scene.add(eastBadge);
}

// ==============================================================================
// 2. DATA PROCESSING & 3D BALE GENERATION (16 BLOK DI SELATAN GUDANG DARI BARAT KE TIMUR)
// ==============================================================================
function loadInitialData() {
  if (window.WAREHOUSE_DATA) {
    buildWarehouse3D(window.WAREHOUSE_DATA);
  } else {
    fetch('warehouse_data.json')
      .then(res => res.json())
      .then(data => {
        window.WAREHOUSE_DATA = data;
        buildWarehouse3D(data);
      })
      .catch(err => {
        console.error('Failed to load warehouse data', err);
      });
  }
}

function getBlockPosition(blockId) {
  // 16 BLOK BERJAJAR DI SISI SELATAN GUDANG (Z = +5.0m):
  // Blok 1 di paling BARAT (X negatif), berurutan ke Timur s/d Blok 16 di paling TIMUR (X positif).
  const totalBlocks = 16;
  const blockSpanX = BALE_WIDTH + BALE_GAP_X + 0.35; // Jarak rapi antar blok
  const startX = -((totalBlocks - 1) / 2) * blockSpanX;
  const x = startX + (blockId - 1) * blockSpanX;
  const z = 4.8; // Posisi di sisi Selatan gudang
  return { x, z };
}

function buildWarehouse3D(rawData) {
  // Clear previous meshes
  baleMeshes.forEach(mesh => scene.remove(mesh));
  baleMeshes = [];
  balesData = [];
  Object.values(blockGroups).forEach(grp => scene.remove(grp));
  blockGroups = {};

  const blocks = rawData.blocks;
  const master = rawData.master || {};

  let totalBalesCount = 0;
  let totalWeightKg = 0;
  let gradeCounts = {};

  const boxGeometry = new THREE.BoxGeometry(BALE_WIDTH, BALE_HEIGHT, BALE_DEPTH);

  // Iterate all 16 blocks
  for (let bId = 1; bId <= 16; bId++) {
    const blockKey = String(bId);
    const bInfo = blocks[blockKey];
    if (!bInfo) continue;

    const blockGroup = new THREE.Group();
    const pos = getBlockPosition(bId);
    blockGroup.position.set(pos.x, 0, pos.z);
    blockGroups[bId] = blockGroup;
    scene.add(blockGroup);

    const headers = bInfo.headers; // Saf 1, Saf 2, etc.
    const numSaf = headers.length;
    const blockWidth = BALE_WIDTH + 0.2;
    const blockDepth = numSaf * (BALE_DEPTH + BALE_GAP_Z) + 0.4;

    // 1. Block Floor Pad & Border
    const padGeo = new THREE.BoxGeometry(blockWidth, 0.08, blockDepth);
    const padMat = new THREE.MeshStandardMaterial({
      color: 0x1e293b,
      roughness: 0.6,
      metalness: 0.2
    });
    const pad = new THREE.Mesh(padGeo, padMat);
    pad.position.set(0, 0.04, 0);
    pad.receiveShadow = true;
    blockGroup.add(pad);

    // Glowing border for Block
    // 2. Marka Lantai Blok Flat di Depan Saf 1 (Menempel di lantai, tidak menghalangi pandangan)
    const floorLabel = createFloorTextPlane(bInfo.title, blockWidth * 0.95, 0.7, {
      bgColor: 'rgba(14, 165, 233, 0.25)',
      borderColor: '#38bdf8',
      textColor: '#38bdf8',
      fontSize: 34
    });
    floorLabel.position.set(0, 0.05, -(blockDepth / 2) - 0.45);
    blockGroup.add(floorLabel);

    // 3. Process Bale Grid in Block
    // data is array of 4 rows: [T4, T3, T2, T1]
    const dataRows = bInfo.data; // 4 rows
    dataRows.forEach((row, rowIdx) => {
      const layerLevel = 4 - rowIdx; // 4, 3, 2, 1 (T4 at top, T1 at bottom)
      
      row.forEach((noGudVal, safIdx) => {
        if (noGudVal === "" || noGudVal === null || noGudVal === undefined) return;
        
        const noGudStr = String(noGudVal).trim();
        if (!noGudStr) return;

        // Lookup master metadata
        let masterInfo = master[noGudStr] || master[noGudStr.replace(/^0+/, '') || noGudStr];
        let barkot = masterInfo?.barkot || '-';
        let kg = masterInfo?.kg !== undefined && masterInfo?.kg !== '' ? masterInfo.kg : '-';
        let grade = masterInfo?.grade || 'UNGRADED';
        let status = masterInfo?.status || 'NORMAL';

        // Calculate numeric stats
        totalBalesCount++;
        const numericKg = parseFloat(kg);
        if (!isNaN(numericKg)) {
          totalWeightKg += numericKg;
        }
        gradeCounts[grade] = (gradeCounts[grade] || 0) + 1;

        // Spatial position inside block:
        // Saf 1 di UTARA (Z negatif) -> Saf N di SELATAN (Z positif)
        const safZ = (safIdx - (numSaf - 1) / 2) * (BALE_DEPTH + BALE_GAP_Z);
        const baseY = 0.08 + (layerLevel - 0.5) * (BALE_HEIGHT + BALE_GAP_Y);

        // Bale entity data
        const baleData = {
          id: `b_${bId}_s${safIdx + 1}_t${layerLevel}`,
          blockId: bId,
          blockTitle: bInfo.title,
          safIndex: safIdx + 1,
          safName: headers[safIdx] || `Saf ${safIdx + 1}`,
          layerLevel: layerLevel,
          layerName: `T${layerLevel}`,
          noGud: noGudStr,
          barkot: barkot,
          kg: kg,
          grade: grade,
          status: status,
          baseY: baseY,
          currentY: baseY,
          localX: 0,
          localZ: safZ,
          worldPos: new THREE.Vector3(pos.x, baseY, pos.z + safZ)
        };

        // Mesh creation
        const mat = getBaleMaterial(baleData, 'grade');
        const mesh = new THREE.Mesh(boxGeometry, mat);
        mesh.position.set(0, baseY, safZ);
        mesh.castShadow = true;
        mesh.receiveShadow = true;
        mesh.userData = baleData;

        // Add strapping bands (realistic look)
        addStrappingBands(mesh);

        blockGroup.add(mesh);
        baleMeshes.push(mesh);
        balesData.push(baleData);
      });
    });
  }

  // Update UI Stats & Dropdowns
  updateStatsDashboard(totalBalesCount, totalWeightKg, gradeCounts);
  populateFilterOptions(gradeCounts);
}

// Realistic Tobacco Bale Texture Generator & Cache (High DPI & Big Sharp Numbers)
const baleTextureCache = {};

function getBaleCanvasTexture(baleData, colorMode) {
  const cacheKey = `${baleData.noGud}_${baleData.grade}_${colorMode}`;
  if (baleTextureCache[cacheKey]) return baleTextureCache[cacheKey];

  const canvas = document.createElement('canvas');
  canvas.width = 512;
  canvas.height = 512;
  const ctx = canvas.getContext('2d');

  // Base Burlap / Jute Sack Color
  let baseColor = '#c8a876'; // Natural golden tan burlap (Karung Goni)

  if (colorMode === 'grade') {
    const gKey = baleData.grade ? baleData.grade.trim().toUpperCase() : 'DEFAULT';
    const hex = GRADE_COLORS[gKey] || stringToColor(gKey);
    baseColor = '#' + hex.toString(16).padStart(6, '0');
  } else if (colorMode === 'layer') {
    const hex = LAYER_COLORS[baleData.layerLevel] || 0x64748b;
    baseColor = '#' + hex.toString(16).padStart(6, '0');
  } else if (colorMode === 'weight') {
    const w = parseFloat(baleData.kg);
    if (!isNaN(w)) {
      const ratio = Math.min(Math.max((w - 20) / 45, 0), 1);
      const c1 = new THREE.Color(0x38bdf8);
      const c2 = new THREE.Color(0xf59e0b);
      const c3 = new THREE.Color(0xef4444);
      const col = (ratio < 0.5) ? c1.lerp(c2, ratio * 2) : c2.lerp(c3, (ratio - 0.5) * 2);
      baseColor = '#' + col.getHexString();
    }
  } else if (colorMode === 'realistic') {
    baseColor = '#c49a65'; // Warm tobacco burlap
  }

  // 1. Fill base jute tone
  ctx.fillStyle = baseColor;
  ctx.fillRect(0, 0, 512, 512);

  // 2. Add Jute Woven Pattern (Fibers cross-hatching)
  ctx.fillStyle = 'rgba(0, 0, 0, 0.08)';
  for (let x = 0; x < 512; x += 10) {
    ctx.fillRect(x, 0, 5, 512);
  }
  ctx.fillStyle = 'rgba(255, 255, 255, 0.07)';
  for (let y = 0; y < 512; y += 10) {
    ctx.fillRect(0, y, 512, 5);
  }

  // 3. Stitched border edge
  ctx.strokeStyle = 'rgba(0, 0, 0, 0.35)';
  ctx.lineWidth = 8;
  ctx.strokeRect(6, 6, 500, 500);

  // 4. White Warehouse Stencil / Paper Label (GIANT & HIGH CONTRAST)
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(36, 50, 440, 412);
  ctx.strokeStyle = '#0f172a';
  ctx.lineWidth = 8;
  ctx.strokeRect(36, 50, 440, 412);

  // Print BIG BOLD No. Gudang
  ctx.fillStyle = '#000000';
  ctx.font = "900 125px 'Inter', sans-serif";
  ctx.textAlign = 'center';
  ctx.fillText(String(baleData.noGud), 256, 195);

  // Print Grade and Weight
  ctx.font = "bold 44px 'Inter', sans-serif";
  ctx.fillStyle = '#1e293b';
  const tagGrade = baleData.grade && baleData.grade !== 'UNGRADED' ? `GR: ${baleData.grade}` : '';
  const tagKg = (baleData.kg && baleData.kg !== '-') ? `${baleData.kg}kg` : '';
  const subInfo = [tagGrade, tagKg].filter(Boolean).join(' • ') || 'TEMBAKAU';
  ctx.fillText(subInfo, 256, 295);

  // Print Barcode
  if (baleData.barkot && baleData.barkot !== '-') {
    ctx.font = "bold 36px monospace";
    ctx.fillStyle = '#0284c7';
    ctx.fillText(`*${baleData.barkot}*`, 256, 385);
  }

  const texture = new THREE.CanvasTexture(canvas);
  texture.minFilter = THREE.LinearFilter;
  baleTextureCache[cacheKey] = texture;
  return texture;
}

function addStrappingBands(mesh) {
  // Realistic dual nylon strapping bands around tobacco bale
  const bandMat = new THREE.MeshStandardMaterial({
    color: 0x18181b, // Dark durable strapping band
    roughness: 0.3,
    metalness: 0.2
  });

  // 1. Horizontal center band
  const hBandGeo = new THREE.BoxGeometry(BALE_WIDTH + 0.015, 0.04, BALE_DEPTH + 0.015);
  const hBand = new THREE.Mesh(hBandGeo, bandMat);
  hBand.position.set(0, 0, 0);
  mesh.add(hBand);

  // 2. Vertical band Left
  const vBandGeo = new THREE.BoxGeometry(0.04, BALE_HEIGHT + 0.015, BALE_DEPTH + 0.015);
  const vBand1 = new THREE.Mesh(vBandGeo, bandMat);
  vBand1.position.set(-BALE_WIDTH * 0.28, 0, 0);
  mesh.add(vBand1);

  // 3. Vertical band Right
  const vBand2 = new THREE.Mesh(vBandGeo, bandMat);
  vBand2.position.set(BALE_WIDTH * 0.28, 0, 0);
  mesh.add(vBand2);
}

function getBaleMaterial(bale, colorMode) {
  const texture = getBaleCanvasTexture(bale, colorMode);

  return new THREE.MeshStandardMaterial({
    map: texture,
    roughness: 0.75,
    metalness: 0.1,
    transparent: true,
    opacity: 0.95
  });
}

function stringToColor(str) {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = str.charCodeAt(i) + ((hash << 5) - hash);
  }
  const c = (hash & 0x00ffffff).toString(16).toUpperCase();
  return parseInt('00000'.substring(0, 6 - c.length) + c, 16);
}

// Flat Floor Text Plane Helper (Never blocks the camera or 3D view)
function createFloorTextPlane(text, width = 2.0, height = 0.8, options = {}) {
  const canvas = document.createElement('canvas');
  canvas.width = 512;
  canvas.height = 256;
  const ctx = canvas.getContext('2d');

  const fontsize = options.fontSize || 32;
  const textColor = options.textColor || '#38bdf8';
  const bgColor = options.bgColor || 'rgba(15, 23, 42, 0.6)';
  const borderColor = options.borderColor || 'rgba(56, 189, 248, 0.6)';

  ctx.clearRect(0, 0, canvas.width, canvas.height);
  
  // Floor background patch
  ctx.fillStyle = bgColor;
  ctx.fillRect(8, 8, 496, 240);

  ctx.strokeStyle = borderColor;
  ctx.lineWidth = 6;
  ctx.strokeRect(8, 8, 496, 240);

  ctx.font = `Bold ${fontsize}px 'Inter', sans-serif`;
  ctx.fillStyle = textColor;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 256, 128);

  const texture = new THREE.CanvasTexture(canvas);
  texture.minFilter = THREE.LinearFilter;
  const mat = new THREE.MeshBasicMaterial({
    map: texture,
    transparent: true,
    side: THREE.DoubleSide
  });

  const geo = new THREE.PlaneGeometry(width, height);
  const mesh = new THREE.Mesh(geo, mat);
  mesh.rotation.set(-Math.PI / 2, 0, Math.PI); // Flat on the floor and right-side up facing North
  return mesh;
}

// ==============================================================================
// 3. UI, FILTERING & INTERACTION LOGIC
// ==============================================================================
function updateStatsDashboard(totalBales, totalWeight, gradeCounts) {
  document.getElementById('stat-total-bales').innerText = totalBales;
  document.getElementById('stat-total-weight').innerText = (totalWeight / 1000).toFixed(2);
  const avg = totalBales > 0 ? (totalWeight / totalBales).toFixed(1) : '0';
  document.getElementById('stat-avg-weight').innerText = avg;
  
  // Capacity calculation (16 blocks * 6 saf * 4 levels = 384 max active capacity)
  const occupancy = Math.min(Math.round((totalBales / 384) * 100), 100);
  document.getElementById('stat-occupancy').innerText = occupancy;

  // Grade Legend update
  updateGradeLegend(gradeCounts);
}

function updateGradeLegend(gradeCounts) {
  const legendContainer = document.getElementById('grade-legend-list');
  if (!legendContainer) return;
  legendContainer.innerHTML = '';

  const sortedGrades = Object.entries(gradeCounts).sort((a, b) => b[1] - a[1]);
  sortedGrades.slice(0, 8).forEach(([grade, count]) => {
    const item = document.createElement('div');
    item.className = 'legend-item';
    const hexColor = (GRADE_COLORS[grade] !== undefined) 
      ? '#' + GRADE_COLORS[grade].toString(16).padStart(6, '0') 
      : '#64748b';
    
    item.innerHTML = `
      <span class="legend-color" style="background-color: ${hexColor}"></span>
      <span>${grade}: <b>${count}</b></span>
    `;
    legendContainer.appendChild(item);
  });
}

function populateFilterOptions(gradeCounts) {
  const gradeSelect = document.getElementById('filter-grade');
  if (!gradeSelect) return;
  
  // Keep the 'All Grades' option
  gradeSelect.innerHTML = '<option value="all">Semua Grade</option>';
  Object.keys(gradeCounts).sort().forEach(g => {
    const opt = document.createElement('option');
    opt.value = g;
    opt.innerText = `Grade ${g} (${gradeCounts[g]})`;
    gradeSelect.appendChild(opt);
  });
}

function setupUIEventListeners() {
  // Mobile Panel Toggle & Close
  const togglePanelBtn = document.getElementById('btn-toggle-panel');
  const closePanelBtn = document.getElementById('btn-close-panel');
  const panelBackdrop = document.getElementById('panel-backdrop');
  const leftPanel = document.getElementById('left-panel');

  if (togglePanelBtn && leftPanel) {
    togglePanelBtn.addEventListener('click', () => {
      leftPanel.classList.add('open');
      if (panelBackdrop) panelBackdrop.classList.add('active');
    });
  }

  function closeMobilePanel() {
    if (leftPanel) leftPanel.classList.remove('open');
    if (panelBackdrop) panelBackdrop.classList.remove('active');
  }

  if (closePanelBtn) closePanelBtn.addEventListener('click', closeMobilePanel);
  if (panelBackdrop) panelBackdrop.addEventListener('click', closeMobilePanel);

  // Color Mode selection
  document.querySelectorAll('.color-mode-pills .pill-option').forEach(pill => {
    pill.addEventListener('click', (e) => {
      document.querySelectorAll('.color-mode-pills .pill-option').forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      activeFilter.colorMode = pill.getAttribute('data-mode');
      applyFilters();
    });
  });

  // Layer level buttons
  document.querySelectorAll('.layer-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const lvl = parseInt(btn.getAttribute('data-layer'));
      activeFilter.layers[lvl] = !activeFilter.layers[lvl];
      btn.classList.toggle('active', activeFilter.layers[lvl]);
      applyFilters();
    });
  });

  // Block Select
  const blockSelect = document.getElementById('filter-block');
  if (blockSelect) {
    blockSelect.addEventListener('change', (e) => {
      activeFilter.block = e.target.value;
      if (e.target.value !== 'all') {
        focusOnBlock(parseInt(e.target.value));
      } else {
        resetCameraOverview();
      }
      applyFilters();
    });
  }

  // Grade Select
  const gradeSelect = document.getElementById('filter-grade');
  if (gradeSelect) {
    gradeSelect.addEventListener('change', (e) => {
      activeFilter.grade = e.target.value;
      applyFilters();
    });
  }

  // Explode Layers Slider
  const explodeSlider = document.getElementById('slider-explode');
  if (explodeSlider) {
    explodeSlider.addEventListener('input', (e) => {
      activeFilter.explodeFactor = parseFloat(e.target.value);
      document.getElementById('explode-val').innerText = `${activeFilter.explodeFactor}x`;
      updateLayerExplosion();
    });
  }

  // Search Input (Real-time unified search for No Gud & Barkot with Suggestions)
  const searchInput = document.getElementById('input-search');
  const suggestionsBox = document.getElementById('search-suggestions');

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      activeFilter.searchQuery = e.target.value.trim().toLowerCase();
      applyFilters();
      
      if (activeFilter.searchQuery) {
        renderSearchSuggestions(activeFilter.searchQuery);
        highlightFirstSearchMatch(activeFilter.searchQuery);
      } else {
        if (suggestionsBox) {
          suggestionsBox.innerHTML = '';
          suggestionsBox.classList.remove('show');
        }
        hideLocationFinderHUD();
        removeBeacon();
        resetCameraOverview();
      }
    });

    searchInput.addEventListener('focus', () => {
      if (activeFilter.searchQuery) renderSearchSuggestions(activeFilter.searchQuery);
    });

    // Close suggestions when clicked outside
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.search-wrapper') && suggestionsBox) {
        suggestionsBox.classList.remove('show');
      }
    });
  }

  // 2D Denah Blueprint Modal Toggle & Close
  const toggle2dBtn = document.getElementById('btn-toggle-2d-view');
  const close2dBtn = document.getElementById('close-denah-2d-btn');
  const denahModal = document.getElementById('denah-2d-modal');

  if (toggle2dBtn && denahModal) {
    toggle2dBtn.addEventListener('click', () => {
      renderAllBlocks2DGrid();
      denahModal.classList.add('show');
    });
  }

  if (close2dBtn && denahModal) {
    close2dBtn.addEventListener('click', () => {
      denahModal.classList.remove('show');
    });
  }

  // Mobile Bottom Bar Navigation Actions
  const mBtnDenah = document.getElementById('m-btn-denah');
  const mBtnSearch = document.getElementById('m-btn-search');
  const mBtnFilter = document.getElementById('m-btn-filter');
  const mBtnReset = document.getElementById('m-btn-reset');

  if (mBtnDenah && denahModal) {
    mBtnDenah.addEventListener('click', () => {
      renderAllBlocks2DGrid();
      denahModal.classList.add('show');
    });
  }

  if (mBtnSearch && searchInput) {
    mBtnSearch.addEventListener('click', () => {
      searchInput.focus();
      window.scrollTo({ top: 0, behavior: 'smooth' });
      if (activeFilter.searchQuery) renderSearchSuggestions(activeFilter.searchQuery);
    });
  }

  if (mBtnFilter && leftPanel) {
    mBtnFilter.addEventListener('click', () => {
      leftPanel.classList.add('open');
      if (panelBackdrop) panelBackdrop.classList.add('active');
    });
  }

  if (mBtnReset) {
    mBtnReset.addEventListener('click', resetAllUI);
  }

  const btnResetAll = document.getElementById('btn-reset-all');
  if (btnResetAll) {
    btnResetAll.addEventListener('click', resetAllUI);
  }

  // Close Location HUD Button
  const closeHudBtn = document.getElementById('close-hud-btn');
  if (closeHudBtn) {
    closeHudBtn.addEventListener('click', () => {
      hideLocationFinderHUD();
      removeBeacon();
      applyFilters();
    });
  }

  // Camera preset buttons
  document.querySelectorAll('.cam-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.cam-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const view = btn.getAttribute('data-cam');
      setCameraPreset(view);
    });
  });

  // Quick Block Bar Tab Buttons
  document.querySelectorAll('.block-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.block-tab-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const bVal = btn.getAttribute('data-block');
      const blockSelect = document.getElementById('filter-block');
      if (blockSelect) blockSelect.value = bVal;

      activeFilter.block = bVal;
      if (bVal !== 'all') {
        focusOnBlock(parseInt(bVal));
      } else {
        resetCameraOverview();
      }
      applyFilters();
    });
  });

  // Drawer Close Button
  const closeBtn = document.getElementById('close-drawer-btn');
  if (closeBtn) {
    closeBtn.addEventListener('click', () => {
      document.getElementById('inspector-drawer').classList.remove('open');
      selectedBale = null;
      removeBeacon();
      hideLocationFinderHUD();
      applyFilters();
    });
  }

  // Excel Upload Modal trigger
  const uploadBtn = document.getElementById('btn-upload-excel');
  const modal = document.getElementById('upload-modal');
  const closeModalBtn = document.getElementById('close-modal-btn');
  if (uploadBtn && modal) {
    uploadBtn.addEventListener('click', () => modal.classList.add('show'));
    closeModalBtn.addEventListener('click', () => modal.classList.remove('show'));
  }

  // Dropzone file handling
  setupExcelDropzone();
}

function updateLayerExplosion() {
  baleMeshes.forEach(mesh => {
    const data = mesh.userData;
    // T1 stays at floor, T2 rises +1*gap, T3 rises +2*gap, T4 rises +3*gap
    const explodeOffset = (data.layerLevel - 1) * (activeFilter.explodeFactor * 1.5);
    mesh.position.y = data.baseY + explodeOffset;
  });
}

function applyFilters() {
  const query = activeFilter.searchQuery;
  const targetGrade = activeFilter.grade;
  const targetBlock = activeFilter.block;

  baleMeshes.forEach(mesh => {
    const d = mesh.userData;
    let isVisible = true;
    let isMatch = true;

    // Check layer visibility
    if (!activeFilter.layers[d.layerLevel]) {
      isVisible = false;
    }

    // Check block filter
    if (targetBlock !== 'all' && String(d.blockId) !== String(targetBlock)) {
      isVisible = false;
    }

    // Check grade filter
    if (targetGrade !== 'all' && d.grade !== targetGrade) {
      isVisible = false;
    }

    // Check search query (No Gud or Barkot)
    if (query) {
      const matchNoGud = d.noGud && d.noGud.toLowerCase().includes(query);
      const matchBarkot = d.barkot && d.barkot.toLowerCase().includes(query);
      if (!matchNoGud && !matchBarkot) {
        isMatch = false;
      }
    }

    mesh.visible = isVisible;

    // Update material according to active color mode & match state
    const baseMat = getBaleMaterial(d, activeFilter.colorMode);
    
    if (query && !isMatch) {
      // Ghosted mode for non-matching bales
      baseMat.opacity = 0.12;
      baseMat.transparent = true;
    } else if (selectedBale && selectedBale.id === d.id) {
      // Highlight selected with glowing emissive color
      baseMat.emissive = new THREE.Color(0x38bdf8);
      baseMat.emissiveIntensity = 0.7;
    } else if (query && isMatch) {
      // Highlight search result
      baseMat.emissive = new THREE.Color(0x10b981);
      baseMat.emissiveIntensity = 0.85;
    }

    mesh.material = baseMat;
  });
}

function highlightFirstSearchMatch(query) {
  const match = baleMeshes.find(m => {
    const d = m.userData;
    return m.visible && ((d.noGud && d.noGud.toLowerCase().includes(query)) || (d.barkot && d.barkot.toLowerCase().includes(query)));
  });

  if (match) {
    selectBale(match.userData, match);
  }
}

// ==============================================================================
// 4. RAYCASTING, TOUCH TAP & INTERACTIVE INSPECTOR
// ==============================================================================
let beaconMesh = null;
let pointerDownPos = { x: 0, y: 0 };

function onPointerDown(event) {
  pointerDownPos.x = event.clientX;
  pointerDownPos.y = event.clientY;
}

function onPointerUp(event) {
  // Hanya proses jika bukan gerakan drag/orbit (jarak gerak < 10px)
  const dist = Math.hypot(event.clientX - pointerDownPos.x, event.clientY - pointerDownPos.y);
  if (dist > 10) return;

  // Lakukan raycast langsung dari koordinat tap/klik layar
  mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
  mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;

  raycaster.setFromCamera(mouse, camera);
  const visibleMeshes = baleMeshes.filter(m => m.visible);
  const intersects = raycaster.intersectObjects(visibleMeshes, false);

  if (intersects.length > 0) {
    const hitMesh = intersects[0].object;
    if (hitMesh && hitMesh.userData && hitMesh.userData.id) {
      selectBale(hitMesh.userData, hitMesh);
    }
  }
}

function onMouseMove(event) {
  mouse.x = (event.clientX / window.innerWidth) * 2 - 1;
  mouse.y = -(event.clientY / window.innerHeight) * 2 + 1;

  raycaster.setFromCamera(mouse, camera);
  const visibleMeshes = baleMeshes.filter(m => m.visible);
  const intersects = raycaster.intersectObjects(visibleMeshes, false);

  const tooltip = document.getElementById('hover-tooltip');

  if (intersects.length > 0) {
    const hitMesh = intersects[0].object;
    hoveredBale = hitMesh.userData;
    document.body.style.cursor = 'pointer';

    // Show floating tooltip (hanya di desktop/mouse)
    if (tooltip && window.innerWidth > 1024) {
      tooltip.innerHTML = `
        <div style="font-weight:700; color:#38bdf8;">No. Gudang: ${hoveredBale.noGud}</div>
        <div style="font-size:11px; color:#cbd5e1;">Grade: <b>${hoveredBale.grade}</b> | ${hoveredBale.kg} kg</div>
        <div style="font-size:10px; color:#94a3b8;">${hoveredBale.blockTitle} - ${hoveredBale.safName} (${hoveredBale.layerName})</div>
      `;
      tooltip.style.left = `${event.clientX}px`;
      tooltip.style.top = `${event.clientY - 10}px`;
      tooltip.style.opacity = '1';
    }
  } else {
    hoveredBale = null;
    document.body.style.cursor = 'default';
    if (tooltip) tooltip.style.opacity = '0';
  }
}

function selectBale(baleData, mesh) {
  selectedBale = baleData;
  applyFilters();

  // Populate Right Drawer
  const drawer = document.getElementById('inspector-drawer');
  document.getElementById('inspect-nogud').innerText = baleData.noGud;
  document.getElementById('inspect-grade').innerText = baleData.grade || '-';
  document.getElementById('inspect-barkot').innerText = baleData.barkot || '-';
  document.getElementById('inspect-kg').innerText = baleData.kg ? `${baleData.kg} Kg` : '-';
  document.getElementById('inspect-block').innerText = `${baleData.blockTitle}`;
  document.getElementById('inspect-saf').innerText = `${baleData.safName}`;
  document.getElementById('inspect-tingkat').innerText = `${baleData.layerName}`;
  document.getElementById('inspect-status').innerText = baleData.status || 'NORMAL';

  drawer.classList.add('open');

  // Smooth camera pan:
  // Jika sedang memfilter blok tertentu -> KAMERA OTOMATIS MELIHAT DARI SAMPING (Side Profile View)
  // sehingga terlihat jelas seluruh Saf 1 (Kiri) s/d Saf 6 (Kanan) dan Tingkat T1 (Bawah) s/d T4 (Atas)
  if (mesh) {
    const worldPos = new THREE.Vector3();
    mesh.getWorldPosition(worldPos);
    
    if (activeFilter.block !== 'all') {
      // Sudut pandang samping satunya (Sisi Timur/Selatan) memperlihatkan profil blok
      smoothCameraFly(worldPos.x + 8.0, 4.5, worldPos.z + 1.5, worldPos.x, 1.8, worldPos.z);
    } else {
      // Jarak menengah lega dari sisi Utara jika mode semua blok
      smoothCameraFly(worldPos.x, 10.0, worldPos.z - 15.0, worldPos.x, 2.0, worldPos.z);
    }
    createOrUpdateBeacon(worldPos);
  }

  // Tampilkan HUD Banner Lokasi Cerdas di Tengah Layar
  showLocationFinderHUD(baleData);

  // Render Denah Matriks Blok
  renderBlockMatrix(baleData);
}

function showLocationFinderHUD(baleData) {
  const hud = document.getElementById('location-finder-hud');
  if (!hud) return;

  const barkotText = (baleData.barkot && baleData.barkot !== '-') ? ` • BARCODE: ${baleData.barkot}` : '';
  const kgText = (baleData.kg && baleData.kg !== '-') ? ` • ${baleData.kg} Kg` : '';
  const gradeText = (baleData.grade && baleData.grade !== 'UNGRADED') ? ` • Grade ${baleData.grade}` : '';

  document.getElementById('hud-match-title').innerText = `LOKASI BAL NO. GUDANG #${baleData.noGud}${barkotText}${gradeText}${kgText}`;
  document.getElementById('hud-block').innerText = baleData.blockTitle;
  document.getElementById('hud-saf').innerText = baleData.safName;
  document.getElementById('hud-tingkat').innerText = `TINGKAT ${baleData.layerLevel} (${baleData.layerName})`;

  hud.classList.add('show');
}

function hideLocationFinderHUD() {
  const hud = document.getElementById('location-finder-hud');
  if (hud) hud.classList.remove('show');
}

function createOrUpdateBeacon(worldPos) {
  removeBeacon();

  const beaconGroup = new THREE.Group();

  // 1. Vertical glowing laser beam cylinder setinggi 12 meter
  const beamGeo = new THREE.CylinderGeometry(0.05, 0.28, 12, 16);
  const beamMat = new THREE.MeshBasicMaterial({
    color: 0x10b981,
    transparent: true,
    opacity: 0.75
  });
  const beam = new THREE.Mesh(beamGeo, beamMat);
  beam.position.set(0, 6.5, 0);
  beaconGroup.add(beam);

  // 2. Target indicator ring at top of bale
  const ringGeo = new THREE.RingGeometry(0.4, 0.65, 24);
  const ringMat = new THREE.MeshBasicMaterial({ color: 0x38bdf8, side: THREE.DoubleSide });
  const ring = new THREE.Mesh(ringGeo, ringMat);
  ring.rotation.x = -Math.PI / 2;
  ring.position.set(0, 0.48, 0);
  beaconGroup.add(ring);

  beaconGroup.position.set(worldPos.x, worldPos.y + BALE_HEIGHT / 2, worldPos.z);
  scene.add(beaconGroup);
  beaconMesh = beaconGroup;
}

function removeBeacon() {
  if (beaconMesh) {
    scene.remove(beaconMesh);
    beaconMesh = null;
  }
}

function renderBlockMatrix(baleData) {
  const raw = window.WAREHOUSE_DATA;
  if (!raw || !raw.blocks) return;
  const bInfo = raw.blocks[String(baleData.blockId)];
  if (!bInfo) return;

  document.getElementById('matrix-block-name').innerText = bInfo.title;

  const container = document.getElementById('matrix-grid-view');
  const headers = bInfo.headers; // Saf 1 .. Saf N
  const dataRows = bInfo.data;   // [T4, T3, T2, T1]

  let html = `<table class="matrix-grid-table"><thead><tr><th>Tingkat</th>`;
  headers.forEach(h => {
    html += `<th>${h}</th>`;
  });
  html += `</tr></thead><tbody>`;

  let obstructingBales = [];

  dataRows.forEach((row, rowIdx) => {
    const layerLevel = 4 - rowIdx; // 4, 3, 2, 1
    html += `<tr><td style="font-weight:700; background:rgba(30,41,59,0.7); color:#94a3b8;">T${layerLevel}</td>`;
    
    row.forEach((noGud, safIdx) => {
      const isSelected = (baleData.safIndex === safIdx + 1 && baleData.layerLevel === layerLevel);
      const isAbove = (baleData.safIndex === safIdx + 1 && layerLevel > baleData.layerLevel && noGud);

      let cellClass = '';
      if (isSelected) {
        cellClass = 'selected-cell';
      } else if (isAbove) {
        cellClass = 'obstruct-cell';
        obstructingBales.push(`T${layerLevel} (#${noGud})`);
      } else if (!noGud) {
        cellClass = 'empty-cell';
      }

      html += `<td class="${cellClass}" onclick="handleMatrixCellClick(${baleData.blockId}, ${safIdx + 1}, ${layerLevel})">${noGud || '-'}</td>`;
    });
    html += `</tr>`;
  });

  html += `</tbody></table>`;
  container.innerHTML = html;

  // Unstack Alert Card
  const alertBox = document.getElementById('unstack-alert-box');
  const alertDesc = document.getElementById('unstack-desc');
  if (obstructingBales.length > 0) {
    alertBox.style.display = 'block';
    alertDesc.innerHTML = `Untuk mengambil bal <b>#${baleData.noGud}</b> (${baleData.safName} - T${baleData.layerLevel}), Anda perlu memindahkan bal di atasnya:<br><span style="color:#fef08a; font-weight:700;">${obstructingBales.join(' ➔ ')}</span>.`;
  } else {
    alertBox.style.display = 'none';
  }
}

window.handleMatrixCellClick = function(blockId, safIndex, layerLevel) {
  const targetBale = balesData.find(b => b.blockId === blockId && b.safIndex === safIndex && b.layerLevel === layerLevel);
  if (targetBale) {
    const targetMesh = baleMeshes.find(m => m.userData.id === targetBale.id);
    selectBale(targetBale, targetMesh);
  }
};

// ==============================================================================
// 5. SEARCH AUTOCOMPLETE & 2D BLUEPRINT MATRIX
// ==============================================================================
function renderSearchSuggestions(query) {
  const box = document.getElementById('search-suggestions');
  if (!box) return;

  const matches = balesData.filter(d => {
    const matchNoGud = d.noGud && String(d.noGud).toLowerCase().includes(query);
    const matchBarkot = d.barkot && String(d.barkot).toLowerCase().includes(query);
    const matchGrade = d.grade && String(d.grade).toLowerCase().includes(query);
    return matchNoGud || matchBarkot || matchGrade;
  }).slice(0, 10);

  if (matches.length === 0) {
    box.innerHTML = `<div style="padding:12px; font-size:12px; color:var(--text-muted); text-align:center;">Tidak ada bal yang cocok</div>`;
    box.classList.add('show');
    return;
  }

  let html = '';
  matches.forEach(bale => {
    const barkotInfo = (bale.barkot && bale.barkot !== '-') ? ` • Barkot: ${bale.barkot}` : '';
    const kgInfo = (bale.kg && bale.kg !== '-') ? ` • ${bale.kg}kg` : '';
    html += `
      <div class="search-suggestion-item" onclick="selectSuggestionBale(${bale.id})">
        <div>
          <div class="sugg-main"><i class="fa-solid fa-cube" style="color:var(--accent-cyan);"></i> Bal #${bale.noGud}</div>
          <div class="sugg-loc">${bale.blockTitle} ➔ ${bale.safName} (T${bale.layerLevel})</div>
        </div>
        <div class="sugg-meta">
          <span style="color:#fef08a; font-weight:700;">${bale.grade || '-'}</span>${kgInfo}${barkotInfo}
        </div>
      </div>
    `;
  });

  box.innerHTML = html;
  box.classList.add('show');
}

window.selectSuggestionBale = function(baleId) {
  const targetBale = balesData.find(b => b.id === baleId);
  if (targetBale) {
    const targetMesh = baleMeshes.find(m => m.userData.id === baleId);
    selectBale(targetBale, targetMesh);
    const box = document.getElementById('search-suggestions');
    if (box) box.classList.remove('show');
  }
};

function renderAllBlocks2DGrid() {
  const container = document.getElementById('all-blocks-matrix-container');
  if (!container) return;

  let html = '';
  for (let bId = 1; bId <= 16; bId++) {
    const blockBales = balesData.filter(b => b.blockId === bId);
    if (blockBales.length === 0) continue;

    html += `
      <div class="block-section-card">
        <div class="block-section-title">
          <i class="fa-solid fa-cubes-stacked"></i> BLOK ${String(bId).padStart(2, '0')} (Total: ${blockBales.length} Bal)
        </div>
        <table class="matrix-grid-table" style="font-size:11px;">
          <thead>
            <tr>
              <th style="width: 50px;">Tingkat</th>
              <th>Saf 1 (Utara)</th>
              <th>Saf 2</th>
              <th>Saf 3</th>
              <th>Saf 4</th>
              <th>Saf 5</th>
              <th>Saf 6 (Selatan)</th>
            </tr>
          </thead>
          <tbody>
    `;

    for (let lvl = 4; lvl >= 1; lvl--) {
      html += `<tr><td style="font-weight:700; color:var(--text-muted);">T${lvl}</td>`;
      for (let sIdx = 1; sIdx <= 6; sIdx++) {
        const found = blockBales.find(b => b.safIndex === sIdx && b.layerLevel === lvl);
        if (found) {
          const noGud = found.noGud;
          const gr = (found.grade && found.grade !== 'UNGRADED') ? ` <span style="font-size:9px; color:#fef08a;">(${found.grade})</span>` : '';
          html += `<td style="cursor:pointer; background:rgba(16,185,129,0.22); border:1px solid rgba(16,185,129,0.5); font-weight:700; color:#fff; padding:6px 2px;" onclick="jumpFrom2DTo3D(${found.id})">#${noGud}${gr}</td>`;
        } else {
          html += `<td style="color:#475569; background:rgba(15,23,42,0.4); border:1px solid rgba(51,65,85,0.3);">-</td>`;
        }
      }
      html += `</tr>`;
    }

    html += `</tbody></table></div>`;
  }

  container.innerHTML = html;
}

window.jumpFrom2DTo3D = function(baleId) {
  const modal = document.getElementById('denah-2d-modal');
  if (modal) modal.classList.remove('show');
  const targetBale = balesData.find(b => b.id === baleId);
  if (targetBale) {
    const targetMesh = baleMeshes.find(m => m.userData.id === baleId);
    selectBale(targetBale, targetMesh);
  }
};

// ==============================================================================
// 6. CAMERA ANIMATION & PRESETS
// ==============================================================================
function setCameraPreset(preset) {
  const isMobile = window.innerWidth <= 768;
  if (preset === 'iso') {
    smoothCameraFly(0, isMobile ? 18 : 20, isMobile ? -24 : -18, 0, 1.8, 4.8);
  } else if (preset === 'top') {
    smoothCameraFly(0, isMobile ? 48 : 42, 4.8, 0, 0, 4.8);
  } else if (preset === 'front') {
    smoothCameraFly(0, 5, -8, 0, 2, 4.8);
  } else if (preset === 'west') {
    smoothCameraFly(-10, 12, -6, -10, 2, 4.8);
  } else if (preset === 'east') {
    smoothCameraFly(10, 12, -6, 10, 2, 4.8);
  }
}

function resetAllUI() {
  // 1. Reset Camera to Centered Overview
  resetCameraOverview();

  // 2. Clear Search & Suggestions
  const searchInput = document.getElementById('input-search');
  if (searchInput) searchInput.value = '';
  activeFilter.searchQuery = '';
  const suggestionsBox = document.getElementById('search-suggestions');
  if (suggestionsBox) {
    suggestionsBox.innerHTML = '';
    suggestionsBox.classList.remove('show');
  }

  // 3. Hide Location Finder HUD Banner
  hideLocationFinderHUD();

  // 4. Clear Bale Selection & Laser Beacon
  selectedBale = null;
  removeBeacon();

  // 5. Close All Drawers & Modals
  const drawer = document.getElementById('inspector-drawer');
  if (drawer) drawer.classList.remove('open');
  const leftPanel = document.getElementById('left-panel');
  if (leftPanel) leftPanel.classList.remove('open');
  const panelBackdrop = document.getElementById('panel-backdrop');
  if (panelBackdrop) panelBackdrop.classList.remove('active');
  const denahModal = document.getElementById('denah-2d-modal');
  if (denahModal) denahModal.classList.remove('show');
  const uploadModal = document.getElementById('upload-modal');
  if (uploadModal) uploadModal.classList.remove('show');

  // 6. Reset Filters (Block, Grade, Layers, Explode, ColorMode)
  activeFilter.block = 'all';
  activeFilter.grade = 'all';
  activeFilter.layers = { 1: true, 2: true, 3: true, 4: true };
  activeFilter.explodeFactor = 0;
  activeFilter.colorMode = 'grade';

  // Update UI Inputs & Buttons
  const blockSelect = document.getElementById('filter-block');
  if (blockSelect) blockSelect.value = 'all';

  const gradeSelect = document.getElementById('filter-grade');
  if (gradeSelect) gradeSelect.value = 'all';

  document.querySelectorAll('.block-tab-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('data-block') === 'all');
  });

  document.querySelectorAll('.layer-btn').forEach(b => {
    b.classList.add('active');
  });

  const explodeSlider = document.getElementById('slider-explode');
  if (explodeSlider) explodeSlider.value = 0;
  const explodeVal = document.getElementById('explode-val');
  if (explodeVal) explodeVal.innerText = '0x';
  updateLayerExplosion();

  document.querySelectorAll('.color-mode-pills .pill-option').forEach(p => {
    p.classList.toggle('active', p.getAttribute('data-mode') === 'grade');
  });

  document.querySelectorAll('.cam-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('data-cam') === 'iso');
  });

  // 7. Re-apply Filters to Restore 100% Opacity and Materials
  applyFilters();
}

function resetCameraOverview() {
  const isMobile = window.innerWidth <= 768;
  if (isMobile) {
    smoothCameraFly(0, 18, -24, 0, 1.8, 4.8);
  } else {
    smoothCameraFly(0, 20, -18, 0, 2, 4.8);
  }
}

function focusOnBlock(blockId) {
  const grp = blockGroups[blockId];
  if (!grp) return;
  const targetX = grp.position.x;
  const targetZ = grp.position.z;
  smoothCameraFly(targetX + 8.5, 4.8, targetZ + 1.5, targetX, 1.8, targetZ);
}

function smoothCameraFly(camX, camY, camZ, targetX, targetY, targetZ) {
  const startPos = camera.position.clone();
  const endPos = new THREE.Vector3(camX, camY, camZ);
  const startTarget = controls.target.clone();
  const endTarget = new THREE.Vector3(targetX, targetY, targetZ);

  let progress = 0;
  const speed = 0.04;

  function stepFly() {
    progress += speed;
    if (progress <= 1) {
      camera.position.lerpVectors(startPos, endPos, easeOutCubic(progress));
      controls.target.lerpVectors(startTarget, endTarget, easeOutCubic(progress));
      requestAnimationFrame(stepFly);
    } else {
      camera.position.copy(endPos);
      controls.target.copy(endTarget);
    }
  }
  stepFly();
}

function easeOutCubic(x) {
  return 1 - Math.pow(1 - x, 3);
}

// ==============================================================================
// 6. EXCEL (.XLSX) DYNAMIC PARSER & DROPZONE
// ==============================================================================
function setupExcelDropzone() {
  const dropzone = document.getElementById('excel-dropzone');
  const fileInput = document.getElementById('excel-file-input');

  if (!dropzone || !fileInput) return;

  dropzone.addEventListener('click', () => fileInput.click());

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = '#38bdf8';
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.style.borderColor = 'rgba(56, 189, 248, 0.4)';
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.style.borderColor = 'rgba(56, 189, 248, 0.4)';
    if (e.dataTransfer.files.length > 0) {
      handleExcelFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleExcelFile(e.target.files[0]);
    }
  });
}

function handleExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      // Parse Sheet into warehouse structure
      const parsedData = parseWorkbookToWarehouseData(workbook);
      window.WAREHOUSE_DATA = parsedData;
      buildWarehouse3D(parsedData);
      
      document.getElementById('upload-modal').classList.remove('show');
      alert(`Berhasil memuat data dari ${file.name}! Visualisasi 3D telah diperbarui.`);
    } catch (err) {
      console.error('Error parsing Excel', err);
      alert('Gagal memproses file Excel. Pastikan format file sesuai.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseWorkbookToWarehouseData(wb) {
  // If master sheet is present, build master dict
  let master = {};
  wb.SheetNames.forEach(sname => {
    const ws = wb.Sheets[sname];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    rows.forEach(r => {
      if (r && r.length >= 5) {
        const noGud = String(r[1] || r[2] || '').trim();
        const barkot = String(r[2] || r[3] || '').trim();
        const grade = String(r[3] || r[4] || '').trim();
        const kg = r[4] || r[5] || '';
        if (noGud && noGud !== 'No Gud' && noGud !== 'undefined') {
          master[noGud] = { no_gud: noGud, barkot: barkot, grade: grade, kg: kg, status: 'NORMAL' };
        }
      }
    });
  });

  // Preserve existing blocks structure if only master is updated, or fallback to current WAREHOUSE_DATA
  return {
    blocks: window.WAREHOUSE_DATA ? window.WAREHOUSE_DATA.blocks : {},
    master: Object.keys(master).length > 0 ? master : (window.WAREHOUSE_DATA ? window.WAREHOUSE_DATA.master : {})
  };
}

// ==============================================================================
// 7. RENDER LOOP
// ==============================================================================
function onWindowResize() {
  camera.aspect = window.innerWidth / window.innerHeight;
  camera.updateProjectionMatrix();
  renderer.setSize(window.innerWidth, window.innerHeight);
}

function animate() {
  requestAnimationFrame(animate);
  controls.update();
  renderer.render(scene, camera);
}
