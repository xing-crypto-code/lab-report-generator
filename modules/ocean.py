"""
ASCII 海洋背景 —— 采用 stackedlol/ascii-sharks 仓库 README 的手绘鲨鱼字符画。

关于路线的说明（重要）：
原仓库引擎 renderer.ts 是**程序化光栅化**（脊线采样 + 超采样 + 逐格 fillRect/fillText）。
我严格照做了一版（见 git 历史），但在「10~12px 格子 + 半透明面板遮挡」这个场景下，
程序化剪影会退化成模糊色团，辨识不出是鲨鱼。

因此这里改用仓库 **README 里那幅手绘 ASCII 鲨鱼**作为精灵：
- 形状是手绘的，辨识度远高于程序化采样
- 每帧只画几十个字符，性能比逐格采样高一个量级
- 字符本身带密度语义（@#%8 深、;:. 浅），延续原版 SKIN_RAMP 的观感

保留的原版要素：纯黑底 + 可见字符网格、青色尾迹（wake）、鱼群 `-`、密度梯度配色。
亮暗两套配色，两种主题都可见。
"""

OCEAN_JS = r"""
function(){
  if (window.__oceanInit) return [];
  window.__oceanInit = true;

  // ── 仓库 README 里的鲨鱼字符画（原样保留）──────────────────
  const SHARK_ART = [
    "                            ;'",
    "                             %%'",
    "                              .t;.               .oooo",
    "                               S;:S%              %8S8.",
    "                                ;:tt%t '..        XX%%S",
    "t'                            ''XSSS8SSX%XX8#.    %XSX%",
    " :;'               ;.''X8%8XX%%S%%S%SStt%tSX8S#X8't%ttt.",
    "  ::%'       ''8XXXSSt%Stttttt%8@%%@%8@8%Sttt%%%%SSSSS%tt.",
    "  'SXXtt%t%%%S%:;StS%ttSS%tSXSt88@@%8%@8%tStS%%X%SS%%S%;;t",
    " .@##88#@...'@@@@#@8@88X88#XXXSX8SX88SX8%SS%S%%S88St%t%%%.",
    "  %;%t             :.    '.'##@###@X#888XX##8@@@#.X%XS%.",
    " .%t                            ;%%%#@@##@@@.     SXS%%",
    " t.                            .%;%%%.            %%XXS",
    "                              '%t:t              .SXSS'",
    "                             '::.                ooooo",
    "                            .S'",
    "                            .",
  ];
  const ART_W = 58, ART_H = SHARK_ART.length;

  // 字符 → 亮度档（0 暗 → 1 亮），对应原版 SKIN_RAMP 的密度语义
  const CHAR_LUM = {
    "@": 1.00, "#": 0.96, "%": 0.90, "8": 0.84, "X": 0.76,
    "S": 0.66, "t": 0.54, ";": 0.42, ":": 0.34, "'": 0.26,
    ".": 0.20, ",": 0.22, "`": 0.20, "o": 0.30,
  };

  const CELL_W = 9, CELL_H = 15;      // 字符网格（接近原版 1:1.45）
  const MAX_FISH = 110;
  const WAKE_LEVELS = 4;

  function palette() {
    return isDark
      ? {
          art: (l) => "rgba(" + Math.round(52 + l * 118) + "," +
                       Math.round(60 + l * 126) + "," +
                       Math.round(74 + l * 138) + ",0.42)",
          fishRGB: "198,218,255", fishA0: 0.34, fishA1: 0.82,
          wakeRGB: "110,230,255",
        }
      : {
          art: (l) => "rgba(" + Math.round(128 - l * 78) + "," +
                       Math.round(136 - l * 80) + "," +
                       Math.round(150 - l * 82) + ",0.46)",
          fishRGB: "58,72,98", fishA0: 0.5, fishA1: 0.9,
          wakeRGB: "20,108,146",
        };
  }

  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const rnd = (a, b) => a + Math.random() * (b - a);

  // ── 画布 ────────────────────────────────────────────────────
  const cv = document.createElement("canvas");
  cv.id = "asciiOcean";
  cv.setAttribute("aria-hidden", "true");
  cv.style.cssText =
    "position:fixed;inset:0;width:100%;height:100%;z-index:0;pointer-events:none;";
  document.body.insertBefore(cv, document.body.firstChild);
  const ctx = cv.getContext("2d", { alpha: true });

  let W = 0, H = 0, dpr = 1, cols = 0, rows = 0;
  let running = false, raf = 0, last = 0, isDark = true;
  let frames = 0, fpsT = performance.now();
  window.__oceanFps = 0;
  const fish = [], sharks = [];
  let energy = new Float32Array(0), energyTmp = new Float32Array(0);

  function resize() {
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = window.innerWidth; H = window.innerHeight;
    cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr);
    cols = Math.ceil(W / CELL_W); rows = Math.ceil(H / CELL_H);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    energy = new Float32Array(cols * rows);
    energyTmp = new Float32Array(cols * rows);
  }

  function seed() {
    fish.length = 0;
    const n = clamp(Math.round((W * H) / 13000), 40, MAX_FISH);
    for (let i = 0; i < n; i++) {
      const dir = (Math.random() < 0.5 ? 0 : Math.PI) + rnd(-0.35, 0.35);
      fish.push({ x: rnd(0, W), y: rnd(0, H),
                  vx: Math.cos(dir) * 32, vy: Math.sin(dir) * 32, alarm: 0 });
    }
    // 三条鲨鱼：一大两小，横向巡游（字符画是侧视，横着游最自然）
    sharks.length = 0;
    const baseScale = clamp(Math.min(W, H) / 900, 0.6, 1.15);
    const specs = [
      { s: 1.00, y: 0.30, dir: 1 },
      { s: 0.5, y: 0.82, dir: -1 },
    ];
    for (const sp of specs) {
      sharks.push({
        scale: baseScale * sp.s,
        x: rnd(0, W), y: H * sp.y, baseY: H * sp.y,
        dir: sp.dir,
        speed: 16 + sp.s * 18,
        bob: rnd(0, 6.28),
        bobAmp: 5 + sp.s * 10,
        flip: sp.dir < 0,
      });
    }
  }

  // ── 鱼群 ────────────────────────────────────────────────────
  function stepFish(dt) {
    const R = 62, R2 = R * R;
    for (const f of fish) {
      let ax = 0, ay = 0, sx = 0, sy = 0, avx = 0, avy = 0, cx = 0, cy = 0, n = 0;
      for (const o of fish) {
        if (o === f) continue;
        const dx = o.x - f.x, dy = o.y - f.y, d2 = dx * dx + dy * dy;
        if (d2 > R2 || d2 < 1e-6) continue;
        const d = Math.sqrt(d2);
        if (d < 22) { const k = (22 - d) / d * 0.9; sx -= dx * k; sy -= dy * k; }
        avx += o.vx; avy += o.vy; cx += dx; cy += dy; n++;
      }
      ax += sx * 320; ay += sy * 320;
      if (n > 0) {
        ax += (avx / n - f.vx) * 1.5 + (cx / n) * 0.9;
        ay += (avy / n - f.vy) * 1.5 + (cy / n) * 0.9;
      }
      // 鲨鱼靠近时惊散
      for (const sh of sharks) {
        const artW = ART_W * CELL_W * sh.scale;
        const artH = ART_H * CELL_H * sh.scale;
        const fearR = artW * 0.3 + 36;
        const ddx = f.x - sh.x, ddy = f.y - sh.y;
        const d = Math.hypot(ddx, ddy);
        if (d < fearR && d > 1e-3 && Math.abs(ddy) < artH * 0.7) {
          const w = (1 - d / fearR) ** 2;
          ax += (ddx / d) * 900 * w;
          ay += (ddy / d) * 900 * w;
          f.alarm = Math.max(f.alarm, w);
        }
      }
      f.alarm *= Math.exp(-dt * 0.8);
      const m = 40;
      if (f.x < m) ax += (m - f.x) * 5; else if (f.x > W - m) ax -= (f.x - (W - m)) * 5;
      if (f.y < m) ay += (m - f.y) * 5; else if (f.y > H - m) ay -= (f.y - (H - m)) * 5;
      ay -= f.vy * 2.2;
      const sp = Math.hypot(f.vx, f.vy) || 1;
      const want = 28 * (1 + f.alarm);
      ax += (want - sp) * (f.vx / sp) * 1.5;
      ay += (want - sp) * (f.vy / sp) * 1.5;
      f.vx += ax * dt; f.vy += ay * dt;
      const s2 = Math.hypot(f.vx, f.vy), cap = 44 * (1 + f.alarm);
      if (s2 > cap) { f.vx *= cap / s2; f.vy *= cap / s2; }
      f.x = clamp(f.x + f.vx * dt, -40, W + 40);
      f.y = clamp(f.y + f.vy * dt, -40, H + 40);
    }
  }

  function stepSharks(dt, t) {
    for (const sh of sharks) {
      const artW = ART_W * CELL_W * sh.scale;
      sh.x += sh.dir * sh.speed * dt;
      if (sh.dir > 0 && sh.x > W + artW * 0.8) sh.x = -artW;
      if (sh.dir < 0 && sh.x < -artW) sh.x = W + artW * 0.8;
      sh.y = sh.baseY + Math.sin(t * 0.35 + sh.bob) * sh.bobAmp;
    }
  }

  // ── 尾迹 ────────────────────────────────────────────────────
  function cellIndex(x, y) {
    const c = Math.floor(x / CELL_W), r = Math.floor(y / CELL_H);
    if (c < 0 || r < 0 || c >= cols || r >= rows) return -1;
    return r * cols + c;
  }
  function depositWake() {
    for (const f of fish) {
      const i = cellIndex(f.x, f.y);
      if (i >= 0) energy[i] = Math.min(1.2, energy[i] + 0.012 + f.alarm * 0.03);
    }
    for (const sh of sharks) {
      const artW = ART_W * CELL_W * sh.scale;
      const artH = ART_H * CELL_H * sh.scale;
      // 尾部（游动方向的反侧）扫过的水最多
      const tailX = sh.x + (sh.dir > 0 ? artW * 0.1 : artW * 0.9);
      for (let k = 0; k < 4; k++) {
        const i = cellIndex(tailX, sh.y + (k - 1.5) * artH * 0.22);
        if (i >= 0) energy[i] = Math.min(1.2, energy[i] + 0.05);
      }
    }
  }
  function diffuse(dt) {
    const decay = Math.exp(-dt * 1.7);
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const i = r * cols + c;
        const l = c > 0 ? energy[i - 1] : energy[i];
        const rt = c < cols - 1 ? energy[i + 1] : energy[i];
        const u = r > 0 ? energy[i - cols] : energy[i];
        const d = r < rows - 1 ? energy[i + cols] : energy[i];
        energyTmp[i] = (energy[i] * 4 + l + rt + u + d) * 0.125 * decay;
      }
    }
    const t = energy; energy = energyTmp; energyTmp = t;
  }

  // ── 渲染 ────────────────────────────────────────────────────
  function render(t) {
    const P = palette();
    ctx.clearRect(0, 0, W, H);

    // 1) 尾迹点（分桶，减少 fillStyle 切换）
    const buckets = [[], [], [], []];
    for (let i = 0; i < energy.length; i++) {
      const e = energy[i];
      if (e < 0.10) continue;
      buckets[clamp(Math.floor(e * WAKE_LEVELS), 0, WAKE_LEVELS - 1)].push(i);
    }
    const dot = 2;
    for (let b = 0; b < WAKE_LEVELS; b++) {
      const arr = buckets[b];
      if (!arr.length) continue;
      const wa = (0.10 + b * 0.15) * (isDark ? 1 : 0.45);
      ctx.fillStyle = "rgba(" + P.wakeRGB + "," + wa.toFixed(2) + ")";
      for (const i of arr) {
        const c = i % cols, r = (i / cols) | 0;
        ctx.fillRect(Math.round(c * CELL_W + CELL_W * 0.6 - dot / 2),
                     Math.round(r * CELL_H + CELL_H * 0.6 - dot / 2), dot, dot);
      }
    }

    // 2) 鱼群
    ctx.font = "600 11px ui-monospace, Menlo, Consolas, monospace";
    ctx.textAlign = "center"; ctx.textBaseline = "middle";
    for (const f of fish) {
      ctx.fillStyle = "rgba(" + P.fishRGB + "," +
        (f.alarm > 0.3 ? P.fishA1 : P.fishA0).toFixed(2) + ")";
      const cx = Math.round(f.x / CELL_W) * CELL_W + CELL_W / 2;
      const cy = Math.round(f.y / CELL_H) * CELL_H + CELL_H / 2;
      ctx.fillText("-", cx, cy);
    }

    // 3) 鲨鱼字符画
    ctx.textAlign = "left"; ctx.textBaseline = "top";
    for (const sh of sharks) {
      const sc = sh.scale;
      const cw = CELL_W * sc, chh = CELL_H * sc;
      ctx.font = "600 " + Math.round(chh * 1.05) + "px ui-monospace, Menlo, Consolas, monospace";
      for (let li = 0; li < SHARK_ART.length; li++) {
        const line = SHARK_ART[li];
        for (let ci = 0; ci < line.length; ci++) {
          const chr = line[ci];
          if (chr === " ") continue;
          const lum = CHAR_LUM[chr];
          if (lum === undefined) continue;
          const col = sh.flip ? (ART_W - 1 - ci) : ci;
          const px = sh.x + col * cw;
          const py = sh.y + li * chh;
          if (px < -cw * 2 || px > W + cw * 2) continue;
          ctx.fillStyle = P.art(lum);
          ctx.fillText(chr, px, py);
        }
      }
    }
  }

  // ── 主循环 ──────────────────────────────────────────────────
  const STEP = 1 / 30;
  let acc = 0;
  const reduce = window.matchMedia("(prefers-reduced-motion: reduce)");

  function frame(now) {
    if (!running) return;
    const dt = Math.min(0.1, (now - last) / 1000);
    last = now; acc += dt;
    let guard = 0;
    while (acc >= STEP && guard < 3) {
      stepFish(STEP);
      stepSharks(STEP, now / 1000);
      depositWake();
      diffuse(STEP);
      acc -= STEP; guard++;
    }
    if (guard === 3) acc = 0;
    render(now / 1000);
    frames++;
    if (now - fpsT > 1000) {
      window.__oceanFps = Math.round(frames * 1000 / (now - fpsT));
      frames = 0; fpsT = now;
    }
    raf = requestAnimationFrame(frame);
  }

  function start() {
    if (running) return;
    if (reduce.matches) { render(performance.now() / 1000); return; }
    running = true; last = performance.now();
    raf = requestAnimationFrame(frame);
  }
  function stop() { running = false; cancelAnimationFrame(raf); }

  function syncTheme() {
    isDark = document.documentElement.getAttribute("data-ui-theme") === "dark";
    if (document.hidden) { stop(); return; }
    start();
    if (reduce.matches) render(performance.now() / 1000);
  }

  resize(); seed(); syncTheme();

  new MutationObserver(syncTheme).observe(document.documentElement,
    { attributes: true, attributeFilter: ["data-ui-theme"] });
  window.addEventListener("resize", () => { resize(); seed(); });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) stop(); else syncTheme();
  });
  reduce.addEventListener("change", syncTheme);

  const gc = document.querySelector(".gradio-container");
  if (gc) { gc.style.position = "relative"; gc.style.zIndex = "1"; }

  return [];
}
"""
