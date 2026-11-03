"""
ASCII 海洋背景动画（移植自 stackedlol/ascii-sharks 的视觉语言）。

原项目是 TypeScript + canvas 的完整人工生命模拟（boids 鱼群 + 程序化鲨鱼光栅化）。
这里保留它的核心观感，做了三处简化以适应「报告生成工具的背景层」这个用途：

1. 鲨鱼用脊线 + 半宽函数逐格采样（同原版思路），但节点数 10 而非 20、帧率降到 30fps
2. 鱼群是精简版 Reynolds boids（分离/对齐/聚合 + 躲避鲨鱼），不带 bait ball 环绕
3. 不做「捕食」事件与重生，鱼被冲散后会自己聚回来

字符密度梯度 `:;t%SX8#@`、青色 wake `#6ee6ff`、格子比例 13×19 都与原版一致。
只在暗色主题下显示（由 data-ui-theme 控制透明度），并遵守 prefers-reduced-motion。
"""

OCEAN_JS = r"""
function(){
  if (window.__oceanInit) return [];
  window.__oceanInit = true;

  const CELL_W = 13, CELL_H = 19;
  const DENSITY = ":;t%SX8#@";            // 亮 → 暗的字符梯度（同原版 SKIN_RAMP）
  const WAKE = "#6ee6ff";                  // 青色尾迹
  const FISH_CH = "-";
  const MAX_FISH = 150;
  const SHARK_LEN = [300, 240, 190];       // 三条不同体长的鲨鱼

  let cv = document.createElement("canvas");
  cv.id = "asciiOcean";
  cv.setAttribute("aria-hidden", "true");
  cv.style.cssText =
    "position:fixed;inset:0;width:100%;height:100%;z-index:0;" +
    "pointer-events:none;opacity:0;transition:opacity .35s ease;";
  // 必须是 body 的第一个子元素，内容层才能盖在它上面
  document.body.insertBefore(cv, document.body.firstChild);

  const ctx = cv.getContext("2d", { alpha: true });
  let W = 0, H = 0, dpr = 1, cols = 0, rows = 0;
  let running = false, raf = 0, last = 0;
  const fish = [];
  const sharks = [];

  const rnd = (a, b) => a + Math.random() * (b - a);
  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);

  // ── 初始化 ────────────────────────────────────────────────
  function resize() {
    dpr = Math.min(window.devicePixelRatio || 1, 2);
    W = window.innerWidth;
    H = window.innerHeight;
    cv.width = Math.round(W * dpr);
    cv.height = Math.round(H * dpr);
    cols = Math.ceil(W / CELL_W);
    rows = Math.ceil(H / CELL_H);
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  function seed() {
    fish.length = 0;
    const n = clamp(Math.round((W * H) / 12000), 40, MAX_FISH);
    for (let i = 0; i < n; i++) {
      const dir = (Math.random() < 0.5 ? 0 : Math.PI) + rnd(-0.4, 0.4);
      fish.push({
        x: rnd(0, W), y: rnd(0, H),
        vx: Math.cos(dir) * 34, vy: Math.sin(dir) * 34,
        alarm: 0,
      });
    }
    sharks.length = 0;
    SHARK_LEN.forEach((len, i) => {
      const heading = rnd(0, Math.PI * 2);
      const x = rnd(W * 0.2, W * 0.8), y = rnd(H * 0.2, H * 0.8);
      const nodes = [];
      for (let k = 0; k < 10; k++) {
        nodes.push([x - Math.cos(heading) * (len / 9) * k,
                    y - Math.sin(heading) * (len / 9) * k]);
      }
      sharks.push({
        len, heading, speed: 26 + i * 6, phase: rnd(0, 6.28),
        nodes, x, y,
      });
    });
  }

  // ── 模拟 ──────────────────────────────────────────────────
  function stepFish(dt) {
    const R = 62, R2 = R * R;
    for (const f of fish) {
      let ax = 0, ay = 0, sx = 0, sy = 0, avx = 0, avy = 0, cx = 0, cy = 0, n = 0;
      for (const o of fish) {
        if (o === f) continue;
        const dx = o.x - f.x, dy = o.y - f.y;
        const d2 = dx * dx + dy * dy;
        if (d2 > R2 || d2 < 1e-6) continue;
        const d = Math.sqrt(d2);
        if (d < 26) { const k = (26 - d) / d * 0.9; sx -= dx * k; sy -= dy * k; }
        avx += o.vx; avy += o.vy; cx += dx; cy += dy; n++;
      }
      ax += sx * 320; ay += sy * 320;
      if (n > 0) {
        ax += (avx / n - f.vx) * 1.5 + (cx / n) * 0.9;
        ay += (avy / n - f.vy) * 1.5 + (cy / n) * 0.9;
      }
      // 躲鲨鱼
      for (const sh of sharks) {
        const fearR = sh.len * 0.55 + 50;
        const d = Math.hypot(f.x - sh.nodes[0][0], f.y - sh.nodes[0][1]);
        if (d < fearR && d > 1e-3) {
          const w = (1 - d / fearR) ** 2;
          ax += ((f.x - sh.nodes[0][0]) / d) * 900 * w;
          ay += ((f.y - sh.nodes[0][1]) / d) * 900 * w;
          f.alarm = Math.max(f.alarm, w);
        }
      }
      f.alarm *= Math.exp(-dt * 0.8);
      // 软边界
      const m = 40;
      if (f.x < m) ax += (m - f.x) * 5;
      else if (f.x > W - m) ax -= (f.x - (W - m)) * 5;
      if (f.y < m) ay += (m - f.y) * 5;
      else if (f.y > H - m) ay -= (f.y - (H - m)) * 5;
      ay -= f.vy * 2.2;                       // 保持横游
      const sp = Math.hypot(f.vx, f.vy) || 1;
      const want = 30 * (1 + f.alarm);
      ax += (want - sp) * (f.vx / sp) * 1.5;
      ay += (want - sp) * (f.vy / sp) * 1.5;
      f.vx += ax * dt; f.vy += ay * dt;
      const s2 = Math.hypot(f.vx, f.vy);
      const cap = 46 * (1 + f.alarm);
      if (s2 > cap) { f.vx *= cap / s2; f.vy *= cap / s2; }
      f.x = clamp(f.x + f.vx * dt, -40, W + 40);
      f.y = clamp(f.y + f.vy * dt, -40, H + 40);
    }
  }

  function stepSharks(dt, t) {
    for (const sh of sharks) {
      // 缓慢漫游 + 绕中心巡游
      const dx = W / 2 - sh.x, dy = H / 2 - sh.y;
      const d = Math.hypot(dx, dy) || 1;
      const desired = Math.atan2(dy, dx) + 1.4 + Math.sin(t * 0.3 + sh.phase) * 0.5;
      let diff = desired - sh.heading;
      while (diff > Math.PI) diff -= Math.PI * 2;
      while (diff < -Math.PI) diff += Math.PI * 2;
      sh.heading += clamp(diff, -0.9 * dt, 0.9 * dt);
      // 出界就往回游
      if (sh.x < -sh.len || sh.x > W + sh.len || sh.y < -sh.len || sh.y > H + sh.len) {
        const back = Math.atan2(H / 2 - sh.y, W / 2 - sh.x);
        let bd = back - sh.heading;
        while (bd > Math.PI) bd -= Math.PI * 2;
        while (bd < -Math.PI) bd += Math.PI * 2;
        sh.heading += clamp(bd, -2.4 * dt, 2.4 * dt);
      }
      sh.phase += dt * 2.2;
      sh.x += Math.cos(sh.heading) * sh.speed * dt;
      sh.y += Math.sin(sh.heading) * sh.speed * dt;
      sh.nodes[0][0] = sh.x;
      sh.nodes[0][1] = sh.y;
      // follow-the-leader 脊线，带身体波动
      const seg = sh.len / 9;
      let prev = sh.heading + Math.PI;
      for (let k = 1; k < 10; k++) {
        const px = sh.nodes[k - 1][0], py = sh.nodes[k - 1][1];
        const amp = sh.len * 0.045 * (0.1 + 0.9 * Math.pow(k / 9, 1.8));
        const off = amp * Math.sin(sh.phase - (k / 9) * 5.2);
        const a = prev + Math.sin(sh.phase - k * 0.5) * 0.06;
        sh.nodes[k][0] = px + Math.cos(a) * seg - Math.sin(a) * (off * 0.35);
        sh.nodes[k][1] = py + Math.sin(a) * seg + Math.cos(a) * (off * 0.35);
        prev = a;
      }
    }
  }

  // ── 渲染 ──────────────────────────────────────────────────
  const halfWidth = (s) => {
    // 经典鲨鱼俯视轮廓：吻端 → 胸鳍区 → 尾柄
    if (s < 0.12) return 0.13 * Math.pow(Math.sin((s / 0.12) * Math.PI / 2), 0.85);
    if (s < 0.28) return 0.13 + 0.035 * (s - 0.12) / 0.16;
    if (s < 0.72) return 0.055 + 0.110 * Math.pow(1 - (s - 0.28) / 0.44, 1.15);
    const k = (s - 0.72) / 0.28;
    return 0.055 * (1 - k * 0.65) + 0.020 * Math.sin(k * Math.PI);
  };

  function sampleShark(px, py, sh) {
    let best = Infinity, bestS = 0, bestSign = 1;
    for (let k = 0; k < 9; k++) {
      const ax = sh.nodes[k][0], ay = sh.nodes[k][1];
      const bx = sh.nodes[k + 1][0], by = sh.nodes[k + 1][1];
      const sx = bx - ax, sy = by - ay;
      const len2 = sx * sx + sy * sy || 1;
      let u = ((px - ax) * sx + (py - ay) * sy) / len2;
      u = u < 0 ? 0 : u > 1 ? 1 : u;
      const qx = ax + sx * u, qy = ay + sy * u;
      const d2 = (px - qx) ** 2 + (py - qy) ** 2;
      if (d2 < best) {
        best = d2;
        bestS = (k + u) / 9;
        bestSign = sx * (py - ay) - sy * (px - ax) >= 0 ? 1 : -1;
      }
    }
    const a = Math.sqrt(best);
    const w = halfWidth(bestS) * sh.len;
    return a <= w ? { s: bestS, across: a / (w || 1), v: a * bestSign } : null;
  }
  function draw(t) {
    ctx.clearRect(0, 0, W, H);
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    // 鱼：暗蓝灰短横线，受惊时变亮
    ctx.font = "600 13px ui-monospace, Menlo, Consolas, monospace";
    for (const f of fish) {
      const c = f.alarm > 0.3 ? 0.85 : 0.42;
      ctx.fillStyle = "rgba(190,212,255," + c.toFixed(2) + ")";
      const cx = Math.round(f.x / CELL_W) * CELL_W + CELL_W / 2;
      const cy = Math.round(f.y / CELL_H) * CELL_H + CELL_H / 2;
      ctx.fillText(FISH_CH, cx, cy + 1);
    }

    // 鲨鱼：逐格采样，用密度字符拼出形体
    ctx.font = "600 14px ui-monospace, Menlo, Consolas, monospace";
    for (const sh of sharks) {
      const pad = sh.len * 0.4;
      const c0 = Math.max(0, Math.floor((sh.x - pad) / CELL_W));
      const c1 = Math.min(cols - 1, Math.floor((sh.x + pad) / CELL_W));
      const r0 = Math.max(0, Math.floor((sh.y - pad) / CELL_H));
      const r1 = Math.min(rows - 1, Math.floor((sh.y + pad) / CELL_H));
      for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
          const x = c * CELL_W + CELL_W / 2;
          const y = r * CELL_H + CELL_H / 2;
          const hit = sampleShark(x, y, sh);
          if (!hit) continue;
          // 受光：靠背亮、靠腹暗（across 0 中线 → 1 边缘）
          const shade = 0.62 - hit.across * hit.across * 0.45;
          const idx = clamp(Math.floor((1 - shade) * (DENSITY.length - 0.01)), 0, DENSITY.length - 1);
          let ch = DENSITY[idx];
          let color;
          if (hit.s < 0.06 && hit.across < 0.5) { ch = "o"; color = "#b9c0cc"; }
          else if (hit.across > 0.86) { ch = hit.s > 0.5 ? "%" : ":"; color = "#79839a"; }
          else {
            // 少数格子闪青色，模拟原版的 shimmer
            const flick = ((c * 7 + r * 13 + Math.floor(t * 2.4)) % 97) < 2;
            const g = 46 + shade * 160;
            color = flick ? WAKE : "rgb(" +
              Math.round(g * 0.95) + "," +
              Math.round(g * 1.02) + "," +
              Math.min(255, Math.round(g * 1.14 + 8)) + ")";
          }
          ctx.fillStyle = color;
          ctx.fillText(ch, x, y + 1);
        }
      }
    }
  }

  // ── 主循环 ─────────────────────────────────────────────────
  const STEP = 1 / 30;
  let acc = 0;
  const reduce = window.matchMedia("(prefers-reduced-motion: reduce)");

  function frame(now) {
    if (!running) return;
    const dt = Math.min(0.1, (now - last) / 1000);
    last = now;
    acc += dt;
    let guard = 0;
    while (acc >= STEP && guard < 3) {
      const t = now / 1000;
      stepFish(STEP);
      stepSharks(STEP, t);
      acc -= STEP;
      guard++;
    }
    if (guard === 3) acc = 0;
    draw(now / 1000);
    raf = requestAnimationFrame(frame);
  }

  function start() {
    if (running || reduce.matches) return;
    running = true;
    last = performance.now();
    raf = requestAnimationFrame(frame);
  }
  function stop() {
    running = false;
    cancelAnimationFrame(raf);
  }

  // 亮暗主题联动：只有暗色才显示这片海
  function syncTheme() {
    const dark = document.documentElement.getAttribute("data-ui-theme") === "dark";
    cv.style.opacity = dark ? "1" : "0";
    if (dark && !document.hidden) start(); else stop();
  }

  resize();
  seed();
  syncTheme();

  new MutationObserver(syncTheme).observe(document.documentElement, {
    attributes: true, attributeFilter: ["data-ui-theme"],
  });
  window.addEventListener("resize", () => { resize(); seed(); });
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) stop(); else syncTheme();
  });
  reduce.addEventListener("change", syncTheme);

  // 内容层必须盖在 canvas 之上
  const gc = document.querySelector(".gradio-container");
  if (gc) { gc.style.position = "relative"; gc.style.zIndex = "1"; }

  return [];
}
"""
