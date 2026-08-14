/* Temper 前端:加载流派运行时模块,渲染输入表单,实时求值。
   计算全部发生在浏览器内(dist/flows/*.mjs 由 codegen 管线生成)。 */

const $ = (sel) => document.querySelector(sel);

const state = {
  index: null,        // flows/index.json
  constants: null,    // data/constants.json(心法候选)
  manifest: null,     // data/upstream-manifest.json(抓取时间)
  mod: null,          // 当前流派模块
  flowName: null,
  raf: 0,
};

const BLOCK_ORDER = ["面板", "标量", "怪物/杂项", "团队增益", "开关", "心法", "属性/套装", "武器"];

async function boot() {
  const [index, constants, manifest] = await Promise.all([
    fetch("flows/index.json").then((r) => r.json()),
    fetch("data/constants.json").then((r) => r.json()).catch(() => null),
    fetch("data/upstream-manifest.json").then((r) => r.json()).catch(() => null),
  ]);
  state.index = index;
  state.constants = constants;
  state.manifest = manifest;

  const sel = $("#flow-select");
  sel.innerHTML = "";
  for (const f of index.flows) {
    const opt = document.createElement("option");
    opt.value = f.flow;
    opt.textContent = `${f.flow} v${f.calcVersion}`;
    sel.appendChild(opt);
  }
  sel.disabled = false;
  sel.addEventListener("change", () => loadFlow(sel.value));
  $("#reset-btn").addEventListener("click", () => {
    if (!state.mod) return;
    state.mod.reset();
    renderForm();
    recompute();
  });
  await loadFlow(sel.value);
}

const modCache = new Map();

async function loadFlow(name) {
  const entry = state.index.flows.find((f) => f.flow === name);
  if (!modCache.has(name)) {
    modCache.set(name, await import(`./flows/${encodeURIComponent(entry.module)}`));
  }
  state.mod = modCache.get(name);
  state.flowName = name;
  state.mod.reset();
  renderForm();
  renderMeta(entry);
  recompute();
}

function xinfaCandidates(fieldName) {
  const flows = state.constants?.flows || [];
  const flow = flows.find((f) => f.name === state.flowName);
  const hit = (flow?.xinfaInputs || []).find((x) => x.label === fieldName);
  return hit?.candidates?.length ? hit.candidates : null;
}

function renderForm() {
  const fields = state.mod.fields();
  const wrap = $("#inputs");
  wrap.innerHTML = "";

  const byBlock = new Map();
  for (const [name, f] of Object.entries(fields)) {
    const block = f.kind === "text" && name.endsWith("心法") ? "心法" : (fieldBlock(name, f) || "其他");
    if (!byBlock.has(block)) byBlock.set(block, []);
    byBlock.get(block).push([name, f]);
  }

  const blocks = [...byBlock.keys()].sort(
    (a, b) => idx(BLOCK_ORDER, a) - idx(BLOCK_ORDER, b)
  );
  for (const block of blocks) {
    const sec = document.createElement("div");
    sec.className = "block";
    const h = document.createElement("h2");
    h.textContent = block;
    sec.appendChild(h);
    const grid = document.createElement("div");
    grid.className = "grid";
    for (const [name, f] of byBlock.get(block)) {
      grid.appendChild(fieldControl(name, f));
    }
    sec.appendChild(grid);
    wrap.appendChild(sec);
  }
}

// 字段块信息不在 fields() 里,按名字前缀/形态归组(与 flow-inputs 的 block 对齐)
function fieldBlock(name, f) {
  if (name.startsWith("面板.")) return "面板";
  if (name.startsWith("团队增益.")) return "团队增益";
  if (name.endsWith("心法")) return "心法";
  if (name.startsWith("武器")) return "武器";
  if (name === "武学属性" || name === "套装") return "属性/套装";
  if (["副本天赋", "战斗时间", "机制时间"].includes(name)) return "怪物/杂项";
  if (f.kind === "toggle" || name === "高伤取值") return "开关";
  return "标量";
}

const idx = (arr, v) => { const i = arr.indexOf(v); return i < 0 ? arr.length : i; };

function fieldControl(name, f) {
  const row = document.createElement("div");
  row.className = "field";
  const label = document.createElement("label");
  label.textContent = name.replace(/^面板\.|^团队增益\./, "");
  label.title = f.note || name;
  row.appendChild(label);

  let ctl;
  // 动态表单 + Chrome 的按位置状态恢复会造成「表单显示恢复值、模块
  // 仍是默认值」的错位;autocomplete=off 让 Chromium 跳过状态保存
  if (f.kind === "toggle") {
    ctl = document.createElement("input");
    ctl.type = "checkbox";
    ctl.checked = f.default === "√";
    ctl.addEventListener("change", () => {
      state.mod.set(name, ctl.checked ? "√" : "×");
      schedule();
    });
  } else if (f.kind === "number") {
    ctl = document.createElement("input");
    ctl.type = "number";
    ctl.step = "any";
    ctl.value = f.default ?? "";
    ctl.addEventListener("input", () => {
      const v = parseFloat(ctl.value);
      if (Number.isFinite(v)) {
        state.mod.set(name, v);
        schedule();
      }
    });
  } else {
    const candidates = xinfaCandidates(name);
    if (candidates) {
      ctl = document.createElement("select");
      for (const c of candidates) {
        const o = document.createElement("option");
        o.value = o.textContent = c;
        ctl.appendChild(o);
      }
      if (f.default != null) ctl.value = f.default;
      ctl.addEventListener("change", () => {
        state.mod.set(name, ctl.value);
        schedule();
      });
    } else {
      ctl = document.createElement("input");
      ctl.type = "text";
      ctl.value = f.default ?? "";
      if (f.note) ctl.placeholder = "见悬停说明";
      ctl.addEventListener("change", () => {
        state.mod.set(name, ctl.value);
        schedule();
      });
    }
  }
  ctl.setAttribute("autocomplete", "off");
  row.appendChild(ctl);
  return row;
}

function schedule() {
  // rAF 在后台/隐藏标签页会被节流甚至冻结,退化为 setTimeout 兜底
  clearTimeout(state.raf);
  state.raf = setTimeout(recompute, 30);
}

function recompute() {
  if (!state.mod) return;
  const out = state.mod.evaluate();
  setOut("毕业率", out["毕业率"], (v) => `${(v * 100).toFixed(2)}%`);
  setOut("ADPS", out["ADPS"], fmtInt);
  setOut("RDPS", out["RDPS"], fmtInt);
  setOut("总伤", out["总伤"], fmtInt);
  setOut("真气比例", out["真气比例"], (v) => `${(v * 100).toFixed(1)}%`);
}

const fmtInt = (v) => Math.round(v).toLocaleString("zh-CN");

function setOut(key, v, fmt) {
  const el = document.getElementById(`out-${key}`);
  el.textContent = typeof v === "number" && Number.isFinite(v) ? fmt(v) : "—";
}

function renderMeta(entry) {
  $("#meta-version").textContent = `${entry.flow} v${entry.calcVersion}`;
  $("#meta-synced").textContent =
    state.manifest?._meta?.syncedAt?.slice(0, 10) ?? "未知";
  $("#meta-generated").textContent =
    (state.index.generatedAt ?? "").slice(0, 10) || "未知";
  $("#meta-box").hidden = false;
}

window.__temper = state; // 调试句柄(控制台可用,不影响功能)

boot().catch((e) => {
  console.error(e);
  $("#inputs").innerHTML =
    `<p class="hint">加载失败:${e.message}。请刷新重试;若持续出现,请提 issue。</p>`;
});
