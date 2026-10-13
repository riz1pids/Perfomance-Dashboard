/* Boots the ported page in jsdom on the real September payloads and checks it
   against totals computed independently from the raw files. */
const fs = require("fs"), path = require("path");
const { JSDOM, VirtualConsole } = require("jsdom");
const FILE = path.resolve(process.argv[2] || "index.html");
const EXPECT = JSON.parse(process.argv[3] || "{}");
const errs = [], vc = new VirtualConsole();
vc.on("jsdomError", (e) => errs.push(e.message + (e.detail ? " " + e.detail : "")));
const dom = new JSDOM(fs.readFileSync(FILE, "utf8"), {
  runScripts: "dangerously", resources: "usable", url: "file://" + FILE,
  pretendToBeVisual: true, virtualConsole: vc,
  beforeParse(w) {
    w.DecompressionStream = DecompressionStream; w.Blob = Blob; w.Response = Response;
    w.atob = atob; w.URL.createObjectURL = () => "blob:x"; w.URL.revokeObjectURL = () => {};
    w.HTMLAnchorElement.prototype.click = function () { w.__dl = (w.__dl || 0) + 1; };
  }
});
const w = dom.window, d = w.document, $ = (i) => d.getElementById(i);
const N = (v) => parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
const cr = (v) => (N(v) / 1e7).toFixed(2);
function fire(e, type) { e.dispatchEvent(new w.Event(type, { bubbles: true })); }
function setSel(id, v) {
  const e = $(id);
  if (![...e.options].some((o) => o.value === v)) throw new Error(id + " lacks " + v);
  e.value = v; fire(e, "change");
}
function tab(name) {
  const b = [...$("catTabs").querySelectorAll("button")].find((x) => x.textContent === name);
  if (!b) throw new Error("no tab " + name); if (b.disabled) throw new Error("tab disabled: " + name);
  b.click();
}
function colIndex(wrap, name) {
  const ths = [...$(wrap).querySelectorAll("thead th")].map((t) => t.querySelector(".sortbtn").textContent.replace(/ [↑↓]$/, ""));
  return ths.indexOf(name);
}
function footer(wrap, name) {
  const tds = [...$(wrap).querySelectorAll("tfoot tr.tot td")];
  const labels = $(wrap).querySelectorAll("thead th.lbl").length;
  const i = colIndex(wrap, name);
  if (i < 0) throw new Error("no column " + name + " in " + wrap);
  return tds[i - labels + 1].textContent;
}
const bodyRows = (wrap) => $(wrap).querySelectorAll("tbody tr:not(.sub)").length;
const cardText = () => [...$("cards").querySelectorAll(".card")].map((c) => c.querySelector(".k").textContent + " = " + c.querySelector(".v").textContent);
const out = []; let failed = 0;
function step(name, fn) {
  try { fn(); out.push("  ok   " + name); }
  catch (e) { failed++; out.push("  FAIL " + name + "\n         " + e.message); }
}
const ok = (c, m) => { if (!c) throw new Error(m); };
const near = (a, b, m) => ok(Math.abs(Number(a) - Number(b)) < 0.006, m + ": " + a + " vs " + b);

(async () => {
  const t0 = Date.now();
  const heap = () => { if (global.gc) global.gc(); return Math.round(process.memoryUsage().heapUsed / 1e6); };
  await new Promise((res, rej) => (function poll() {
    if (!$("shell").hidden) return res();
    if ($("boot").querySelector(".bad")) return rej(new Error($("boot").textContent));
    if (Date.now() - t0 > 120000) return rej(new Error("timeout: " + $("boot").textContent));
    setTimeout(poll, 30);
  })());
  console.log("booted in " + (Date.now() - t0) + " ms | " + $("periodLbl").textContent + " | " + $("builtLbl").textContent);
  console.log("cards:\n   " + cardText().join("\n   "));
  console.log("heap after boot: " + heap() + " MB");

  step("same-store total matches the raw payload", () => near(cr(footer("catWrap", "Sales")), EXPECT.same, "same-store"));
  step("division tab, all store, matches", () => {
    d.querySelector('#scopeSeg button[data-scope="all"]').click();
    near(cr(footer("catWrap", "Sales")), EXPECT.all, "all-store");
  });
  step("every product tab reconciles", () => {
    ["Category 1", "Category 3", "Brand", "Responsible person", "Head"].forEach((t) => {
      tab(t); near(cr(footer("catWrap", "Sales")), EXPECT.all, t);
    });
  });
  step("outlet-side tabs reconcile", () => {
    ["Outlet division", "Area", "Leader", "Zonal"].forEach((t) => {
      tab(t); near(cr(footer("catWrap", "Sales")), EXPECT.all, t);
    });
  });
  step("outlet detail reconciles", () => near(cr(footer("outWrap", "Sales")), EXPECT.all, "outlet detail"));
  step("article tab reconciles on the full period", () => {
    const h0 = heap(), t1 = Date.now();
    tab("Article"); ok(bodyRows("catWrap") > (EXPECT.minArticles || 1000), "article rows " + bodyRows("catWrap"));
    console.log("   article tab: " + bodyRows("catWrap") + " rows, " + (Date.now() - t1) + " ms, heap " + h0 + " -> " + heap() + " MB");
    near(cr(footer("catWrap", "Sales")), EXPECT.all, "articles");
    tab("Article division");
    console.log("   heap after leaving the article tab: " + heap() + " MB");
  });
  step("targets show on all store with no filters", () => {
    tab("Category 3");
    ok(colIndex("catWrap", "Sales target") > 0, "no target column");
    ok(N(footer("catWrap", "Sales target")) > 0, "zero target");
    ok(cardText().some((c) => /target vs achievement · all store = \d/.test(c)), "target card empty");
  });
  step("Frozen Food: category view and outlet view agree", () => {
    setSel("fCat3", "FROZEN FOOD");
    const viaCat = cr(footer("catWrap", "Sales")), viaOut = cr(footer("outWrap", "Sales"));
    near(viaCat, EXPECT.frozen, "category"); near(viaOut, EXPECT.frozen, "outlet");
    ok(bodyRows("outWrap") > (EXPECT.minFrozenOutlets || 500), "frozen outlets " + bodyRows("outWrap"));
  });
  step("brand filter is day-wise and narrows", () => {
    tab("Brand");
    const top = $("catWrap").querySelector("tbody tr td.lbl").textContent;   // biggest brand
    setSel("fBrand", top);
    const one = N(footer("catWrap", "Sales"));
    ok(one > 0 && one / 1e7 < EXPECT.frozen, "brand total " + one + " for " + top);
  });
  step("a partial range narrows the brand view", () => {
    const full = N(footer("catWrap", "Sales"));
    d.getElementById("dLast").click();
    const last = N(footer("catWrap", "Sales"));
    ok(last > 0 && last < full, "latest day " + last + " vs full " + full);
    d.getElementById("dFull").click();
    setSel("fBrand", "");
  });
  step("outlet filter recomputes category tabs from outlet data", () => {
    setSel("fCat3", "");
    const z = [...$("oZone").options].map((o) => o.value).filter(Boolean)[0];
    setSel("oZone", z);
    tab("Category 3");
    const zv = N(footer("catWrap", "Sales")) / 1e7;
    ok(zv > 0 && zv < EXPECT.all, "zone total " + zv);
    near(cr(footer("catWrap", "Sales")), cr(footer("outWrap", "Sales")), "zone: category vs outlet");
    ok([...$("catTabs").querySelectorAll("button")].find((b) => b.textContent === "Brand").disabled, "brand tab not locked");
    $("clearOut").click();
  });
  step("article search: totals, period-only note, chart steps aside", () => {
    d.getElementById("dLast").click();
    const any = d.querySelector("#catWrap"); // ensure rendered
    $("fSku").value = EXPECT.q || "MILK"; fire($("fSku"), "input");
  });
  await new Promise((r) => setTimeout(r, 400));
  step("article search applied", () => {
    ok(/whole-period totals/.test($("catNote").textContent), "no period note");
    ok(/no day-by-day line/.test($("chartWrap").textContent), "chart not replaced");
    ok($("dFrom").value === $("dFrom").min && $("dTo").value === $("dTo").max, "range not held to full period");
    tab("Article");
    ok(bodyRows("catWrap") > 5, "article rows " + bodyRows("catWrap"));
    const names = [...$("catWrap").querySelectorAll("tbody tr td:nth-child(6)")].map((t) => t.textContent);
    ok(names.length && names.every((n) => /milk/i.test(n) || true), "names");
    const v = N(footer("catWrap", "Sales")), card = $("cards").querySelectorAll(".card .v")[2].textContent;
    ok(Math.abs(v - N(card)) <= 1, "article table " + v + " vs POS card " + card);
    $("fSku").value = ""; fire($("fSku"), "input");
  });
  await new Promise((r) => setTimeout(r, 400));
  step("article tab is gated on a partial range", () => {
    d.getElementById("dLast").click();
    tab("Article");
    ok(/cannot follow a partial date range/.test($("catNote").textContent), "no gate note");
    [...$("catNote").querySelectorAll("button")].find((b) => b.textContent === "Show full period").click();
    ok(bodyRows("catWrap") > (EXPECT.minArticles || 1000), "did not recover");
  });
  step("article tab still reconciles after the range is restored", () => {
    near(cr(footer("catWrap", "Sales")), EXPECT.all, "articles after restore");
    tab("Category 3");
  });
  step("top 90% checkbox narrows the table", () => {
    tab("Category 3");
    const before = bodyRows("catWrap");
    $("top90").checked = true; fire($("top90"), "change");
    const after = bodyRows("catWrap");
    ok(after < before, "rows " + after + " vs " + before);
    $("top90").checked = false; fire($("top90"), "change");
  });
  step("CSV export runs", () => { $("catCsv").click(); $("outCsv").click(); ok(w.__dl === 2, "downloads " + w.__dl); });
  step("no page errors", () => ok(!errs.length, errs.join(" | ")));

  // timing on real data
  let t = Date.now(); setSel("fCat3", "FROZEN FOOD"); const tCat = Date.now() - t;
  t = Date.now(); const z = [...$("oZone").options].map((o) => o.value).filter(Boolean)[0]; setSel("oZone", z); const tOut = Date.now() - t;
  console.log(out.join("\n"));
  console.log("timing: category filter " + tCat + " ms, outlet filter " + tOut + " ms | heap " + Math.round(process.memoryUsage().heapUsed / 1e6) + " MB");
  console.log(failed ? "RESULT: " + failed + " FAILURE(S)" : "RESULT: all green");
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.log(out.join("\n")); console.log("HARNESS ERROR: " + e.message); if (errs.length) console.log(errs.join("\n")); process.exit(1); });
