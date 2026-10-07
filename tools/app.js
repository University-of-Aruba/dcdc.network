(function () {
  "use strict";

  var CFG = window.REGISTRY_CONFIG || {};
  // showInstitutions: false hides every institution column, for a public copy published
  // before the institutions have confirmed their entries.
  var SHOW_INST = CFG.showInstitutions !== false;
  var VIEWS = SHOW_INST ? ["start", "theme", "explore", "matrix", "profiles", "ai", "about"] : ["start", "theme", "explore", "ai", "about"];

  var INSTITUTIONS = [
    { key: "UA", name: "University of Aruba" },
    { key: "UoC", name: "University of Curaçao" },
    { key: "USM", name: "University of St. Martin" },
    { key: "IPA", name: "Instituto Pedagogico Arubano" },
    { key: "DCDC network", name: "DCDC network (services the network runs or teaches)", short: "DCDC" }
  ];
  var STATUS = [
    ["Supported", "s-supported", "Provided by the institution, with help from IT, library or training."],
    ["Provided", "s-provided", "Provided by the institution, with little or no help."],
    ["Training only", "s-training", "A free tool the institution trains people in, without providing it."],
    ["Informal use", "s-informal", "Used on people's own initiative or own licence. Demand without an offer."],
    ["Evaluating", "s-evaluating", "Being considered or piloted."],
    ["Not offered", "s-notoffered", "Asked and confirmed: not provided."]
  ];
  var STATUS_MAP = {};
  STATUS.forEach(function (s) { STATUS_MAP[s[0].toLowerCase()] = { label: s[0], cls: s[1], meaning: s[2] }; });
  var FAIR = [["F", "Findable"], ["A", "Accessible"], ["I", "Interoperable"], ["R", "Reusable"]];
  var FREE_COSTS = ["free", "open source (free)", "free tier + paid"];
  var TABS = { catalogue: "Catalogue", offering: "Institutional offering", needs: "Needs", profile: "Institution profile", themes: "Themes" };
  var SNAPSHOT = { catalogue: "data/catalogue.csv", offering: "data/offering.csv", needs: "data/needs.csv", profile: "data/profile.csv", themes: "data/themes.csv" };

  var model = null;
  var state = { view: "start", theme: "", need: "", q: "", tool: "", stream: "", free: false, oss: false, net: false };

  // ---- Data -------------------------------------------------------------------
  function parseCSV(text) {
    var rows = [], row = [], field = "", inQuotes = false, i = 0, c;
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    while (i < text.length) {
      c = text.charAt(i);
      if (inQuotes) {
        if (c === '"') {
          if (text.charAt(i + 1) === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ",") { row.push(field); field = ""; i++; continue; }
      if (c === "\r") { i++; continue; }
      if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; i++; continue; }
      field += c; i++;
    }
    if (field !== "" || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
  // Header names are matched loosely so the Sheet can gain or reword columns:
  // "Tool ID (automatic)" becomes "tool id", "Arrangement (licence, ...)" becomes "arrangement".
  function normKey(h) { return String(h).replace(/\(.*?\)/g, "").replace(/\s+/g, " ").trim().toLowerCase(); }
  function toObjects(rows) {
    if (!rows.length) return [];
    var head = rows[0].map(normKey);
    return rows.slice(1).map(function (r) {
      var o = {};
      head.forEach(function (h, j) { if (h && !(h in o)) o[h] = (r[j] || "").trim(); });
      return o;
    });
  }
  function fetchText(url) {
    return fetch(url, { cache: "no-store" }).then(function (r) {
      if (!r.ok) throw new Error("HTTP " + r.status);
      return r.text();
    }).then(function (t) {
      if (/^\s*</.test(t)) throw new Error("received a web page instead of data; check the Sheet's sharing setting");
      return t;
    });
  }
  function sheetUrl(tab) {
    return "https://docs.google.com/spreadsheets/d/" + encodeURIComponent(CFG.sheetId) +
      "/gviz/tq?tqx=out:csv&headers=1&sheet=" + encodeURIComponent(tab);
  }
  function loadSet(urls) {
    var keys = Object.keys(urls);
    // Themes are optional, so a Sheet without a Themes tab still loads
    return Promise.all(keys.map(function (k) {
      return k === "themes" ? fetchText(urls[k]).catch(function () { return ""; }) : fetchText(urls[k]);
    })).then(function (texts) {
      var d = {};
      keys.forEach(function (k, i) { d[k] = toObjects(parseCSV(texts[i])); });
      if (!d.catalogue.length || !("id" in d.catalogue[0]) || !("tool" in d.catalogue[0])) {
        throw new Error("the Catalogue tab does not have the expected columns");
      }
      return d;
    });
  }
  function load() {
    if (!CFG.sheetId) return loadSet(SNAPSHOT).then(function (d) { d.source = "snapshot"; return d; });
    var live = {};
    Object.keys(TABS).forEach(function (k) { live[k] = sheetUrl(TABS[k]); });
    return loadSet(live).then(function (d) { d.source = "live"; return d; }).catch(function (err) {
      return loadSet(SNAPSHOT).then(function (d) { d.source = "fallback"; d.error = err; return d; });
    });
  }

  function buildModel(d) {
    var tools = d.catalogue.filter(function (r) { return r.id && r.tool; }).map(function (r, i) {
      return {
        order: i, id: r.id, stream: r.stream, category: r.category, name: r.tool, what: r["what it is"],
        needs: (r["needs served"] || "").split(";").map(function (s) { return s.trim(); }).filter(Boolean),
        cost: r["cost model"], open: r["open source"], platform: r["runs on"],
        learning: r["learning curve"], teaching: r["teaching fit"], pros: r.pros, cons: r.cons,
        fair: { F: r.f, A: r.a, I: r.i, R: r.r }, fairNote: r["fair note"],
        smallNote: r["fit for small institutions"], link: r.link, offer: {},
        contacts: (r["network contacts"] || "").split(";").map(function (s) { return s.trim(); }).filter(Boolean)
      };
    });
    var byId = {}, byName = {};
    tools.forEach(function (t) { byId[t.id] = t; byName[t.name.toLowerCase()] = t; });
    var insts = SHOW_INST ? INSTITUTIONS.slice() : [];
    (SHOW_INST ? d.offering : []).forEach(function (r) {
      var inst = r.institution;
      if (!inst) return;
      var t = byId[r["tool id"]] || byName[(r.tool || "").toLowerCase()];
      if (!t) return;
      if (!insts.some(function (x) { return x.key === inst; })) insts.push({ key: inst, name: inst });
      t.offer[inst] = {
        status: r["offering status"] || "", draft: !r["confirmed on"], confirmed: r["confirmed on"] || "",
        who: r["who can use it"] || "", arrangement: r.arrangement || "", support: r["support available"] || "",
        source: r.source || "", notes: r["notes and questions to ask"] || "", contact: r.contact || ""
      };
    });
    var needs = d.needs.filter(function (r) { return r.need; }).map(function (r) {
      return { need: r.need, meaning: r["what it means"] || "", example: r["example question"] || "" };
    });
    var profile = d.profile.filter(function (r) { return r.question; });
    // Themes: a handful of entry points, each with two or three recommended tools ("picks").
    // Picks are written as "ID=reason | ID+ID=reason"; a "+" joins tools that fill the same role.
    var themes = (d.themes || []).filter(function (r) { return r.key && r.theme; }).map(function (r) {
      var streams = (r.streams || "").split(";").map(function (s) { return s.trim(); }).filter(Boolean);
      var picks = (r.picks || "").split("|").map(function (p) {
        var i = p.indexOf("=");
        var ids = (i < 0 ? p : p.slice(0, i)).split("+").map(function (s) { return s.trim(); });
        return { tools: ids.map(function (id) { return byId[id]; }).filter(Boolean), why: i < 0 ? "" : p.slice(i + 1).trim() };
      }).filter(function (p) { return p.tools.length; });
      return { key: r.key, title: r.theme, summary: r.summary || "", streams: streams, picks: picks,
        guidance: r.guidance || "", linkLabel: r["link label"] || "", link: r.link || "" };
    });
    return { tools: tools, insts: insts, needs: needs, profile: profile, themes: themes, source: d.source, error: d.error };
  }

  // ---- Helpers ------------------------------------------------------------------
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function $(id) { return document.getElementById(id); }
  function statusOf(o) {
    if (!o || !o.status) return null;
    return STATUS_MAP[o.status.toLowerCase()] || { label: o.status, cls: "s-other", meaning: "" };
  }
  function offeredCount(t) {
    return model.insts.reduce(function (n, inst) {
      var s = statusOf(t.offer[inst.key]);
      return n + (s && s.cls !== "s-notoffered" ? 1 : 0);
    }, 0);
  }
  function fairCls(v) {
    v = (v || "").toLowerCase();
    return v === "strong" ? "f-strong" : v === "partial" ? "f-partial" : v === "weak" ? "f-weak" : "f-na";
  }
  // Licence, kept apart from cost: free is not the same as open source (REDCap, QDA Miner Lite),
  // and open source can still be sold as a service (GitLab, Posit Cloud).
  function licence(t) {
    var v = (t.open || "").toLowerCase();
    if (v === "yes") return { label: "Open source", cls: "lic-open" };
    if (v.indexOf("partly") === 0) return { label: "Open core", cls: "lic-core" };
    if (v === "no") return { label: "Proprietary", cls: "lic-prop" };
    return null;
  }
  function contactHTML(c) {
    var m = /^(.*?)\s*<([^>]+@[^>]+)>$/.exec(c);
    return m ? '<a href="mailto:' + esc(m[2]) + '">' + esc(m[1]) + '</a>' : esc(c);
  }
  function badge(inst, o) {
    var s = statusOf(o), label = inst.short || inst.key;
    if (!s) {
      return '<span class="badge s-none" title="' + esc(inst.name) + ': not yet asked">' +
        '<b>' + esc(label) + '</b><span>not yet asked</span></span>';
    }
    return '<span class="badge ' + s.cls + (o.draft ? " is-draft" : "") + '" title="' + esc(inst.name + ": " + s.label +
      (o.draft ? " (draft, unconfirmed)" : "")) + '"><b>' + esc(label) + '</b><span>' + esc(s.label) +
      (o.draft ? ' <i>draft</i>' : "") + '</span></span>';
  }
  function legendHTML(long) {
    var items = STATUS.map(function (s) {
      return '<span class="lg"><span class="sw ' + s[1] + '"></span>' + esc(s[0]) +
        (long ? '<small>' + esc(s[2]) + '</small>' : "") + '</span>';
    });
    items.push('<span class="lg"><span class="sw s-none"></span>Not yet asked' +
      (long ? '<small>Nobody has asked this institution yet.</small>' : "") + '</span>');
    items.push('<span class="lg"><span class="sw s-supported is-draft"></span>Draft' +
      (long ? '<small>Recorded but not yet confirmed by the institution.</small>' : "") + '</span>');
    return items.join("");
  }

  // ---- Explore view -----------------------------------------------------------------
  function renderChips() {
    var html = ['<button type="button" class="chip' + (state.need ? "" : " on") + '" data-need="">All tools</button>'];
    model.needs.forEach(function (n) {
      var count = model.tools.filter(function (t) { return t.needs.indexOf(n.need) >= 0; }).length;
      html.push('<button type="button" class="chip' + (state.need === n.need ? " on" : "") + '" data-need="' +
        esc(n.need) + '">' + esc(n.need) + ' <span class="n">' + count + '</span></button>');
    });
    $("need-chips").innerHTML = html.join("");
    var info = model.needs.filter(function (n) { return n.need === state.need; })[0];
    $("need-info").innerHTML = info
      ? '<p><strong>' + esc(info.need) + ':</strong> ' + esc(info.meaning) + '</p>' +
        (info.example ? '<p class="example">A question someone might bring: “' + esc(info.example) + '”</p>' : "")
      : "";
  }
  function matches(t) {
    if (state.tool) return t.id === state.tool;
    if (state.need && t.needs.indexOf(state.need) < 0) return false;
    if (state.stream && t.stream !== state.stream) return false;
    if (state.free && FREE_COSTS.indexOf((t.cost || "").toLowerCase()) < 0) return false;
    if (state.oss && !(licence(t) && licence(t).cls !== "lic-prop")) return false;
    if (state.net && offeredCount(t) === 0) return false;
    if (state.q) {
      var hay = [t.name, t.what, t.category, t.stream, t.needs.join(" "), t.pros, t.cons].join(" ").toLowerCase();
      var words = state.q.toLowerCase().split(/\s+/).filter(Boolean);
      for (var i = 0; i < words.length; i++) if (hay.indexOf(words[i]) < 0) return false;
    }
    return true;
  }
  function offerDetails(t) {
    var rows = model.insts.filter(function (inst) { return t.offer[inst.key]; }).map(function (inst) {
      var o = t.offer[inst.key], s = statusOf(o);
      var bits = [];
      if (o.arrangement) bits.push(esc(o.arrangement));
      if (o.who) bits.push("For: " + esc(o.who));
      if (o.support) bits.push("Support: " + esc(o.support));
      if (o.contact) bits.push("Contact: " + contactHTML(o.contact));
      var src = o.source ? esc(o.source) : "";
      src += o.confirmed ? " Confirmed " + esc(o.confirmed) + "." : "";
      return '<li><div class="oh">' + esc(inst.short || inst.key) + ' <span class="st ' + (s ? s.cls : "s-none") +
        (o.draft && s ? " is-draft" : "") + '">' + (s ? esc(s.label) + (o.draft ? " · draft" : "") : "Open question") +
        '</span></div>' + (bits.length ? '<p>' + bits.join(". ") + '</p>' : "") +
        (o.notes ? '<p class="notes">' + esc(o.notes) + '</p>' : "") +
        (src ? '<p class="src">' + src + '</p>' : "") + '</li>';
    });
    return rows.length ? '<h4>What the network has recorded</h4><ul class="offers">' + rows.join("") + '</ul>'
      : '<p class="muted">No institution has recorded anything for this tool yet.</p>';
  }
  function card(t) {
    var lic = licence(t);
    var meta = (lic ? '<li class="lic ' + lic.cls + '">' + esc(lic.label) + '</li>' : "") +
      [/^open source \(free\)$/i.test(t.cost || "") ? "Free" : t.cost, t.learning ? "Learning curve: " + t.learning.toLowerCase() : "", t.teaching && t.teaching !== "n/a" ? "Teaching fit: " + t.teaching.toLowerCase() : ""]
      .filter(Boolean).map(function (m) { return '<li>' + esc(m) + '</li>'; }).join("");
    var fair = FAIR.map(function (f) {
      return '<span class="pill ' + fairCls(t.fair[f[0]]) + '" title="' + esc(f[1] + ": " + (t.fair[f[0]] || "not rated")) + '">' +
        f[0] + '<span class="sr"> ' + esc(f[1]) + ': ' + esc(t.fair[f[0]] || "not rated") + '</span></span>';
    }).join("");
    var badges = SHOW_INST ? '<div class="badges" aria-label="Offered by">' +
      model.insts.map(function (inst) { return badge(inst, t.offer[inst.key]); }).join("") + '</div>' : "";
    return '<article class="card" id="tool-' + esc(t.id) + '">' +
      '<p class="stream">' + esc(t.stream) + (t.category ? " · " + esc(t.category) : "") + '</p>' +
      '<h3>' + esc(t.name) + '</h3>' +
      '<p class="what">' + esc(t.what) + '</p>' +
      '<ul class="meta">' + meta + '</ul>' +
      badges +
      '<div class="fairrow"><span class="fl">FAIR support</span>' + fair + '</div>' +
      '<details><summary>' + (SHOW_INST ? "Pros, cons and what each institution offers" : "Pros and cons") + '</summary><div class="det">' +
      '<div class="pc"><div><h4>Pros</h4><p>' + esc(t.pros) + '</p></div><div><h4>Cons</h4><p>' + esc(t.cons) + '</p></div></div>' +
      (t.fairNote ? '<h4>FAIR</h4><p>' + esc(t.fairNote) + '</p>' : "") +
      (t.smallNote ? '<h4>Fit for small institutions</h4><p>' + esc(t.smallNote) + '</p>' : "") +
      (t.contacts.length ? '<h4>Ask someone in the network</h4><p>' + t.contacts.map(contactHTML).join(", ") + '</p>' : "") +
      (SHOW_INST ? offerDetails(t) : "") +
      '<p class="tail">' + (t.platform ? 'Runs on: ' + esc(t.platform) + '. ' : "") +
      (t.link ? '<a href="' + esc(t.link) + '" target="_blank" rel="noopener">Website</a>' : "") +
      ' <span class="id">' + esc(t.id) + '</span></p>' +
      '</div></details></article>';
  }
  function renderCards() {
    var list = model.tools.filter(matches);
    var q = state.q.toLowerCase();
    var nameHit = function (t) { return q && t.name.toLowerCase().indexOf(q) >= 0 ? 1 : 0; };
    list.sort(function (a, b) {
      return nameHit(b) - nameHit(a) ||
        ((state.need || state.net) ? offeredCount(b) - offeredCount(a) : 0) || a.order - b.order;
    });
    $("count").textContent = list.length === model.tools.length ? "All " + list.length + " tools"
      : list.length + " of " + model.tools.length + " tools" + (state.need ? ", tools already offered in the network first" : "");
    $("cards").innerHTML = list.length ? list.map(card).join("")
      : '<p class="empty">Nothing matches. Clear a filter, or add the tool to the registry.</p>';
    // A single result is usually a click from the matrix: show it opened
    if (list.length === 1) $("cards").querySelector("details").open = true;
  }
  function renderExplore() { renderChips(); renderCards(); }

  // ---- Start and theme views --------------------------------------------------------
  // Short names on the theme pages: "Google Workspace (Drive, ...)" becomes "Google Workspace"
  function shortName(t) { return t.name.replace(/\s*\(.*\)\s*$/, ""); }
  function pickNames(p) { return p.tools.map(function (t) { return esc(shortName(t)); }).join(" or "); }
  function themeTools(th) { return model.tools.filter(function (t) { return th.streams.indexOf(t.stream) >= 0; }); }
  function renderStart() {
    $("themes").innerHTML = model.themes.map(function (th) {
      return '<a class="theme" href="#theme?t=' + encodeURIComponent(th.key) + '">' +
        '<h3>' + esc(th.title) + '</h3><p class="sum">' + esc(th.summary) + '</p>' +
        '<p class="picks-l">Our picks</p><ul class="picks-s">' + th.picks.map(function (p) {
          return '<li>' + pickNames(p) + '</li>';
        }).join("") + '</ul><span class="more">' + themeTools(th).length + ' tools in this theme &rarr;</span></a>';
    }).join("");
  }
  // A tool collapsed to one line; opening it shows the full card.
  function toolRow(t, label) {
    var lic = licence(t);
    return '<details class="trow"><summary><span class="tn">' + esc(label || shortName(t)) + '</span>' +
      (lic ? '<span class="lic ' + lic.cls + '">' + esc(lic.label) + '</span>' : "") +
      '<span class="tw">' + esc(t.what) + '</span></summary>' + card(t) + '</details>';
  }
  function renderTheme() {
    var th = model.themes.filter(function (x) { return x.key === state.theme; })[0];
    if (!th) { state.view = "start"; show(); return; }
    var picked = [];
    th.picks.forEach(function (p) { p.tools.forEach(function (t) { picked.push(t.id); }); });
    var rest = themeTools(th).filter(function (t) { return picked.indexOf(t.id) < 0; });
    var link = th.link ? '<p class="tlink"><a href="' + esc(th.link) + '"' + (th.link.charAt(0) === "#" ? "" : ' target="_blank" rel="noopener"') +
      '>' + esc(th.linkLabel || th.link) + '</a></p>' : "";
    $("theme-body").innerHTML =
      '<p class="back"><a href="#start">&larr; All themes</a></p>' +
      '<h2>' + esc(th.title) + '</h2><p class="intro">' + esc(th.summary) + '</p>' +
      (th.guidance ? '<p class="guidance">' + esc(th.guidance) + '</p>' : "") + link +
      '<h3 class="sub">Our picks</h3><div class="picks">' + th.picks.map(function (p) {
        return '<div class="pick"><p class="pn">' + pickNames(p) + '</p><p class="pw">' + esc(p.why) + '</p>' +
          p.tools.map(function (t) { return toolRow(t, p.tools.length > 1 ? "" : "Details, pros and cons"); }).join("") + '</div>';
      }).join("") + '</div>' +
      (rest.length ? '<details class="rest"><summary>Other tools in this theme (' + rest.length + ')</summary>' +
        '<p class="muted">Also in use across the network. Open a tool for its pros, cons and FAIR notes.</p>' +
        rest.map(function (t) { return toolRow(t); }).join("") + '</details>' : "");
    window.scrollTo(0, 0);
  }

  // ---- Matrix view ------------------------------------------------------------------
  function renderMatrix() {
    var summary = model.insts.map(function (inst) {
      var rec = 0, conf = 0;
      model.tools.forEach(function (t) {
        var o = t.offer[inst.key];
        if (o && o.status) { rec++; if (!o.draft) conf++; }
      });
      return '<div class="is"><b>' + esc(inst.short || inst.key) + '</b><span>' + rec + ' recorded</span><span>' +
        conf + ' confirmed</span></div>';
    }).join("");
    $("inst-summary").innerHTML = summary;
    var head = '<thead><tr><th class="tcol">Tool</th>' + model.insts.map(function (inst) {
      return '<th title="' + esc(inst.name) + '">' + esc(inst.short || inst.key) + '</th>';
    }).join("") + '<th class="fcol">FAIR</th></tr></thead>';
    var body = [], lastStream = null, ncol = model.insts.length + 2;
    model.tools.forEach(function (t) {
      if (t.stream !== lastStream) {
        body.push('<tr class="grp"><th colspan="' + ncol + '">' + esc(t.stream) + '</th></tr>');
        lastStream = t.stream;
      }
      var cells = model.insts.map(function (inst) {
        var o = t.offer[inst.key], s = statusOf(o);
        if (!s) return '<td class="c s-none" title="' + esc(inst.name) + ': not yet asked"><span class="sr">not yet asked</span></td>';
        return '<td class="c ' + s.cls + (o.draft ? " is-draft" : "") + '" title="' + esc(inst.name + ": " + s.label +
          (o.draft ? " (draft)" : "")) + '">' + esc(s.label) + (o.draft ? '<i>draft</i>' : "") + '</td>';
      }).join("");
      var fair = FAIR.map(function (f) { return '<span class="mini ' + fairCls(t.fair[f[0]]) + '" title="' + esc(f[1] + ": " + (t.fair[f[0]] || "")) + '">' + f[0] + '</span>'; }).join("");
      body.push('<tr><th class="tcol"><a href="#explore?tool=' + encodeURIComponent(t.id) + '">' + esc(t.name) + '</a></th>' +
        cells + '<td class="fcol">' + fair + '</td></tr>');
    });
    $("matrix").innerHTML = head + '<tbody>' + body.join("") + '</tbody>';
  }

  // ---- Profiles view ----------------------------------------------------------------
  function renderProfiles() {
    var cols = INSTITUTIONS.filter(function (i) { return i.key !== "DCDC network"; });
    var head = '<thead><tr><th class="qcol">Question</th>' + cols.map(function (i) {
      return '<th title="' + esc(i.name) + '">' + esc(i.key) + '</th>';
    }).join("") + '</tr></thead>';
    var body = [], last = null;
    model.profile.forEach(function (r) {
      if (r.section && r.section !== last) {
        body.push('<tr class="grp"><th colspan="' + (cols.length + 1) + '">' + esc(r.section) + '</th></tr>');
        last = r.section;
      }
      body.push('<tr><th class="qcol">' + esc(r.question) + (r["why it matters"] ? '<small>' + esc(r["why it matters"]) + '</small>' : "") +
        '</th>' + cols.map(function (i) {
          var v = r[i.key.toLowerCase()] || "";
          return '<td class="' + (v ? "" : "blank") + '">' + (v ? esc(v) : '<span class="sr">not yet asked</span>') + '</td>';
        }).join("") + '</tr>');
    });
    $("profile").innerHTML = head + '<tbody>' + body.join("") + '</tbody>';
  }

  // ---- Routing ------------------------------------------------------------------------
  function readHash() {
    var h = location.hash.replace(/^#/, "");
    var parts = h.split("?");
    var p = new URLSearchParams(parts[1] || "");
    state.view = VIEWS.indexOf(parts[0]) >= 0 ? parts[0] : "start";
    if (state.view === "theme") state.theme = p.get("t") || "";
    if (state.view === "explore") {
      state.need = p.get("need") || "";
      state.q = p.get("q") || "";
      state.tool = p.get("tool") || "";
      $("q").value = state.q;
    }
  }
  function writeHash() {
    var p = new URLSearchParams();
    if (state.need) p.set("need", state.need);
    if (state.q) p.set("q", state.q);
    if (state.tool) p.set("tool", state.tool);
    var qs = p.toString();
    history.replaceState(null, "", "#explore" + (qs ? "?" + qs : ""));
  }
  function show() {
    ["start", "theme", "explore", "matrix", "profiles", "ai", "about"].forEach(function (v) {
      $("view-" + v).hidden = v !== state.view;
    });
    Array.prototype.forEach.call(document.querySelectorAll(".tabs a"), function (a) {
      if (a.getAttribute("data-view") === (state.view === "theme" ? "start" : state.view)) a.setAttribute("aria-current", "page");
      else a.removeAttribute("aria-current");
    });
    if (state.view === "start") renderStart();
    if (state.view === "theme") renderTheme();
    if (state.view === "explore") renderExplore();
    if (state.view === "matrix") renderMatrix();
    if (state.view === "profiles") renderProfiles();
  }

  function sourceText() {
    var t = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    if (model.source === "live") return "Live from the registry Google Sheet, loaded at " + t + ".";
    if (model.source === "fallback") return "Could not reach the Google Sheet (" + model.error.message + "). Showing the " + (CFG.snapshotLabel || "snapshot") + ".";
    return "Showing the " + (CFG.snapshotLabel || "snapshot") + ". Not yet connected to the Google Sheet.";
  }

  function init() {
    load().then(function (d) {
      model = buildModel(d);
      $("loading").hidden = true;
      $("source").textContent = sourceText();
      $("source2").textContent = model.source === "live" ? "Data: live Google Sheet" : "Data: " + (CFG.snapshotLabel || "snapshot");
      if (model.source === "fallback") document.body.classList.add("is-fallback");
      var streams = [];
      model.tools.forEach(function (t) { if (t.stream && streams.indexOf(t.stream) < 0) streams.push(t.stream); });
      $("stream").innerHTML += streams.map(function (s) { return '<option>' + esc(s) + '</option>'; }).join("");
      if (!SHOW_INST) {
        document.body.classList.add("no-inst");
        Array.prototype.forEach.call(document.querySelectorAll('.tabs a[data-view="matrix"], .tabs a[data-view="profiles"]'),
          function (a) { a.remove(); });
      }
      $("legend").innerHTML = legendHTML(false);
      $("legend2").innerHTML = legendHTML(false);
      $("legend3").innerHTML = legendHTML(true);

      $("need-chips").addEventListener("click", function (e) {
        var b = e.target.closest("button[data-need]");
        if (!b) return;
        state.need = b.getAttribute("data-need");
        state.tool = "";
        writeHash(); renderExplore();
      });
      $("q").addEventListener("input", function () { state.q = this.value.trim(); state.tool = ""; writeHash(); renderCards(); });
      $("stream").addEventListener("change", function () { state.stream = this.value; renderCards(); });
      $("free").addEventListener("change", function () { state.free = this.checked; renderCards(); });
      $("oss").addEventListener("change", function () { state.oss = this.checked; renderCards(); });
      $("net").addEventListener("change", function () { state.net = this.checked; renderCards(); });
      window.addEventListener("hashchange", function () { readHash(); show(); });
      readHash(); show();
    }).catch(function (err) {
      $("loading").textContent = "The registry could not be loaded: " + err.message;
    });
  }
  init();
})();
