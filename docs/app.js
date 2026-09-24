(() => {
  const STORAGE_KEY = "manual_apply_v1";
  let allJobs = [];
  let currentTab = "new";
  let currentFilter = "all";

  const $jobs = document.getElementById("jobs");
  const $meta = document.getElementById("meta");
  const $empty = document.getElementById("empty");

  function jobId(j) {
    return [j.company, j.title, j.location, j.apply_url || ""]
      .join("|")
      .toLowerCase();
  }

  function loadApplied() {
    try {
      return new Set(JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"));
    } catch {
      return new Set();
    }
  }

  function saveApplied(set) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify([...set]));
  }

  function markApplied(id) {
    const s = loadApplied();
    s.add(id);
    saveApplied(s);
    render();
  }

  function matchesFilter(j) {
    const loc = `${j.location || ""} ${j.country || ""}`.toLowerCase();
    switch (currentFilter) {
      case "netherlands":
        return (
          loc.includes("netherlands") ||
          loc.includes("amsterdam") ||
          loc.includes("rotterdam") ||
          loc.includes("utrecht") ||
          loc.includes("eindhoven") ||
          (j.country || "").toUpperCase() === "NL"
        );
      case "uk":
        return (
          loc.includes("united kingdom") ||
          loc.includes("uk") ||
          loc.includes("london") ||
          loc.includes("england") ||
          loc.includes("scotland") ||
          loc.includes("wales") ||
          (j.country || "").toUpperCase() === "GB" ||
          (j.country || "").toUpperCase() === "UK"
        );
      case "europe":
        return (
          loc.includes("europe") ||
          loc.includes("eu ") ||
          loc.includes("emea") ||
          ["AT","BE","BG","HR","CY","CZ","DK","EE","FI","FR","DE","GR","HU","IS","IE","IT","LV","LT","LU","MT","NL","NO","PL","PT","RO","SK","SI","ES","SE","CH","GB","UK"].includes((j.country || "").toUpperCase()) ||
          /austria|belgium|bulgaria|croatia|cyprus|czech|denmark|estonia|finland|france|germany|greece|hungary|iceland|ireland|italy|latvia|lithuania|luxembourg|malta|netherlands|norway|poland|portugal|romania|slovakia|slovenia|spain|sweden|switzerland/.test(loc)
        );
      case "remote":
        return (j.remote || "").toLowerCase() === "yes" || loc.includes("remote");
      case "visa":
        return (j.visa || "").toLowerCase().includes("yes") || (j.visa || "").toLowerCase().includes("mentioned");
      default:
        return true;
    }
  }

  function render() {
    const applied = loadApplied();
    let list = allJobs.filter((j) => {
      const id = jobId(j);
      const isApplied = applied.has(id);
      if (currentTab === "new" && isApplied) return false;
      if (currentTab === "applied" && !isApplied) return false;
      return matchesFilter(j);
    });

    $meta.textContent = `${list.length} job${list.length === 1 ? "" : "s"} · ${currentTab} · ${currentFilter}`;
    $jobs.innerHTML = "";
    $empty.classList.toggle("hidden", list.length > 0);

    for (const j of list) {
      const id = jobId(j);
      const card = document.createElement("article");
      card.className = "card";

      const badges = [];
      if ((j.remote || "").toLowerCase() === "yes") {
        badges.push('<span class="badge remote">Remote</span>');
      }
      if ((j.visa || "").toLowerCase().includes("yes") || (j.visa || "").toLowerCase().includes("mentioned")) {
        badges.push('<span class="badge visa">Visa mentioned</span>');
      }

      card.innerHTML = `
        <h2>${escapeHtml(j.title || "")}</h2>
        <div class="company">${escapeHtml(j.company || "")}</div>
        <div class="meta-row">
          <span>${escapeHtml(j.location || "—")}</span>
          <span>${escapeHtml(j.country || "—")}</span>
          <span>${escapeHtml(j.posted_date || "—")}</span>
          ${badges.join(" ")}
        </div>
        <div class="actions">
          <a class="btn btn-apply" href="${escapeAttr(j.apply_url || "#")}" target="_blank" rel="noopener noreferrer">APPLY</a>
          ${currentTab === "new" ? `<button class="btn btn-mark" type="button" data-id="${escapeAttr(id)}">MARK APPLIED</button>` : ""}
        </div>
      `;
      $jobs.appendChild(card);
    }

    $jobs.querySelectorAll(".btn-mark").forEach((btn) => {
      btn.addEventListener("click", () => markApplied(btn.dataset.id));
    });
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&" + "amp;")
      .replace(/</g, "&" + "lt;")
      .replace(/>/g, "&" + "gt;")
      .replace(/"/g, "&" + "quot;");
  }

  function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, "&#39;");
  }

  document.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      currentTab = btn.dataset.tab;
      render();
    });
  });

  document.querySelectorAll(".filter").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".filter").forEach((b) => b.classList.remove("active"));
      btn.classList.add("active");
      currentFilter = btn.dataset.filter;
      render();
    });
  });

  fetch("jobs_latest.json", { cache: "no-store" })
    .then((r) => {
      if (!r.ok) throw new Error("Failed to load jobs");
      return r.json();
    })
    .then((data) => {
      allJobs = Array.isArray(data) ? data : [];
      render();
    })
    .catch((err) => {
      $meta.textContent = "Could not load jobs_latest.json";
      console.error(err);
    });
})();
