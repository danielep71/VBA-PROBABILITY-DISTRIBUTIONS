"use strict";

// Progressive enhancement: the catalogue and first example remain readable
// without JavaScript. No probability calculation or VBA execution occurs here.
const rows = [...document.querySelectorAll("#distribution-rows tr")];
const filters = [...document.querySelectorAll("[data-filter]")];
const search = document.querySelector("#distribution-search");
let selectedFilter = "all";
function filterCatalogue() {
  const query = search.value.trim().toLowerCase();
  let visible = 0;
  for (const row of rows) {
    const matches = (selectedFilter === "all" || row.dataset.kind === selectedFilter)
      && row.dataset.search.includes(query);
    row.hidden = !matches;
    if (matches) visible += 1;
  }
  document.querySelector("#catalogue-count").textContent = `${visible} of ${rows.length} distribution surfaces`;
  document.querySelector("#catalogue-empty").hidden = visible !== 0;
}
for (const button of filters) {
  button.addEventListener("click", () => {
    selectedFilter = button.dataset.filter;
    for (const other of filters) other.setAttribute("aria-pressed", String(other === button));
    filterCatalogue();
  });
}
search.addEventListener("input", filterCatalogue);
document.querySelector(".catalogue-tools").hidden = false;

const examples = {
  tail: {
    title: "Upper-tail probability",
    description: "Calculate P(Z > 8) for a standard Normal variable without subtracting a rounded CDF from one.",
    formula: "=K_STATS_NormalStandard_Survival(8)",
    note: "The probability is approximately 6.22 × 10⁻¹⁶. This is a usage example, not a domain-wide accuracy guarantee.",
  },
  inverse: {
    title: "A threshold from a tiny tail",
    description: "Find the standard Normal threshold exceeded with probability 10⁻²⁰. Pass the small tail probability directly.",
    formula: "=K_STATS_NormalStandard_InverseSurvival(1E-20)",
    note: "Forming 1 − 10⁻²⁰ first loses that probability in binary64 arithmetic. The direct inverse-survival API keeps q as its input.",
  },
  poisson: {
    title: "A count threshold",
    description: "Find the 99th percentile of a Poisson variable with mean 25.",
    formula: "=K_STATS_Poisson_InverseCumulative(0.99,25)",
    note: "The inverse returns the smallest integer k for which P(X ≤ k) ≥ 0.99. Parameters are probability, then mean.",
  },
  interval: {
    title: "Probability in an interval",
    description: "Calculate P(−1.96 < Z ≤ 1.96) for a standard Normal variable using the interval API.",
    formula: "=K_STATS_NormalStandard_IntervalProbability(-1.96,1.96)",
    note: "For a continuous distribution the endpoint convention does not change the probability. The bounds are lower, then upper.",
  },
};
let selectedExample = "tail";
const exampleButtons = [...document.querySelectorAll("[data-example]")];
const separator = document.querySelector("#separator");
const copyStatus = document.querySelector("#copy-status");
function renderExample() {
  const example = examples[selectedExample];
  document.querySelector("#example-title").textContent = example.title;
  document.querySelector("#example-description").textContent = example.description;
  document.querySelector("#example-code").textContent = separator.value === "it"
    ? example.formula.replaceAll(",", ";").replaceAll(".", ",")
    : example.formula;
  document.querySelector("#example-note").textContent = example.note;
  copyStatus.textContent = "";
}
for (const button of exampleButtons) {
  button.addEventListener("click", () => {
    selectedExample = button.dataset.example;
    for (const other of exampleButtons) other.setAttribute("aria-pressed", String(other === button));
    renderExample();
  });
}
separator.addEventListener("change", renderExample);
document.querySelector("#copy-formula").addEventListener("click", async () => {
  const code = document.querySelector("#example-code");
  const formula = code.textContent;
  try {
    await navigator.clipboard.writeText(formula);
    copyStatus.textContent = `Copied: ${formula}`;
  } catch {
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(code);
    selection.removeAllRanges();
    selection.addRange(range);
    copyStatus.textContent = "Clipboard unavailable. The formula is selected; use Ctrl+C or Command+C to copy it.";
  }
});
document.querySelector(".example-options").hidden = false;
document.querySelector("#separator-control").hidden = false;
document.querySelector("#copy-formula").hidden = false;
