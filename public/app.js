const cards = [...document.querySelectorAll(".product-card")];
const countEl = document.getElementById("cart-count");
const selectionList = document.getElementById("selection-list");
const emptyState = document.getElementById("selection-empty");
const totalEl = document.getElementById("estimated-total");
const responseBox = document.getElementById("server-response");
const submitBtn = document.getElementById("submit-btn");
let orderSubmissionInProgress = false;
let orderSubmittedSuccessfully = false;
const searchInput = document.getElementById("search-input");
const clearSearch = document.getElementById("clear-search");
const resultCount = document.getElementById("result-count");
const noResults = document.getElementById("no-results");
let activeCategory = "All";

function getQty(card) { return Number(card.querySelector(".qty").textContent) || 0; }
function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, char => ({
    "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;", "'":"&#39;"
  }[char]));
}
function updateSelection() {
  const selected = cards.map(card => ({
    quantity: getQty(card),
    id: card.dataset.productId,
    code: card.dataset.code,
    name: card.querySelector("h3").textContent,
    price: Number(card.querySelector(".product-bottom strong").textContent.replace(/[₹,]/g, "")),
    unit: card.querySelector(".unit-label").textContent
  })).filter(item => item.quantity > 0);

  cards.forEach(card => {
    const plus = card.querySelector(".plus");
    plus.disabled = Number(card.dataset.stock || 0) <= getQty(card);
  });
  const totalQty = selected.reduce((sum, item) => sum + item.quantity, 0);
  const total = selected.reduce((sum, item) => sum + item.price * item.quantity, 0);
  countEl.textContent = totalQty;
  totalEl.textContent = `₹${total.toLocaleString("en-IN")}`;
  emptyState.hidden = selected.length > 0;
  selectionList.innerHTML = selected.map(item =>
    `<div class="selection-row"><span><b>${escapeHtml(item.name)}</b><small>Code ${escapeHtml(item.code)} · ${escapeHtml(item.unit)} × ${item.quantity}</small></span><span>₹${(item.price * item.quantity).toLocaleString("en-IN")}</span></div>`
  ).join("");
}
function filterProducts() {
  const query = searchInput.value.trim().toLowerCase();
  let visible = 0;
  cards.forEach(card => {
    const matchesCategory = activeCategory === "All" || card.dataset.category === activeCategory;
    const matchesQuery = !query || card.dataset.name.includes(query) || card.dataset.code.toLowerCase().includes(query);
    const show = matchesCategory && matchesQuery;
    card.hidden = !show;
    if (show) visible++;
  });
  resultCount.textContent = `Showing ${visible} of ${cards.length} products`;
  noResults.hidden = visible !== 0;
}
cards.forEach(card => {
  card.querySelector(".minus").addEventListener("click", () => {
    const qty = card.querySelector(".qty");
    qty.textContent = Math.max(0, getQty(card) - 1);
    updateSelection();
  });
  card.querySelector(".plus").addEventListener("click", () => {
    const qty = card.querySelector(".qty");
    const stock = Number(card.dataset.stock || 0);
    if (stock <= 0) return;
    qty.textContent = Math.min(stock, getQty(card) + 1);
    updateSelection();
  });
});
document.querySelectorAll(".category-chip").forEach(button => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".category-chip").forEach(b => b.classList.remove("active"));
    button.classList.add("active");
    activeCategory = button.dataset.category;
    filterProducts();
  });
});
searchInput.addEventListener("input", filterProducts);
clearSearch.addEventListener("click", () => {
  searchInput.value = "";
  filterProducts();
  searchInput.focus();
});
document.querySelectorAll(".nav-item").forEach(link => {
  link.addEventListener("click", () => {
    document.querySelectorAll(".nav-item").forEach(item => item.classList.remove("active"));
    link.classList.add("active");
  });
});
document.getElementById("order-form").addEventListener("submit", async event => {
  event.preventDefault();
  // Guard against rapid/double clicks and repeated form submissions.
  if (orderSubmissionInProgress || orderSubmittedSuccessfully) return;
  orderSubmissionInProgress = true;
  submitBtn.disabled = true;
  responseBox.hidden = true;
  responseBox.classList.remove("error");
  const items = cards.map(card => ({product_id: card.dataset.productId, quantity: getQty(card)}))
    .filter(item => item.quantity > 0);
  if (!items.length) {
    responseBox.textContent = "Please add at least one product before sending your selection.";
    responseBox.classList.add("error");
    responseBox.hidden = false;
    orderSubmissionInProgress = false;
    submitBtn.disabled = false;
    submitBtn.innerHTML = 'Send selection to stall <span>→</span>';
    return;
  }
  const mandalName = document.getElementById("mandal-name").value.trim();
  const firmName = document.getElementById("firm-name").value.trim();
  const mobileNo = document.getElementById("mobile-no").value.trim();
  if (!mandalName || !firmName || !/^[0-9]{10}$/.test(mobileNo)) {
    responseBox.textContent = !mandalName ? "Please enter Mandal Name." : (!firmName ? "Please enter Firm Name." : "Please enter a valid 10-digit mobile number.");
    responseBox.classList.add("error");
    responseBox.hidden = false;
    orderSubmissionInProgress = false;
    submitBtn.disabled = false;
    submitBtn.innerHTML = 'Send selection to stall <span>→</span>';
    return;
  }
  const payload = {
    // Keep the API's legacy field names as well as the new customer fields.
    customer_name: mandalName,
    phone: mobileNo,
    mandal_name: mandalName,
    firm_name: firmName,
    mobile_no: mobileNo,
    items
  };
  submitBtn.textContent = "Submitting order…";
  try {
    const response = await fetch("/api/orders", {
      method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(payload)
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.message || "Could not send your selection.");
    // Prefer the PDF bytes returned by Flask. Keep a normal receipt link as a
    // fallback for mobile browsers that block programmatic Blob downloads.
    const pdfRoute = result.order && result.order.order_id
      ? `/my-orders/${encodeURIComponent(result.order.order_id)}/pdf`
      : null;
    let pdfDownloadStarted = false;

    if (typeof result.pdf_base64 === "string" && result.pdf_base64.length > 0) {
      try {
        const binary = atob(result.pdf_base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        const pdfBlob = new Blob([bytes], {type: "application/pdf"});
        const pdfUrl = URL.createObjectURL(pdfBlob);
        const downloadLink = document.createElement("a");
        downloadLink.href = pdfUrl;
        downloadLink.download = result.pdf_filename || "order.pdf";
        downloadLink.style.display = "none";
        document.body.appendChild(downloadLink);
        downloadLink.click();
        downloadLink.remove();
        window.setTimeout(() => URL.revokeObjectURL(pdfUrl), 60000);
        pdfDownloadStarted = true;
      } catch (downloadError) {
        console.warn("Blob PDF download failed; showing receipt link instead.", downloadError);
      }
    }

    // Do not auto-download Excel too: some mobile browsers block the second
    // download, which can make receipt downloads unreliable.
    const receiptLink = pdfRoute
      ? `<a href="${pdfRoute}" download="${escapeHtml(result.pdf_filename || "order.pdf")}" class="pdf-download-link">Download PDF receipt</a>`
      : "";
    const downloadMessage = pdfDownloadStarted
      ? "Your PDF receipt download should begin automatically. If it does not, use the link below."
      : "Use the link below to download your PDF receipt.";
    responseBox.innerHTML = `<strong>${escapeHtml(result.message || "Order saved successfully.")}</strong><br>Reference: ${escapeHtml(result.order?.order_id || "") }<br>Estimated total: ₹${Number(result.order?.total || 0).toLocaleString("en-IN")}<br><span class="pdf-download">${downloadMessage}</span><br>${receiptLink}<br><small>${escapeHtml(result.order?.status || "")}</small>`;
    orderSubmittedSuccessfully = true;
    submitBtn.textContent = "Order submitted ✓";
    cards.forEach(card => card.querySelector(".qty").textContent = "0");
    updateSelection();
    await refreshStock();
    updateSelection();
    document.getElementById("order-form").reset();
  } catch (error) {
    responseBox.textContent = error.message || "Could not connect to the server. Please try again.";
    responseBox.classList.add("error");
  } finally {
    responseBox.hidden = false;
    orderSubmissionInProgress = false;
    // Keep the button locked after success so the same order cannot be submitted twice.
    if (!orderSubmittedSuccessfully) {
      submitBtn.disabled = false;
      submitBtn.innerHTML = 'Send selection to stall <span>→</span>';
    }
  }
});
async function refreshStock() {
  try {
    const response = await fetch("/api/products");
    const result = await response.json();
    if (!response.ok || !result.success) return;
    const stockById = new Map(result.products.map(product => [String(product.id), Number(product.stock)]));
    cards.forEach(card => {
      const stock = stockById.get(String(card.dataset.productId)) ?? 0;
      card.dataset.stock = String(stock);
      const label = card.querySelector("[data-stock-label]");
      if (stock <= 0) {
        label.innerHTML = '<strong class="out-of-stock">Out of stock</strong>';
        card.querySelector(".plus").disabled = true;
        card.querySelector(".plus").title = "Out of stock";
      } else {
        label.innerHTML = `Remaining: <strong>${stock}</strong> units`;
        card.querySelector(".plus").disabled = getQty(card) >= stock;
      }
    });
  } catch (error) { console.warn("Could not refresh stock", error); }
}
updateSelection();
filterProducts();
refreshStock();
