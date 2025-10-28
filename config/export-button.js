(function () {
  "use strict";

  // 🔹 Allowed users (set their Kintone login names or codes here)
  const allowedUsers = ["operations@clavanoprinters.com"];

  // 🔹 Allowed statuses (process management)
  const allowedStatuses = [
    "Under Review",
    "Approved for Client",
    "Request Sent",
  ]; // ✅ customize these

  kintone.events.on("app.record.detail.show", function (event) {
    // Avoid duplicate button
    if (document.getElementById("export-quotation-btn")) return event;

    // 🔹 Get logged-in user
    const currentUser = kintone.getLoginUser().code;

    // 🔹 Get record status (from Process Management field)
    const status = event.record.Status ? event.record.Status.value : null;

    // 🔹 Check conditions
    if (!allowedUsers.includes(currentUser)) {
      console.log(`User ${currentUser} not allowed to export.`);
      return event;
    }

    if (!allowedStatuses.includes(status)) {
      console.log(`Status "${status}" not allowed for export.`);
      return event;
    }

    // ✅ Create modern styled button
    const button = document.createElement("button");
    button.id = "export-quotation-btn";
    button.innerText = "Export"; // ✅ renamed

    // Custom design (Kintone-friendly but cleaner)
    button.style.backgroundColor = "#007bff";
    button.style.color = "#fff";
    button.style.border = "none";
    button.style.padding = "8px 16px";
    button.style.borderRadius = "6px";
    button.style.cursor = "pointer";
    button.style.fontSize = "14px";
    button.style.fontWeight = "bold";
    button.style.margin = "8px"; // ✅ added margin
    button.style.transition = "background 0.3s";

    button.onmouseover = () => (button.style.backgroundColor = "#0056b3");
    button.onmouseout = () => (button.style.backgroundColor = "#007bff");

    // 🔹 Add click action
    button.onclick = async () => {
      const recordId = event.recordId;
      const record = event.record;

      const customer = record.customer.value || "Customer";
      const qty = record.orderQuantity.value || "Qty";

      // 🔹 Build safe filename
      const safeCustomer = customer.replace(/[^a-z0-9]/gi, "_");
      const safeQty = qty.toString().replace(/[^0-9]/g, "");
      const fileName = `${safeCustomer}_${safeQty}_Quotation.xlsx`;

      try {
        const response = await fetch("http://localhost:3000/export", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ recordId }),
        });

        if (!response.ok) throw new Error("Export failed");

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        a.remove();
      } catch (err) {
        alert("❌ Failed to export quotation: " + err.message);
      }
    };

    // Place button in header menu
    kintone.app.record.getHeaderMenuSpaceElement().appendChild(button);

    return event;
  });
})();
