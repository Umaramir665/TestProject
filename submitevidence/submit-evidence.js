Office.onReady(() => {
  const rs = Office.context.roamingSettings;
  const apiKey = rs.get("apiKey") || "";
  const agreementId = rs.get("agreementId") || "";


const getTrimmedValue = (id) => {
  const el = document.getElementById(id);
  const value = (el && typeof el.value === "string") ? el.value.trim() : "";
  return value;
};

const getAuditContextFromInputs = () => {
  const eventStatementDefinition = getTrimmedValue("eventStatementDefinition");
  const proofCertificateForTheAttentionOf = getTrimmedValue("forTheAttentionOf");

  return {
    eventStatementDefinition,
    proofCertificateForTheAttentionOf
  };
};

const validateAuditContext = (auditContext) => {
  if (!auditContext.eventStatementDefinition) return "Event statement definition is required.";
  if (!auditContext.proofCertificateForTheAttentionOf) return "For the attention of is required.";
  return null;
};

const createEmptyAuditContext = () => ({
  eventStatementDefinition: "",
  proofCertificateForTheAttentionOf: ""
});

  
// Check if config is missing
if (!apiKey || !agreementId) {
  const loader = document.getElementById("loader-overlay");
  const content = document.getElementById("content");
  const buttons = document.getElementById("buttons");
  const description = document.getElementById("description");
  const downloadBtn = document.getElementById("downloadBtn");

  if (loader) loader.style.display = "none";
  if (buttons) buttons.style.display = "none";
  if (description) description.style.display = "none";
  if (downloadBtn) downloadBtn.style.display = "none";
document.getElementById("description").style.display = "none";
document.getElementById("downloadBtn").style.display = "none";
  if (content) {
    content.style.display = "block";
   content.innerHTML = `
<div style="
max-width: 400px;
margin: 40px auto;
padding: 20px 25px;
background: #fff3f3;
border: 1px solid #f8d7da;
border-radius: 12px;
box-shadow: 0 4px 12px rgba(0,0,0,0.06);
text-align: center;
font-family: 'Segoe UI', sans-serif;
">
<div style="font-size: 36px; margin-bottom: 16px; color: #d32f2f;">⚠️</div>
<h3 style="color: #b71c1c; font-weight: 600; margin-bottom: 8px;">Configuration Required</h3>
<p style="color: #5f5f5f; font-size: 15px;">
  Your <strong>API Key</strong> or <strong>Agreement ID</strong> is missing.<br>Please complete the setup to continue.
</p>
</div>
`;
  }

  return; // Stop execution
}
const item = Office.context.mailbox.item;
const evidence = [];
let payload = null;
const pushIfValid = (key, value) => {
  if (value && value.trim() !== "") {
    evidence.push({ key, value });
  }
};

 const parseEmail = (input, fallback) => {
  if (!input) return fallback || "";
    const match = /<([^>]+)>/.exec(input);
    if (match && match[1]) {
      return match[1].trim();
    }
    // Fallback if input is already an email (no brackets)
    return input.includes("@") ? input.trim() : (fallback || "");
  };

  const getRecipientEmails = (list) => {
    if (!Array.isArray(list)) return "";
    return list
      .map(r => parseEmail(r.displayName || r.emailAddress?.address || ""))
      .filter(email => !!email)
      .join(", ");
  };

  // From
  if (item.from) {
    //console.log(item.from);
    const email = parseEmail(item.from.emailAddress || item.from.displayName|| "");
    pushIfValid("From", email);
  }

  // To, CC, BCC
  pushIfValid("To", getRecipientEmails(item.to));
  pushIfValid("CC", getRecipientEmails(item.cc));
  pushIfValid("BCC", getRecipientEmails(item.bcc));

  pushIfValid("Subject", item.subject || "");
  pushIfValid("DateTime", item.dateTimeCreated?.toISOString?.() || new Date().toISOString());

  item.body.getAsync("text", result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      pushIfValid("Body", result.value || "");
    }

    payload = {
      evidence,
      tags: ["Outlook", "Evidence"],
      header: {
        sourceSystemDispatchReference: (item.subject && item.subject.trim()) ? item.subject : "OutlookAddin",
        serviceAgreementIdentifier: agreementId,
        where: "Outlook",
        when: new Date().toISOString(),
        auditContext: createEmptyAuditContext()
      }
    };
    document.getElementById("loader-overlay").style.display = "none";
    document.getElementById("content").style.display = "block";
    const auditContextSection = document.getElementById("auditContextSection");
    if (auditContextSection) auditContextSection.style.display = "block";

    try {
      const currentUserEmail = parseEmail(
        Office.context.mailbox.userProfile.emailAddress || ""
      ).toLowerCase();
      const senderEmail = parseEmail(
        item.from?.emailAddress || item.from?.displayName || ""
      ).toLowerCase();
      const isSentByMe = currentUserEmail && senderEmail && currentUserEmail === senderEmail;

      const rawDate = item.dateTimeCreated || item.dateTimeModified || new Date();
      const formattedDate = new Date(rawDate).toLocaleString();
      const dateLabel = isSentByMe ? "Sent" : "Received";
      const subject = (item.subject && item.subject.trim()) ? item.subject.trim() : "(No Subject)";

      const eventDefField = document.getElementById("eventStatementDefinition");
      if (eventDefField) {
        eventDefField.value = subject + " | " + formattedDate;
      }
    } catch (e) {
      console.error("Auto-populate eventStatementDefinition failed:", e);
    }

    // Render evidence in styled blocks (not JSON)
    const contentDiv = document.getElementById("content");
    contentDiv.innerHTML = "";
    evidence.forEach(({ key, value }) => {
      const section = document.createElement("div");
      section.className = "section";

      const title = document.createElement("h4");
      title.textContent = key;

      const content = document.createElement("p");
      content.textContent = value;

      section.appendChild(title);
      section.appendChild(content);
      contentDiv.appendChild(section);
    });

    document.getElementById("buttons").style.display = "flex";
  }); 
    document.getElementById("downloadBtn").addEventListener("click", () => {
if (!evidence || evidence.length === 0) return;

let textContent = "\n";
evidence.forEach(({ key, value }) => {
textContent += `${key}:\n${value}\n\n`;
});

const blob = new Blob([textContent], { type: "text/plain" });
const url = URL.createObjectURL(blob);
const now = new Date();
const timestamp = now.toISOString().replace(/[:.]/g, "-").slice(0, 16);
const filename = `evidence_${timestamp}.txt`;

const isSafariMac = /^((?!chrome|android).)*safari/i.test(navigator.userAgent) && navigator.platform.toUpperCase().indexOf('MAC') >= 0;

if (isSafariMac) {
// Use window.open for Safari on macOS (requires pop-ups enabled)
const newTab = window.open(url);
if (!newTab) {
  alert("Please allow pop-ups in Safari to enable the file download.");
}
} else {
// Default download method for Windows/Chrome/Edge etc.
const link = document.createElement("a");
link.href = url;
link.download = filename;
document.body.appendChild(link);
link.click();
document.body.removeChild(link);
}

URL.revokeObjectURL(url);
});

function showStatusMessage(message, type = "success") {
const statusBox = document.getElementById("status");
statusBox.textContent = message;
statusBox.className = `status ${type}`;
statusBox.style.display = "block";
}

  function dispatchEvidence(payload, apiKey) {
    //console.log(payload);
  const dispatchUrl = "https://localhost:44396/api/v2/Evidence/Dispatch";

  return fetch(dispatchUrl, {
    method: "PUT",
    headers: {
      "Accept": "text/plain",
      "Authorization": apiKey,
      "Content-Type": "application/json-patch+json"
    },
    body: JSON.stringify(payload)
  })
    .then(response => {
      if (!response.ok) {
        throw new Error(`Dispatch failed with status ${response.status}`);
      }
      return response.text();
    });
  }
  
  const dispatchBtn = document.getElementById("dispatchBtn");
  const statusBox = document.getElementById("status");

  dispatchBtn.addEventListener("click", () => {
    if (!payload) return;

    const auditContext = getAuditContextFromInputs();
    const validationError = validateAuditContext(auditContext);
    if (validationError) {
      showStatusMessage(`❌ ${validationError}`, "error");
      return;
    }

    payload.header.auditContext = auditContext;

    // Disable the button to prevent double clicks
    dispatchBtn.disabled = true;
    //dispatchBtn.textContent = "Dispatching...";
// Add spinner + "Dispatching..."
dispatchBtn.innerHTML = `
 <span class="spinner" style="margin: 0;display: inline-block;width: 12px;height: 12px;border: 2px solid #fff;border-top: 2px solid transparent;border-radius: 50%;animation: spin 0.6s linear infinite;margin-right: 8px;"></span>

Dispatching...
`;
    dispatchEvidence(payload, apiKey)
      .then(responseText => {
        try {
          const response = JSON.parse(responseText);
          const receiptId = response?.id;

          if (receiptId) {
            showStatusMessage(`✅ Evidence dispatched successfully. Receipt ID: ${receiptId}`, "success");

  dispatchBtn.style.display = "none";
          } else {
            throw new Error("Invalid response from server.");
          }
        } catch (e) {
         showStatusMessage("❌ Something went wrong. Please try again later.", "error");


                dispatchBtn.disabled = false; 


// Reset button inner HTML (restore icon and text)
dispatchBtn.innerHTML = `
<span style="display: inline-block; margin-right: 8px; transform: rotate(-90deg)">
  ⤵
</span>  
Dispatch Evidence
`;
        }
      })
      .catch(error => {
        console.error("Dispatch error:", error);
     showStatusMessage("❌ Something went wrong. Please try again later.", "error");


       dispatchBtn.disabled = false; 
// Reset button inner HTML (restore icon and text)
dispatchBtn.innerHTML = `
<span style="display: inline-block; margin-right: 8px; transform: rotate(-90deg)">
  ⤵
</span>  
Dispatch Evidence
`;
      })
      .finally(() => { 
        // Keep the button disabled after click no matter the result
                   dispatchBtn.disabled = false; 

      });
  });

}); 
