/* global Office */

let generating = false;

Office.onReady(() => {
  const btn = document.getElementById("generateBtn");
  if (!btn) return;

  // Attach listener
  if (!btn.hasAttribute("data-listener")){
    btn.setAttribute("data-listener", "true")
    btn.addEventListener("click", generateEmail);
  }
});

const templates = {
  acknowledgement: {
    subject: "Request Acknowledgement",
    body: `
      <p>We acknowledge receipt of your request.</p>
      <p>The matter is currently being reviewed, and you will be updated shortly.</p>
    `
  },
  
  courseSwap: {
    subject: "Course Swap - Information Required",
    body: `
      <p>To proceed with your course swap request, please provide:</p>
      <ul>
        <li>Programme offering</li>
        <li>Lecturer</li>
        <li>Semester</li>
      </ul>
    `
  },
  
  courseSwapCompleted: {
    subject: "Course Swap Completed",
    body: `
      <p>Your course swap has been successfully completed.</p>
      <p>Please verify the update and complete the User Satisfaction Survey below:</p>
      <p>
        <a href="https://forms.office.com/r/d50uHS5dPS">
        User Satisfaction Survey
        </a>
      </p>
    `
  },

  furtherAttention: {
    subject: "Issue Requires Further Attention",
    body: `
      <p>Your request requires further investigation.</p>
      <p>The matter has been escalated to <b>INSERTNAMEHERE</b>.</p>
      <p>We will update you once more information becomes available.</p>
      `
    },
    
    iPasswordReset: {
      subject: "Password Reset Completed",
      body: `
        <p>Your password reset request has been completed.</p>
        <p>Please attempt to sign in using your updated credentials.</p>
        <p>If you experience any further issues, kindly let us know.</p>
      `
    },

  issueResolved: {
    subject: "Issue Resolved",
    body: `
    <p>Your issue has been resolved.</p>
    <p>Please complete the User Satisfaction Survey below:</p>
    <p>
    <a href="https://forms.office.com/r/d50uHS5dPS">
    User Satisfaction Survey
    </a>
    </p>
    `
  },

  iUnlock: {
    subject: "iSIMS Account Unlock",
    body: `
      <p>Your iSIMS account has been successfully unlocked.</p>
      <p>You may now attempt to log in.</p>
      <p>Please advise if you encounter any additional difficulties.</p>
    `
  },
  
  loanDevice: {
    subject: "Loan Device Availability",
    body: `
      <p>Thank you for your inquiry regarding a loan device.</p>
      <p>At this time, <b>[a device is available / no devices are currently available]</b>.</p>
      <p>You will be advised should the status change.</p>
    `
  },

  loanDeviceAck: {
    subject: "Loan Device Request Acknowledgement",
    body: `
      <p>We acknowledge receipt of your loan device request.</p>
      <p>Your request is currently being reviewed and you will be contacted with further details.</p>
    `
  },

  MFA: {
    subject: "Two-Factor Authentication Update Required",
    body: `
      <p>To proceed with the two-factor authentication (2FA) update, please provide the correct phone number to be associated with your account.</p>
      <p>Once received, we will complete the update.</p>
    `
  },

  missingGrade: {
    subject: "Missing Grade Investigation",
    body: `
      <p>We acknowledge receipt of your report regarding a missing grade.</p>
      <p>The matter is currently under investigation.</p>
      <p>You will be updated once it has been resolved.</p>
      `
    },

  paSystem: {
    subject: "PA System Setup Availability",
    body: `
    <p>The PA System setup will be available shortly.</p>
    <p>You will be notified once the setup process begins.</p>
    `
  },

  paSystemCompleted: {
    subject: "PA System Setup Completed",
    body: `
    <p>The PA System setup has been successfully completed.</p>
    <p>Please complete the User Satisfaction Survey below:</p>
    <p>
        <a href="https://forms.office.com/r/d50uHS5dPS">
        User Satisfaction Survey
        </a>
      </p>
    `
  },

  postTrainingCompleted: {
    subject: "Training Completed",
    body: `
      <p>The training session has been successfully completed.</p>
      <p>Please complete the User Satisfaction Survey below:</p>
      <p>
      <a href="https://forms.office.com/r/d50uHS5dPS">
        User Satisfaction Survey
        </a>
        </p>
      <p>Thank you for your feedback.</p>
    `
  },
  
  trainingAck: {
    subject: "Training Request Acknowledgement",
    body: `
    <p>We acknowledge receipt of your training request.</p>
    <p>The details are currently being reviewed.</p>
    <p>We will follow up with confirmation and next steps.</p>
    `
  },
  
    vPasswordReset: {
      subject: "Password Reset Completed",
      body: `
        <p>Your password reset request has been completed.</p>
        <p>Please attempt to sign in using your updated credentials.</p>
        <p>If you experience any further issues, kindly let us know.</p>
      `
    },

    vUnlock: {
      subject: "VTDI Account Unlock",
      body: `
        <p>Your VTDI account has been successfully unlocked.</p>
        <p>You may now attempt to log in.</p>
        <p>Please advise if you encounter any additional difficulties.</p>
      `
    },
};

function setStatus(msg, isError = false) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.style.color = isError ? "crimson" : "green";
}

function escapeHtml(str) {
  return (str || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
  }

function emailWrapper(content, data) {
  return `
    <p>Hello ${escapeHtml(data.clientName) || "there"},</p>
    ${content}
  `;
}

function buildTemplate(templateKey, data) {
  const tpl = templates[templateKey];

  if (!tpl) {
    throw new Error("Template not found: " + templateKey);
  }

  return {
    subject: tpl.subject + (data.refNo ? `: ${escapeHtml(data.refNo)}` : ""),
    bodyHtml: emailWrapper(tpl.body, data)
  };
}

function generateEmail() {
  if (generating) return;
  generating = true;

  try {
    const item = Office.context.mailbox.item;

    // Ensure we are in compose mode
    if (!item) {
      setStatus("Open a new email (compose) then try again.", true);
      return;
    }
    
    const templateKey = document.getElementById("templateSelect").value;
    const toEmail = document.getElementById("toEmail").value.trim();
    const toCC = document.getElementById("toCC").value.trim();
    const clientName = document.getElementById("clientName").value.trim();
    const refNo = document.getElementById("refNo").value.trim();
    const details = document.getElementById("details").value.trim();

    if (!toEmail) {
      setStatus("Please enter a recipient email address.", true);
      return;
    }

    const data = {
      toEmail,
      clientName,
      refNo,
      details,
      signatureName: "David Lynch",
    };

    const tpl = buildTemplate(templateKey, data);

    const isCompose = item?.body?.setAsync !== undefined;

    if (!isCompose) {
      // Read Mode
      Office.context.mailbox.item.displayReplyAllForm({
        htmlBody: tpl.bodyHtml
      });
      setStatus("Reply email generated. Review and send."); 
      return;
    } else {
      // Compose Mode
      item.subject.setAsync(tpl.subject);

      item.body.setAsync(
        tpl.bodyHtml,
        { coercionType: Office.CoercionType.Html }
      );

      item.to.setAsync([toEmail]);

      if (toCC) {
        item.cc.setAsync([toCC]);
      }

      setStatus("New email generated. Review and send.");
      return;
    }
  } catch (e) {
    setStatus("Error: " + (e?.stack || e), true);
  } finally {
    generating = false;
  }
}