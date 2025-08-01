<script>
const localChanges = {};
let currentRequestMode = '';
let pendingAdhocDocuments = [];
let selectedDocuments = {};
let activeRequestId = null;
const sessionAdhocDocs = [];
const opportunityId = getQueryParam('id');
document.getElementById('submit-request-button').style.display = 'none';
document.getElementById('cancel-request-button').style.display = 'none';

function addBusinessDays(startDate, numberOfDays) {
  const result = new Date(startDate);
  let addedDays = 0;

  while (addedDays < numberOfDays) {
    result.setDate(result.getDate() + 1);
    const day = result.getDay();
    if (day !== 0 && day !== 6) { // Skip Sunday (0) and Saturday (6)
      addedDays++;
    }
  }

  return result;
}

async function submitNewRequest() {
  const container = document.getElementById('document-template-container');
  const cbList = container.querySelectorAll('input[type="checkbox"]');

  if (!cbList.length) {
    alert("No documents selected.");
    return;
  }

  try {
    // Fetch Opportunity to get linked Contact (Assignee)
    const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
    const opportunityRes = await fetch(
  `${baseUrl}/api/data/v9.2/cr601_opportunities(${opportunityId})?$select=_cr601_contact_value`,
  { headers: { "Accept": "application/json" } }
);

    if (!opportunityRes.ok) throw new Error("Failed to fetch opportunity contact");

    const oppData = await opportunityRes.json();
    const contactId = oppData._cr601_contact_value; 
    if (!contactId) throw new Error("No contact linked to this opportunity");

    // Get logged-in CRM user
    const userSettings = parent.Xrm.Utility.getGlobalContext().userSettings;
    const userId = userSettings.userId.replace(/[{}]/g, '');

    // Construct payload
    const requestPayload = {
      "cr601_Opportunity@odata.bind": `/cr601_opportunities(${opportunityId})`,
      "cr601_requeststatus": 784160001,
      "cr601_Assignee@odata.bind": `/contacts(${contactId})`,
      "cr601_CRMUser@odata.bind": `/systemusers(${userId})`
    };

    const requestRes = await fetch(`${baseUrl}/api/data/v9.2/cr601_documentrequests`, {
      method: 'POST',
      headers: {
        "Content-Type": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Accept": "application/json"
      },
      body: JSON.stringify(requestPayload)
    });

    if (!requestRes.ok) throw new Error("Failed to create document request");

    const requestLocation = requestRes.headers.get("OData-EntityId");
    const requestId = requestLocation.match(/\(([^)]+)\)/)[1];
    if (!requestId) throw new Error("Unable to retrieve request ID");

    // 2. Loop through checkboxes and create selections
    const selectionPromises = [];
    let noteBody = "Document Request Summary:\n";

    const processedDocIds = new Set();

    for (const cb of cbList) {
      const documentTypeId = cb.value;
      const isSelected = cb.checked;
      const docName = cb.dataset.name || "Unnamed Document";

      const noteInput = container.querySelector(`input.note-input[data-doc-type-id="${documentTypeId}"]`);
      const note = noteInput ? noteInput.value.trim() : "";

      noteBody += `- ${docName} | ${isSelected ? "Selected" : "Not Selected"} | Note: "${note}"\n`;

      processedDocIds.add(documentTypeId);
      selectionPromises.push(
        createOpportunityDocumentSelection(opportunityId, documentTypeId, isSelected, requestId, note)
      );
    }
    // 3. Create selections for any pending ad-hoc docs
    for (const adhoc of pendingAdhocDocuments) {
      const documentTypeId = adhoc.documentTypeId;
      console.log("Adhoc doc id", documentTypeId);
      if (processedDocIds.has(documentTypeId)) continue; // Avoid duplicates

      const checkbox = container.querySelector(`input[type="checkbox"][value="${documentTypeId}"]`);
      const isChecked = checkbox ? checkbox.checked : false;

      const noteInput = container.querySelector(`input.note-input[data-doc-type-id="${documentTypeId}"]`);
      const note = noteInput ? noteInput.value.trim() : "";

      const docName = adhoc.docName || "Ad-hoc Document";

      noteBody += `- ${docName} | ${isChecked ? "Selected" : "Not Selected"} | Note: "${note}"\n`;

      selectionPromises.push(
        createOpportunityDocumentSelection(opportunityId, documentTypeId, isChecked, requestId, note)
      );
    }

    // Wait for all selections to complete
    await Promise.all(selectionPromises);

    // Optional: clear the pending list
    pendingAdhocDocuments = [];

    // 4. Add note
    await createNoteForOpportunity(
      opportunityId,
       noteBody.trim(),
      `New Document Request Created`
    );

    alert("New document request submitted successfully.");
    document.getElementById('document-template-container').style.display = 'none';
    document.getElementById('submit-request-button').style.display = 'none';
    document.getElementById('cancel-request-button').style.display = 'none';

    // Refresh the grid
    await fetchDocumentRequests(opportunityId);

  } catch (err) {
    console.error(err);
    alert("Error submitting new request: " + err.message);
  }
}

async function createOpportunityDocumentSelection(opportunityId, documentTypeId, isSelected = false, requestId = null,note = "") {
  const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
  const apiUrl = `${baseUrl}/api/data/v9.2/cr601_opportunitydocumentselections`;

  const body = {
    "cr601_OpportunityId@odata.bind": `/cr601_opportunities(${opportunityId})`,
    "cr601_DocumentTypeId@odata.bind": `/cr601_documenttypes(${documentTypeId})`,
    cr601_selected: isSelected,
    cr601_notes: note
  };
  if (requestId) {
    body["cr601_DocumentRequest@odata.bind"] = `/cr601_documentrequests(${requestId})`;
  }

  const response = await fetch(apiUrl, {
    method: 'POST',
    headers: {
      "Content-Type": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "Accept": "application/json"
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    throw new Error(`Failed to create document selection: ${response.statusText}`);
  }

  // Fix for empty response body
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

async function submitModifiedRequest() {
  const container = document.getElementById('modifyContainer');
  const cbList = container.querySelectorAll('input[type="checkbox"]');

  if (!cbList.length) {
    alert("No documents available to modify.");
    return;
  }

  try {
    const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
    const updatePromises = [];
    const modifiedDetails = [];

    // Get the requestId stored during startModifyRequest
    const requestId = window.activeModifyRequestId;
    if (!requestId) throw new Error("Missing request ID");

    // Fetch request number
    const reqRes = await fetch(`${baseUrl}/api/data/v9.2/cr601_documentrequests(${requestId})?$select=cr601_requestnumber`);
    const reqData = await reqRes.json();
    const requestNumber = reqData.cr601_requestnumber || requestId;

    // Fetch Opportunity Topic
    const oppRes = await fetch(`${baseUrl}/api/data/v9.2/cr601_opportunities(${opportunityId})?$select=topic`);
    const oppData = await oppRes.json();
    const opportunityTopic = oppData.topic || opportunityId;

    for (const cb of cbList) {
      const selectionId = cb.dataset.selectionId;
      const documentTypeId = cb.dataset.documentTypeId;
      const isSelected = cb.checked;
      const docName = cb.parentNode.textContent.trim();

      const noteInput = container.querySelector(`input.note-input[data-doc-type-id="${documentTypeId}"]`);
      const note = noteInput ? noteInput.value.trim() : "";

      if (!selectionId) continue;

      modifiedDetails.push(`- ${docName}: ${isSelected ? "Yes" : "No"}${note ? ` | Note: ${note}` : ''}`);

      updatePromises.push(updateDocumentSelection(selectionId, {
        cr601_selected: isSelected,
        cr601_notes: note
      }));
    }

    await Promise.all(updatePromises);

    // Create formatted annotation
    const noteText = `Modified Document Request #${requestNumber} for Opportunity "${opportunityTopic}".\n\nRequested Documents:\n${modifiedDetails.join('\n')}`;
    await createNoteForOpportunity(opportunityId, noteText, `Document Request #${requestNumber} Modified`);

    alert("Document selections updated successfully.");
    document.getElementById('document-template-container').style.display = 'none';
    document.getElementById('submit-request-button').style.display = 'none';
    document.getElementById('cancel-request-button').style.display = 'none';
    document.getElementById('modifyContainer').style.display = 'none';

    await fetchDocumentRequests(opportunityId);

  } catch (err) {
    console.error(err);
    alert("Error submitting modified request: " + err.message);
  }
}

function cancelNewRequest() {
  document.getElementById('document-template-container').style.display = 'none';
  document.getElementById('submit-request-button').style.display = 'none';
  document.getElementById('cancel-request-button').style.display = 'none';
  document.getElementById('new-request-button').disabled = false;
  document.getElementById('modifyContainer').style.display = 'none';
  pendingAdhocDocuments = [];
  selectedDocuments = {};
}
function cancelModifiedRequest() {
  document.getElementById('document-template-container').style.display = 'none';
  document.getElementById('submit-request-button').style.display = 'none';
  document.getElementById('cancel-request-button').style.display = 'none';
  document.getElementById('new-request-button').disabled = true;
  document.getElementById('modifyContainer').style.display = 'none';
  pendingAdhocDocuments = [];
  selectedDocuments = {};
}

async function fetchCategories() {
const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
const url = `${baseUrl}/api/data/v9.2/cr601_documentcategories?$select=cr601_documentcategoryid,cr601_name`;
const response = await fetch(url, { headers: { "Accept": "application/json" } });
if (!response.ok) throw new Error('Failed to fetch categories');
const data = await response.json();
return data.value; // [{ cr601_documentcategoryid, cr601_name }, ...]
}

async function fetchDocumentTypes() {
  const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
  const url = `${baseUrl}/api/data/v9.2/cr601_documenttypes` +
              `?$select=cr601_documenttypeid,cr601_name,cr601_isadhoc` + // include isadhoc
              `&$expand=cr601_DocumentCategoryId($select=cr601_name,cr601_documentcategoryid)`;

  const response = await fetch(url, {
    headers: {
      "Accept": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0"
    }
  });

  if (!response.ok) {
    throw new Error('Failed to fetch document types');
  }

  const data = await response.json();
  return data.value;
}

function renderTemplate(categories, selectedIds = [], documentList = []) {
  const container = document.getElementById('document-template-container');
  container.innerHTML = '';

  const lowerSelected = selectedIds.map(id => id.toLowerCase());

  categories.forEach(category => {
    const div = document.createElement('div');
    div.classList.add('category');

    const h3 = document.createElement('h3');
    const headingText = document.createElement('span');
    headingText.textContent = category.cr601_name;
    h3.appendChild(headingText);

    // Add "+" button for 'Others' category
    if (category.cr601_name.toLowerCase() === 'others') {
      const addBtn = document.createElement('button');
      addBtn.textContent = '+';
      addBtn.className = 'plus-btn';
      addBtn.style.marginLeft = '10px';
      addBtn.title = 'Add ad-hoc document';
      addBtn.onclick = () => showAddDocumentPrompt(category.cr601_name);

      h3.appendChild(addBtn);
    }

    div.appendChild(h3);

    category.documentTypes.forEach(dt => {
      const docId = dt.cr601_documenttypeid;
      const docName = dt.cr601_name;

      const row = document.createElement('div');
      row.classList.add('document-row');

      const labelContainer = document.createElement('div');
      labelContainer.classList.add('label-container');

      const label = document.createElement('label');
      label.classList.add('checkbox-label');

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = docId;
      checkbox.dataset.name = docName;

      // Determine initial selection state
      const isChecked = selectedDocuments[docId]?.selected ?? lowerSelected.includes(docId.toLowerCase());
      checkbox.checked = isChecked;

      // Save to selectedDocuments
      selectedDocuments[docId] = selectedDocuments[docId] || {};
      selectedDocuments[docId].selected = isChecked;

      // Create text node and label
      const text = document.createTextNode(docName);
      label.appendChild(checkbox);
      label.appendChild(text);
      labelContainer.appendChild(label);

      // Create note input
      const noteInput = document.createElement('input');
      noteInput.type = 'text';
      noteInput.classList.add('note-input');
      noteInput.dataset.docTypeId = docId;
      noteInput.placeholder = 'Enter notes here...';

      // Initial note value
      const noteVal = selectedDocuments[docId]?.notes || '';
      noteInput.value = noteVal;

      noteInput.style.display = isChecked ? 'inline-block' : 'none';

      // Event: checkbox toggled
      checkbox.addEventListener('change', () => {
        selectedDocuments[docId].selected = checkbox.checked;

        if (checkbox.checked) {
          noteInput.style.display = 'inline-block';
        } else {
          noteInput.value = '';
          noteInput.style.display = 'none';
          selectedDocuments[docId].notes = '';
        }
      });

      // Event: notes input changed
      noteInput.addEventListener('input', () => {
        selectedDocuments[docId].notes = noteInput.value;
      });

      // Add to DOM
      row.appendChild(labelContainer);
      row.appendChild(noteInput);
      div.appendChild(row);
    });

    container.appendChild(div);
  });
}

async function getCategoriesWithTypes() {
  const [categories, documentTypes] = await Promise.all([
    fetchCategories(),
    fetchDocumentTypes()
  ]);

  const groupedByCategory = {};

  documentTypes.forEach(dt => {
    const isAdhoc = dt.cr601_isadhoc === true;
    const category = dt.cr601_DocumentCategoryId;
    const categoryId = category?.cr601_documentcategoryid;

    // Skip unbound types
    if (!categoryId) return;
    // NEW LOGIC
    if (isAdhoc) {
      if (currentRequestMode === 'new') {
        // Show ad-hoc only if manually added during this session
        const isPending = pendingAdhocDocuments.some(p => p.documentTypeId.toLowerCase() === dt.cr601_documenttypeid.toLowerCase());
        if (!isPending) return;
      }
      // In modify mode, show all adhoc types for the request (they're in CRM)
      // So we allow them
    }

    if (!groupedByCategory[categoryId]) {
      groupedByCategory[categoryId] = {
        cr601_documentcategoryid: categoryId,
        cr601_name: category.cr601_name,
        documentTypes: []
      };
    }

    groupedByCategory[categoryId].documentTypes.push(dt);
  });

  // Ensure "Others" category is present
  const othersCategory = categories.find(cat => cat.cr601_name.toLowerCase() === 'others');
  if (othersCategory && !groupedByCategory[othersCategory.cr601_documentcategoryid]) {
    groupedByCategory[othersCategory.cr601_documentcategoryid] = {
      cr601_documentcategoryid: othersCategory.cr601_documentcategoryid,
      cr601_name: othersCategory.cr601_name,
      documentTypes: []
    };
  }

  return Object.values(groupedByCategory);
}

async function startNewRequest() {
  document.getElementById('new-request-button').disabled = true;
  currentRequestMode = 'new';

  // Show template, hide checklist
  document.getElementById('document-template-container').style.display = 'block';

  // Show submit + cancel
  document.getElementById('submit-request-button').style.display = 'inline-block';
  document.getElementById('submit-request-button').onclick = submitNewRequest;
  document.getElementById('cancel-request-button').style.display = 'inline-block';
  document.getElementById('cancel-request-button').onclick = cancelNewRequest;

  try {
    const categories = await getCategoriesWithTypes();
    renderTemplate(categories);
  } catch (e) {
    alert('Error loading template: ' + e.message);
    document.getElementById('new-request-button').disabled = false;
  }
}

async function updateDocumentSelection(id, fieldsToUpdate) {
const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
const apiUrl = `${baseUrl}/api/data/v9.2/cr601_opportunitydocumentselections(${id})`;
const response = await fetch(apiUrl, {
    method: 'PATCH',
    headers: {
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Accept": "application/json"
    },
    body: JSON.stringify(fieldsToUpdate)
});

if (!response.ok) {
    throw new Error(`Failed to update document selection: ${response.statusText}`);
}
}

async function createNoteForOpportunity(opportunityId, noteText, subject) {
const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
const apiUrl = `${baseUrl}/api/data/v9.2/annotations`;

const response = await fetch(apiUrl, {
    method: 'POST',
    headers: {
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Accept": "application/json"
    },
    body: JSON.stringify({
    notetext: noteText,
    subject: subject,
    "objectid_cr601_opportunity@odata.bind": `/cr601_opportunities(${opportunityId})`
    })
});

if (!response.ok) {
    throw new Error(`Failed to create note: ${response.statusText}`);
}
}

function getQueryParam(name) {
const urlParams = new URLSearchParams(window.location.search);
console.log(urlParams);
return urlParams.get(name);
}

async function fetchDocumentRequests(opportunityId) {
const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
const url = `${baseUrl}/api/data/v9.2/cr601_documentrequests` +
  `?$filter=_cr601_opportunity_value eq '${opportunityId}'` +
  `&$select=cr601_requestnumber,cr601_requestduedate,createdon,cr601_requeststatus,` +
  `cr601_colloborator,cr601_documentprogresscomplete,cr601_documentrequestid` +
  `&$expand=cr601_CRMUser($select=fullname),cr601_Assignee($select=fullname)` +
  `&$orderby=createdon desc`;


try {
const response = await fetch(url, {
    headers: {
    "Accept": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"'
    }
});

if (!response.ok) throw new Error("Failed to fetch document requests");

const data = await response.json();
renderRequestGrid(data.value);

// Enable New Request button only if all statuses are not "Document Collection"
const hasInProgress = data.value.some(req => {
const status = req["cr601_requeststatus@OData.Community.Display.V1.FormattedValue"];
return status === 'Document Collection';
});

document.getElementById('new-request-button').disabled = hasInProgress;

} catch (error) {
console.error(error);
document.getElementById('document-request-grid').textContent = "Error loading document requests.";
}
}

function renderRequestGrid(requests) {
  const container = document.getElementById('document-request-grid');
  if (!requests.length) {
    container.textContent = "No document requests available.";
    return;
  }

  const table = document.createElement('table');
  table.classList.add('document-request-table');

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  const headers = [
    'Request #', 'Created On', 'Due Date', 'Status', 
    'CRM User', 'Assignee', 'Collaborator', 'Progress %', 'Action'
  ];

  headers.forEach(text => {
    const th = document.createElement('th');
    th.textContent = text;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');

  requests.forEach(req => {
    const row = document.createElement('tr');

    const getText = (val) => val ?? '';
    const statusText = req["cr601_requeststatus@OData.Community.Display.V1.FormattedValue"] || req.cr601_requeststatus;
    const crmUser = req.cr601_CRMUser?.fullname || '';
    const assignee = req.cr601_Assignee?.fullname || '';
    const cells = [
      getText(req.cr601_requestnumber),
      new Date(req.createdon).toLocaleString(),
      req.cr601_requestduedate ? new Date(req.cr601_requestduedate).toLocaleDateString() : '',
      statusText,
      crmUser,
      assignee,
      getText(req.cr601_colloborator),
      req.cr601_documentprogresscomplete !== null ? `${req.cr601_documentprogresscomplete}%` : ''
    ];

    cells.forEach(cellText => {
      const td = document.createElement('td');
      td.textContent = cellText;
      row.appendChild(td);
    });

    const actionCell = document.createElement('td');
    const modifyBtn = document.createElement('button');
    modifyBtn.textContent = 'Modify';

    const isDisabled = statusText !== 'Document Collection';
    modifyBtn.disabled = isDisabled;

    modifyBtn.onclick = () => startModifyRequest(opportunityId, req.cr601_requestnumber);

    actionCell.appendChild(modifyBtn);
    row.appendChild(actionCell);

    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  container.innerHTML = '';
  container.appendChild(table);
}

async function startModifyRequest(opportunityId, requestNumber) {
  
  document.getElementById('new-request-button').disabled = true;
  currentRequestMode = 'modify';
    // Show submit + cancel
  document.getElementById('submit-request-button').style.display = 'inline-block';
  document.getElementById('submit-request-button').onclick = submitModifiedRequest;
  document.getElementById('cancel-request-button').style.display = 'inline-block';
  document.getElementById('cancel-request-button').onclick = cancelModifiedRequest;
  try {
    const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
    const apiUrl = `${baseUrl}/api/data/v9.2`;

    // 1. Fetch Document Request Id by requestNumber
    const requestRes = await fetch(`${apiUrl}/cr601_documentrequests?$filter=cr601_requestnumber eq '${requestNumber}'&$select=cr601_documentrequestid`);
    const requestData = await requestRes.json();
    if (!requestData.value.length) throw new Error('Document request not found');
    const documentRequestId = requestData.value[0].cr601_documentrequestid;
    window.activeModifyRequestId = documentRequestId;
    // 2. Fetch all selections with document type & category for opportunity and this request
    const selectionUrl = `${apiUrl}/cr601_opportunitydocumentselections` +
      `?$expand=cr601_DocumentTypeId($select=cr601_name;$expand=cr601_DocumentCategoryId($select=cr601_name))` +
      `&$filter=_cr601_opportunityid_value eq '${opportunityId}' and _cr601_documentrequest_value eq '${documentRequestId}'`+
      `&$select=cr601_opportunitydocumentselectionid,cr601_selected,cr601_notes,_cr601_documenttypeid_value`;

    const selRes = await fetch(selectionUrl, {
      headers: { Accept: 'application/json' }
    });

    const selData = await selRes.json();

    // 3. Group by category and prepare array
    const groupedByCategory = {};

    selData.value.forEach(sel => {
      const docType = sel.cr601_DocumentTypeId;
      const category = docType.cr601_DocumentCategoryId?.cr601_name || 'Uncategorized';

      if (!groupedByCategory[category]) {
        groupedByCategory[category] = [];
      }

      groupedByCategory[category].push({
        cr601_documenttypeid: docType.cr601_documenttypeid,
        cr601_name: docType.cr601_name,
        cr601_selected: sel.cr601_selected,
        cr601_opportunitydocumentselectionid: sel.cr601_opportunitydocumentselectionid,
        cr601_notes: sel.cr601_notes || ''
      });
    });

    // Convert to array for rendering
    const groupedArray = Object.keys(groupedByCategory).map(category => ({
      cr601_name: category,
      documentTypes: groupedByCategory[category]
    }));
    renderModifyTemplate(groupedArray);

  } catch (err) {
    console.error(err);
    alert('Failed to load document selections for modification.');
  }
}

function renderModifyTemplate(categories) {
  const container = document.getElementById('modifyContainer');
  container.innerHTML = '';

  categories.forEach(category => {
    const categoryDiv = document.createElement('div');
    const header = document.createElement('h3');
    header.textContent = category.cr601_name;
    // Add "+" button for 'Others' category
    if (category.cr601_name.toLowerCase() === 'others') {
      const addModifyBtn = document.createElement('button');
      addModifyBtn.textContent = '+';
      addModifyBtn.className = 'plus-btn';
      addModifyBtn.style.marginLeft = '10px';
      addModifyBtn.title = 'Add ad-hoc document';
      addModifyBtn.onclick = () => showAddDocumentPrompt(category.cr601_name);

      header.appendChild(addModifyBtn);
    }

    categoryDiv.appendChild(header);

    category.documentTypes.forEach(doc => {
      // Create a flex container for each line
      const lineDiv = document.createElement('div');
      lineDiv.classList.add('modify-line');

      // Checkbox + label container
      const checkboxLabelDiv = document.createElement('div');
      checkboxLabelDiv.classList.add('checkbox-label-container');

      const label = document.createElement('label');
      label.classList.add('checkbox-label');

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = doc.cr601_selected;
      checkbox.value = doc.cr601_documenttypeid;
      checkbox.dataset.selectionId = doc.cr601_opportunitydocumentselectionid || '';
      checkbox.dataset.documentTypeId = doc.cr601_documenttypeid;

      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(' ' + doc.cr601_name));
      checkboxLabelDiv.appendChild(label);

      // Notes input container
      const notesInput = document.createElement('input');
      notesInput.type = 'text';
      notesInput.classList.add('note-input');
      notesInput.dataset.docTypeId = doc.cr601_documenttypeid;
      notesInput.placeholder = 'Enter notes here...';
      notesInput.value = doc.cr601_notes || '';
      if (doc.cr601_selected) {
        notesInput.style.display = 'inline-block';
      }

      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          notesInput.style.display = 'inline-block';
        } else {
          notesInput.value = '';
          notesInput.style.display = 'none';
        }
      });

      lineDiv.appendChild(checkboxLabelDiv);
      lineDiv.appendChild(notesInput);

      categoryDiv.appendChild(lineDiv);
    });

    container.appendChild(categoryDiv);
  });

  container.style.display = 'block';
}

async function createDocumentType(docName, categoryId) {
const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
const apiUrl = `${baseUrl}/api/data/v9.2/cr601_documenttypes`;

const newDocType = {
    cr601_name: docName,
    "cr601_DocumentCategoryId@odata.bind": `/cr601_documentcategories(${categoryId})`,
    cr601_isadhoc: true
};

const response = await fetch(apiUrl, {
    method: 'POST',
    headers: {
    "Content-Type": "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    "Accept": "application/json"
    },
    body: JSON.stringify(newDocType)
});

if (!response.ok) {
    throw new Error(`Failed to create document type: ${response.statusText}`);
}

const entityUrl = response.headers.get('OData-EntityId');
const guidMatch = entityUrl.match(/\(([^)]+)\)/);
if (!guidMatch) throw new Error("Failed to get new document type ID");

    const docTypeId = guidMatch[1];
return docTypeId;
}

async function addAdhocDocument(docName, categoryName) {
  try {
    const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();

    // 1. Fetch category ID
    const categoryResponse = await fetch(
      `${baseUrl}/api/data/v9.2/cr601_documentcategories?$filter=cr601_name eq '${categoryName.replace(/'/g, "''")}'&$select=cr601_documentcategoryid`,
      { headers: { "Accept": "application/json" } }
    );

    if (!categoryResponse.ok) throw new Error("Failed to fetch category ID");

    const categoryData = await categoryResponse.json();
    if (!categoryData.value.length) throw new Error("Category not found");

    const categoryId = categoryData.value[0].cr601_documentcategoryid;

    // 2. Create ad-hoc document type
    const documentTypeId = await createDocumentType(docName, categoryId);

    // 3. Add to pending ad-hoc documents
    pendingAdhocDocuments.push({
      documentTypeId,
      docName,
      categoryName
    });

    if (currentRequestMode === 'new') {
    // 4. Refresh UI
    await refreshDocumentUI();
  } else if (currentRequestMode === 'modify') {
    const container = document.getElementById('modifyContainer');
    const selectionPromises = [];
    let noteBody = "Document Request Summary:\n";
    // Get the requestId stored during startModifyRequest
    const requestId = window.activeModifyRequestId;

    // 1. Create Opportunity Document Selection Entry

    for (const adhoc of pendingAdhocDocuments) {
      const documentTypeId = adhoc.documentTypeId;
      console.log("Adhoc doc id", documentTypeId);

      const checkbox = container.querySelector(`input[type="checkbox"][value="${documentTypeId}"]`);
      const isChecked = checkbox ? checkbox.checked : false;

      const noteInput = container.querySelector(`input.note-input[data-doc-type-id="${documentTypeId}"]`);
      const note = noteInput ? noteInput.value.trim() : "";

      const docName = adhoc.docName || "Ad-hoc Document";

      noteBody += `- ${docName} | ${isChecked ? "Selected" : "Not Selected"} | Note: "${note}"\n`;

      selectionPromises.push(
        createOpportunityDocumentSelection(opportunityId, documentTypeId, isChecked, requestId, note)
      );
    }

    // Wait for all selections to complete
    await Promise.all(selectionPromises);

    // clear the pending list
    pendingAdhocDocuments = [];

    // Add note
    await createNoteForOpportunity(
      opportunityId,
       noteBody.trim(),
      `New Adhoc Document Created for Request - ${requestId}`
    );
  }
  // Refresh UI
    await refreshDocumentUI();

  } catch (error) {
    alert('Error adding document: ' + error.message);
  }
}

async function refreshDocumentUI() {
  const selections = await fetchAdhocDocumentSelections();

  // Merge
  const mergedSelections = selections.map(doc => {
    const localState = selectedDocuments[doc.id];
    return {
      ...doc,
      selected: localState?.selected ?? doc.selected,
      notes: localState?.notes ?? doc.notes
    };
  });

  // Update selectedDocuments with merged values
  mergedSelections.forEach(doc => {
    selectedDocuments[doc.id] = {
      selected: doc.selected,
      notes: doc.notes
    };
  });

  if (currentRequestMode === 'new') {
    const categories = await getCategoriesWithTypes();
    const selectedIds = [
      ...mergedSelections.filter(d => d.selected).map(d => d.id),
      ...pendingAdhocDocuments.map(p => p.documentTypeId)
    ];

    renderTemplate(categories, selectedIds, mergedSelections); 
  } else if (currentRequestMode === 'modify') {
    // Group selections by category name
    const groupedByCategory = {};
    console.log(selections);

    selections.forEach(doc => {
      const catName = doc.category?.name || 'Uncategorized';

      if (!groupedByCategory[catName]) {
        groupedByCategory[catName] = {
          cr601_name: catName,
          documentTypes: []
        };
      }

      groupedByCategory[catName].documentTypes.push({
        cr601_documenttypeid: doc.id,
        cr601_name: doc.name,
        cr601_notes: doc.notes,
        cr601_selected: doc.selected,
        cr601_opportunitydocumentselectionid: doc.id
      });
    });

    const categories = Object.values(groupedByCategory);
    renderModifyTemplate(categories);
    return;
  }
}

async function fetchAdhocDocumentSelections() {
  const baseUrl = parent.Xrm.Utility.getGlobalContext().getClientUrl();
  let filters = [`_cr601_opportunityid_value eq '${opportunityId}'`];

  console.log(currentRequestMode);
  
  if (currentRequestMode === 'modify') {
    const requestId = window.activeModifyRequestId;
    console.log(requestId);
    if(requestId)
  {
    filters.push(`_cr601_documentrequest_value eq '${requestId}'`);
  }
    
  }
  console.log(filters);
  const url = `${baseUrl}/api/data/v9.2/cr601_opportunitydocumentselections` +
    `?$expand=cr601_DocumentTypeId(` +
      `$select=cr601_name,cr601_isadhoc;` +
      `$expand=cr601_DocumentCategoryId($select=cr601_name)` +
    `)` +
    `&$filter=${filters.join(' and ')}`;
  /*const url = `${baseUrl}/api/data/v9.2/cr601_opportunitydocumentselections` +
    `?$expand=cr601_DocumentTypeId(` +
      `$select=cr601_name,cr601_isadhoc;` +
      `$expand=cr601_DocumentCategoryId($select=cr601_name)` +
    `)` +
    `&$filter=_cr601_opportunityid_value eq '${opportunityId}'`;*/

  const response = await fetch(url, {
    headers: {
      "Accept": "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0"
    }
  });

  if (!response.ok) throw new Error("Failed to fetch document selections");

  const data = await response.json();

  return data.value
    .filter(item => item.cr601_DocumentTypeId)
    .map(item => {
      const dt = item.cr601_DocumentTypeId;
      const catName = dt?.cr601_DocumentCategoryId?.cr601_name || "Uncategorized";
      return {
        id: item.cr601_opportunitydocumentselectionid,
        name: dt?.cr601_name || "Unnamed",
        category: { name: catName },
        selected: item.cr601_selected ?? false,
        notes: item.cr601_notes || "",
        isAdhoc: dt?.cr601_isadhoc || false
      };
    });
}

function showAddDocumentPrompt(categoryName) {
  let sessionAdhocDocs = [];
  const docName = prompt(`Add a new document for category "${categoryName}":`);

  if (!docName || !docName.trim()) {
    return; // user cancelled or empty input
  }

  const trimmedName = docName.trim();

  // Check if document with same name already exists in the same category
  const allInputs = document.querySelectorAll('.doc-label-wrapper label');
  const exists = Array.from(allInputs).some(label =>
    label.textContent.trim().toLowerCase() === trimmedName.toLowerCase() &&
    label.closest('.category')?.querySelector('h3')?.textContent?.startsWith(categoryName)
  );

  if (exists) {
    alert(`A document named "${trimmedName}" already exists under "${categoryName}".`);
    return;
  }
  const adhocDoc = {
      docName: trimmedName,
      category: categoryName
    };

  sessionAdhocDocs.push(adhocDoc);
  // Proceed to create the document
  addAdhocDocument(trimmedName, categoryName);
}

(async function () {
const container = document.getElementById('document-template-container');

if (!opportunityId) {
    container.textContent = "Opportunity ID not found in URL.";
    return;
}
// await fetchAndRenderProgress(opportunityId);
await fetchDocumentRequests(opportunityId);
})();
</script>