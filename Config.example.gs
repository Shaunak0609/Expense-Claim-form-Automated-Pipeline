/**
 * Settings for the ECF pipeline: template copy.
 * In Apps Script, add a file named "Config" with this content and fill in your values.
 */

// ---------------------------------------------------------------------------
// CONFIG — fill these in once
// ---------------------------------------------------------------------------
const CONFIG = {
  TEMPLATE_FILE_ID: 'PASTE_ECF_TEMPLATE_XLSX_FILE_ID',   // blank ECF .xlsx in Drive
  OUTPUT_FOLDER_ID: 'PASTE_OUTPUT_FOLDER_ID',  // where filled ECFs are saved

  RESPONSES_SHEET: 'Reimbursment Request',
  LOG_SHEET: 'ECF Log',

  // Leave '' to leave the cell blank on the ECF.
  TEAM_NAME: '',
  WORKTAG: '',  // e.g. 'PM016212 EDT-Smart City' (exact text from the ECF list)

  // Drive file ID of the team's designated approver signature (PNG/JPG), added to every ECF.
  // Leave '' to leave the approver box empty.
  APPROVER_SIGNATURE_FILE_ID: '',

  // Printed under "Print Name" in the team's designated approver box on every ECF.
  APPROVER_NAME: '',

  STATUS_TO_PROCESS: 'Request Made',
  STATUS_DONE: 'ECF Generated',
  STATUS_FAILED: 'Needs Info',

  // The 2026-27 ECF's Expense Category list, spelled exactly as in its dropdown.
  VALID_CATEGORIES: [
    'Prototype Project',
    'Competition - People Costs',
    'Competition - Physical Logistics',
    'Professional Development',
    'Operational/Administrative Expenses',
  ],

  // Form wording → ECF category. Remove an entry once the form uses the ECF wording.
  // A plain "Competition" answer is rejected (Needs Info): the ECF now splits it into
  // People Costs and Physical Logistics, so update the form's options to match.
  CATEGORY_ALIASES: {
    'Prototyping Project': 'Prototype Project',
    'Administrative Expenses': 'Operational/Administrative Expenses',
  },
};
