// Kingsway Dashboard settings. Edit the values between the quotes.
window.KINGSWAY_CONFIG = {
  // 1. From the Azure portal: App registrations > Kingsway Dashboard > Overview > "Application (client) ID".
  clientId: "PASTE-YOUR-CLIENT-ID-HERE",

  // 2. Optional: a OneDrive share link to the workbook (OneDrive > the file > Share > Copy link).
  //    With it, everyone who signs in opens the same file automatically. Leave "" to search your OneDrive.
  shareLink: "",

  // Name the app searches for in OneDrive when no share link is set.
  fileNameHint: "Kingsway",

  // Leave as "common" (works for personal and work Microsoft accounts).
  tenant: "common",

  // Optional: real start month of semester 1 for each batch, used to place the undated archive receipts.
  // Example: semesterStarts: { "1": "2024-04", "2": "2024-10", "3": "2025-10", "4": "2026-08" }
  semesterStarts: {}
};
