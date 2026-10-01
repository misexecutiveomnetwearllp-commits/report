/* =====================================================================
   Nettwear IMS Work — connection setting
   ---------------------------------------------------------------------
   Paste the Apps Script URL from Deploy -> Manage deployments here
   (it ends in …/exec). That is all.

   The website then asks for the password on each device (once; the
   device remembers it) and loads the data from the Google Sheet.
   The password is set in CONFIG.PASSWORD in Code.gs, NOT here.

   Leave it empty and the website works the old way (this browser
   only, no password).
   ===================================================================== */
window.NIMS_CONFIG = {
  webAppUrl: 'https://script.google.com/macros/s/AKfycbxKcH7rql9b6H_q8Y6_XlGMKR7WppEn1a-NLAM4g62HbGdI3nJKYTlcyt9xKPAyvH-W/exec'
};
