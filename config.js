/* =====================================================================
   Nettwear IMS Work — connection settings
   ---------------------------------------------------------------------
   Ye do values bharne ke baad website har device par Google Sheet
   (database) se data khud load karegi, aur sirf permission wale Gmail
   users hi sign in karke data dekh payenge.

   1. webAppUrl      : Apps Script -> Deploy -> Manage deployments wala
                       URL (…/exec par khatam hota hai)
   2. googleClientId : Google Cloud Console -> Credentials -> OAuth
                       Client ID (…apps.googleusercontent.com)

   Dono khaali chhodoge to website purane tareeke se (sirf is browser
   mein, bina login) chalegi.
   ===================================================================== */
window.NIMS_CONFIG = {
  webAppUrl: 'PASTE-YOUR-APPS-SCRIPT-EXEC-URL-HERE',
  googleClientId: 'PASTE-YOUR-CLIENT-ID.apps.googleusercontent.com'
};
