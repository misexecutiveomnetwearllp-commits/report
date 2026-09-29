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
  webAppUrl: 'https://script.google.com/macros/s/AKfycbxKcH7rql9b6H_q8Y6_XlGMKR7WppEn1a-NLAM4g62HbGdI3nJKYTlcyt9xKPAyvH-W/exec',
  googleClientId: 'project-420b7f09-282e-4e96-aeb'
};
