/* =====================================================================
   Nettwear IMS Work — connection setting
   ---------------------------------------------------------------------
   Apps Script -> Deploy -> Manage deployments wala URL yahan daalo
   (…/exec par khatam hota hai). Bas itna.

   Iske baad website har device par password maangegi (ek baar, device
   yaad rakhega) aur data Google Sheet se khud load karegi.
   Password Code.gs ke CONFIG.PASSWORD mein set hota hai, yahan NAHI.

   Khaali chhodoge to website purane tareeke se (sirf is browser mein,
   bina password) chalegi.
   ===================================================================== */
window.NIMS_CONFIG = {
  webAppUrl: 'https://script.google.com/macros/s/AKfycbxKcH7rql9b6H_q8Y6_XlGMKR7WppEn1a-NLAM4g62HbGdI3nJKYTlcyt9xKPAyvH-W/exec'
};
