/* Auto-fill customer details handed off from the RFQ tracker.

   The tracker's "Send to automation" button opens a costing page with the
   customer details in the query string, e.g.
     /costing?company=Acme&contact=Jane&email=jane@acme.com&mobile=98765...&rfq_id=ENC/RFQ/2026/114
   We stash those in localStorage ('rfq_customer') so that WHICHEVER costing tab
   is opened next pre-fills the customer fields — even after navigating between
   tabs, which drops the query string. Edits to a customer field update the stash
   so corrections follow across tabs.

   Customer fields use two id conventions across the costing pages:
     - most pages:  c-company / c-name / c-email / c-phone
     - regen page:  company_name / poc_name / email / mobile_no
*/
(function () {
  var KEY = 'rfq_customer';
  // Attributes carried from the tracker; only the four below have form fields.
  var PARAMS = ['rfq_id', 'enquiry_id', 'company', 'contact', 'email', 'mobile', 'item', 'type'];
  var MAP = {
    company: ['c-company', 'company_name'],
    contact: ['c-name', 'poc_name'],
    email:   ['c-email', 'email'],
    mobile:  ['c-phone', 'mobile_no']
  };

  function load() { try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch (e) { return {}; } }
  function save(d) { try { localStorage.setItem(KEY, JSON.stringify(d)); } catch (e) {} }

  // 1) Capture any details arriving in the URL (a fresh hand-off from the tracker).
  var qp = new URLSearchParams(location.search);
  var stash = load();
  var gotNew = false;
  PARAMS.forEach(function (k) {
    var v = qp.get(k);
    if (v != null && v !== '') { stash[k] = v; gotNew = true; }
  });
  if (gotNew) save(stash);

  function fieldFor(key) {
    var ids = MAP[key] || [];
    for (var i = 0; i < ids.length; i++) {
      var el = document.getElementById(ids[i]);
      if (el) return el;
    }
    return null;
  }

  // 2) Fill the customer fields from the stash (overrides the pages' test defaults).
  function prefill() {
    var d = load();
    Object.keys(MAP).forEach(function (key) {
      if (!d[key]) return;
      var el = fieldFor(key);
      if (el && el.value !== d[key]) {
        el.value = d[key];
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    });
  }

  // 3) Keep the stash in sync when a customer field is edited on any tab.
  function wireWriteback() {
    Object.keys(MAP).forEach(function (key) {
      var el = fieldFor(key);
      if (!el) return;
      el.addEventListener('input', function () {
        var d = load(); d[key] = el.value; save(d);
      });
    });
  }

  function init() { prefill(); wireWriteback(); }

  // Run after the page's own DOMContentLoaded sticky-restore, so RFQ data wins.
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
