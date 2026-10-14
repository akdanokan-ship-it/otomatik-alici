/* global Office */
// Yeni mail / Yanıtla / Tümünü yanıtla / İlet anında otomatik çalışır.
// Kurallar kurallar.json dosyasında tutulur.

var KURAL_ADRESI = "https://akdanokan-ship-it.github.io/bileti-kapat-addin/otomatik-alici/kurallar.json";

function kucuk(s) { return (s || "").trim().toLowerCase(); }

function asenkron(fn) {
  return new Promise(function (coz) {
    fn(function (r) { coz(r && r.status === Office.AsyncResultStatus.Succeeded ? r.value : null); });
  });
}

function kurallariGetir() {
  return fetch(KURAL_ADRESI + "?t=" + Date.now(), { cache: "no-store" })
    .then(function (r) { return r.ok ? r.json() : { kurallar: [] }; })
    .catch(function () { return { kurallar: [] }; });
}

function kuralUyarMi(kural, olay, gonderen, yanitlayan) {
  if (kural.olaylar && kural.olaylar.indexOf(olay) === -1) return false;

  var gonderenListesi = (kural.gonderen || []).map(kucuk);
  var alanAdlari = (kural.gonderenAlanAdi || []).map(kucuk);
  if (gonderenListesi.length || alanAdlari.length) {
    if (!gonderen) return false;
    var alanAdi = gonderen.split("@")[1] || "";
    if (gonderenListesi.indexOf(gonderen) === -1 && alanAdlari.indexOf(alanAdi) === -1) return false;
  }

  var yanitlayanListesi = (kural.yanitlayan || []).map(kucuk);
  if (yanitlayanListesi.length && yanitlayanListesi.indexOf(yanitlayan) === -1) return false;

  return true;
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  Promise.all([
    asenkron(function (cb) { item.getComposeTypeAsync(cb); }),
    asenkron(function (cb) { item.to.getAsync(cb); }),
    asenkron(function (cb) { item.cc.getAsync(cb); }),
    kurallariGetir()
  ]).then(function (s) {
    var tur = s[0] ? s[0].composeType : null;          // "newMail" | "reply" | "forward"
    var to = s[1] || [], cc = s[2] || [], ayar = s[3] || {};
    var olay = tur === "reply" ? "yanit" : tur === "forward" ? "ilet" : "yeni";

    // Yanıtta orijinal gönderen TO'nun ilk sırasındadır
    var gonderen = olay === "yanit" && to.length ? kucuk(to[0].emailAddress) : "";
    var yanitlayan = kucuk(Office.context.mailbox.userProfile.emailAddress);

    var mevcut = {};
    to.concat(cc).forEach(function (r) { mevcut[kucuk(r.emailAddress)] = true; });

    var eklenecekTo = [], eklenecekCc = [];
    (ayar.kurallar || []).forEach(function (k) {
      if (!kuralUyarMi(k, olay, gonderen, yanitlayan)) return;
      var ekle = k.ekle || {};
      (ekle.to || []).forEach(function (a) { if (!mevcut[kucuk(a)]) { mevcut[kucuk(a)] = true; eklenecekTo.push(a); } });
      (ekle.cc || []).forEach(function (a) { if (!mevcut[kucuk(a)]) { mevcut[kucuk(a)] = true; eklenecekCc.push(a); } });
    });

    var isler = [];
    if (eklenecekTo.length) isler.push(asenkron(function (cb) { item.to.addAsync(eklenecekTo, cb); }));
    if (eklenecekCc.length) isler.push(asenkron(function (cb) { item.cc.addAsync(eklenecekCc, cb); }));
    return Promise.all(isler);
  }).then(function () { event.completed(); }, function () { event.completed(); });
}

// Manifest'teki FunctionName ile eşleştirme (klasik Outlook için zorunlu)
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
