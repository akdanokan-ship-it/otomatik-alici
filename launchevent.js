/* global Office */
// Yeni mail / Yanıtla / Tümünü yanıtla / İlet anında otomatik çalışır.
// Kurallar kurallar.json dosyasında tutulur.

var KURAL_ADRESI = "https://akdanokan-ship-it.github.io/otomatik-alici/kurallar.json";
var TESHIS = true; // sorun çözülünce false yapın

// Yanıtta alıntılanan orijinal mailin başladığı yeri gösteren işaretler
var ALINTI_ISARETLERI = [
  /<div[^>]*id=["']?appendonsend["']?[^>]*>/i,
  /<div[^>]*id=["']?divRplyFwdMsg["']?[^>]*>/i,
  /<hr[^>]*display:\s*inline-block[^>]*>/i,
  /<div[^>]*border-top:\s*solid\s+#(E1E1E1|B5C4DF)[^>]*>/i,
  /<div[^>]*class=["']?gmail_quote/i,
  /<blockquote/i
];

function kucuk(s) { return (s || "").replace(/\s+/g, " ").trim().toLocaleLowerCase("tr-TR"); }
function bekle(ms) { return new Promise(function (c) { setTimeout(c, ms); }); }

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

// Yanıt alıcıları bazı istemcilerde olaydan hemen sonra dolar; kısa süre tekrar dene
function toGetir(item, deneme) {
  return asenkron(function (cb) { item.to.getAsync(cb); }).then(function (to) {
    if ((to && to.length) || deneme >= 5) return to || [];
    return bekle(400).then(function () { return toGetir(item, deneme + 1); });
  });
}

function alintiBaslangici(html) {
  var enKucuk = -1;
  ALINTI_ISARETLERI.forEach(function (d) {
    var m = d.exec(html);
    if (m && (enKucuk === -1 || m.index < enKucuk)) enKucuk = m.index;
  });
  return enKucuk;
}

function htmlMetin(html) {
  return html.replace(/<style[\s\S]*?<\/style>/gi, " ")
             .replace(/<[^>]+>/g, " ")
             .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&").replace(/&nbsp;/g, " ");
}

// Yanıtın içindeki orijinal mail başlığını (Kimden/Kime/Bilgi/Konu) metin olarak döndürür
function orijinalBaslik(item, deneme) {
  return asenkron(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }).then(function (html) {
    html = html || "";
    var k = alintiBaslangici(html);
    if (k === -1) {
      if (deneme >= 5) return "";
      return bekle(400).then(function () { return orijinalBaslik(item, deneme + 1); });
    }
    var metin = kucuk(htmlMetin(html.substring(k)));
    // Başlık "Konu/Subject" etiketinde biter (konu metni dahil edilmez); bulunamazsa ilk 1000 karakter
    var konu = metin.search(/(subject|konu)\s*:/);
    return konu > -1 ? metin.substring(0, konu) : metin.substring(0, 1000);
  });
}

// Adres orijinal mail başlığında (Kimden/Kime/Bilgi) geçiyor mu? Sadece e-posta adresiyle kontrol edilir.
function baslikta(adres, baslik) {
  return !!baslik && baslik.indexOf(kucuk(adres)) > -1;
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

function teshis(item, mesaj) {
  console.log("[OtomatikAlici] " + mesaj);
  if (!TESHIS) return;
  try {
    item.notificationMessages.replaceAsync("teshis", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: ("Teşhis | " + mesaj).substring(0, 150),
      icon: "Icon.16",
      persistent: true
    });
  } catch (e) {}
}

function onNewMessageComposeHandler(event) {
  var item = Office.context.mailbox.item;

  Promise.all([
    asenkron(function (cb) { item.getComposeTypeAsync(cb); }),
    toGetir(item, 0),
    asenkron(function (cb) { item.cc.getAsync(cb); }),
    kurallariGetir()
  ]).then(function (s) {
    var tur = s[0] ? s[0].composeType : null;          // "newMail" | "reply" | "forward"
    var to = s[1] || [], cc = s[2] || [], ayar = s[3] || {};
    var olay = tur === "reply" ? "yanit" : tur === "forward" ? "ilet" : "yeni";
    var gonderen = olay === "yanit" && to.length ? kucuk(to[0].emailAddress) : "";
    var yanitlayan = kucuk(Office.context.mailbox.userProfile.emailAddress);

    var uyanlar = (ayar.kurallar || []).filter(function (k) { return kuralUyarMi(k, olay, gonderen, yanitlayan); });
    if (!uyanlar.length) { teshis(item, "tür: " + tur + " | gönderen: " + (gonderen || "-") + " | uyan kural yok"); return; }

    // Yanıtlarda orijinal maildeki kişileri de kontrol et
    var baslikIsi = olay === "yanit" ? orijinalBaslik(item, 0) : Promise.resolve("");
    return baslikIsi.then(function (baslik) {
      var mevcut = {};
      to.concat(cc).forEach(function (r) { mevcut[kucuk(r.emailAddress)] = true; });

      var eklenecekTo = [], eklenecekCc = [], atlanan = [];
      function ekle(liste, a) {
        var k = kucuk(a);
        if (mevcut[k]) return;
        if (baslikta(a, baslik)) { atlanan.push(a); mevcut[k] = true; return; }
        mevcut[k] = true; liste.push(a);
      }
      uyanlar.forEach(function (k) {
        var e = k.ekle || {};
        (e.to || []).forEach(function (a) { ekle(eklenecekTo, a); });
        (e.cc || []).forEach(function (a) { ekle(eklenecekCc, a); });
      });

      teshis(item, "gönderen: " + gonderen + " | başlık: " + (baslik ? "okundu" : "YOK") +
        " | eklenecek: " + (eklenecekTo.concat(eklenecekCc).join(", ") || "yok") +
        " | orijinalde var: " + (atlanan.join(", ") || "-"));

      var isler = [];
      if (eklenecekTo.length) isler.push(asenkron(function (cb) { item.to.addAsync(eklenecekTo, cb); }));
      if (eklenecekCc.length) isler.push(asenkron(function (cb) { item.cc.addAsync(eklenecekCc, cb); }));
      return Promise.all(isler);
    });
  }).then(function () { event.completed(); }, function (e) { console.log("[OtomatikAlici] hata", e); event.completed(); });
}

// Manifest'teki FunctionName ile eşleştirme (klasik Outlook için zorunlu)
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
