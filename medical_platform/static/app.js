// 지도 렌더링: data-points 속성의 JSON 배열(name, lat, lng, url, district)을 Leaflet 마커로 표시한다.
document.querySelectorAll("[data-map]").forEach(function (el) {
  if (typeof L === "undefined") return;
  var points = JSON.parse(el.getAttribute("data-points") || "[]");
  var map = L.map(el, { scrollWheelZoom: false }).setView([37.53, 127.0], 11);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 18,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);
  var pin = L.divIcon({ className: "", html: '<div class="pin"></div>', iconSize: [14, 14], iconAnchor: [7, 7], popupAnchor: [0, -8] });
  var bounds = [];
  points.forEach(function (p) {
    var a = document.createElement("a");
    a.href = p.url;
    a.textContent = p.name;
    var popup = document.createElement("div");
    popup.appendChild(a);
    var sub = document.createElement("div");
    sub.className = "small";
    sub.textContent = p.district || "";
    popup.appendChild(sub);
    L.marker([p.lat, p.lng], { icon: pin, title: p.name }).addTo(map).bindPopup(popup);
    bounds.push([p.lat, p.lng]);
  });
  if (bounds.length === 1) map.setView(bounds[0], 15);
  else if (bounds.length > 1) map.fitBounds(bounds, { padding: [30, 30], maxZoom: 13 });
});

// 상담 폼: 메신저 연락 방법을 고르면 메신저 ID 입력을 강조한다.
var channel = document.getElementById("contact_channel");
if (channel) {
  var toggle = function () {
    var row = document.getElementById("messenger_row");
    if (row) row.classList.toggle("highlight", channel.value && channel.value !== "email" && channel.value !== "phone");
  };
  channel.addEventListener("change", toggle);
  toggle();
}
