const map = L.map('map').setView([13.038964, 80.044928], 18);

const lightTiles = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png');
const darkTiles = L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png');
lightTiles.addTo(map);

const locations = [
  { name: "Main Gate 🚪", coords: [13.037939797229118, 80.0452551894564], image: "maingate.jpg" },
  { name: "RIT Bus Parking 🅿", coords: [13.037852391118504, 80.04511565955022], image: "park.jpg" },
  { name: "ICICI Bank 🏦", coords: [13.038147971068447, 80.04523574819733], image: "icici.jpg" },
  { name: "Rajalakshmi Schools of Business 🏢", coords: [13.037989451440975, 80.04482821543021], image: "rsb.jpg" },
  { name: "RIT College 🏫", coords: [13.038274469351844, 80.04538046365047], image: "rit.jpg" },
  { name: "B Block 🏢", coords: [13.03845173206024, 80.04535058010259], image: "bblock.jpg" },
  { name: "C Block 🏢", coords: [13.040002447255425, 80.04539913363176], image: "cblock.jpg" },
  { name: "Steve Jobs Block 🧠", coords: [13.039914313192293, 80.04485074417923], image: "steve.jpg" },
  { name: "Canteen 🍽", coords: [13.039836369406265, 80.04498005893849], image: "canteen.jpg" },
  { name: "Mess 🍛", coords: [13.040223292467353, 80.04535802567622], image: "mess.jpg" },
  { name: "College Ground 🏏", coords: [13.0409904213723, 80.04476883592335], image: "ground.jpg" },
  { name: "Girls Hostel 🏨", coords: [13.040987522569676, 80.04382158303477], image: "hostel.jpg" },
  { name: "Lake View Point 🌊", coords: [13.039152884043148, 80.04597088194272], image: "lake.jpg" }
];

const datalist = document.getElementById("locationList");
const markers = {};
let routingControl = null;


const addedNames = new Set();

locations.forEach(loc => {
  
  const marker = L.marker(loc.coords).addTo(map).bindPopup(`
    <b>${loc.name}</b><br>
    <img src="/static/images/${loc.image}" width="150" />
  `);
  markers[loc.name.toLowerCase()] = marker;

  
  if (!addedNames.has(loc.name)) {
    const option = document.createElement('option');
    option.value = loc.name;
    datalist.appendChild(option);
    addedNames.add(loc.name);
  }
});

document.getElementById("searchBox").addEventListener("input", function () {
  const value = this.value.trim().toLowerCase();
  const loc = locations.find(l => l.name.toLowerCase() === value);

  if (loc) {
    map.setView(loc.coords, 19);
    markers[loc.name.toLowerCase()].openPopup();

    map.locate();
    map.once('locationfound', e => {
      const userLocation = e.latlng;

      if (routingControl) {
        map.removeControl(routingControl);
      }

      routingControl = L.Routing.control({
        waypoints: [userLocation, L.latLng(loc.coords)],
        routeWhileDragging: false,
        addWaypoints: false,
        createMarker: () => null
      }).addTo(map);
    });

    document.querySelector("input[name='location']").value = loc.name;
  }
});

document.getElementById("clearRouteBtn").addEventListener("click", function () {
  if (routingControl) {
    map.removeControl(routingControl);
    routingControl = null;
  }
});


map.locate({ setView: true, maxZoom: 18 });
map.on('locationfound', e => {
  L.marker(e.latlng).addTo(map).bindPopup("📍 You are here").openPopup();
  L.circle(e.latlng, { radius: e.accuracy }).addTo(map);
});


document.getElementById("toggleThemeBtn").addEventListener("click", () => {
  document.body.classList.toggle("dark-mode");
  document.body.classList.toggle("light-mode");

  const isDark = document.body.classList.contains("dark-mode");
  if (isDark) {
    map.removeLayer(lightTiles);
    darkTiles.addTo(map);
    document.getElementById("toggleThemeBtn").textContent = "☀";
  } else {
    map.removeLayer(darkTiles);
    lightTiles.addTo(map);
    document.getElementById("toggleThemeBtn").textContent = "🌙";
  }
});


function submitFeedback() {
  const name = document.querySelector("input[name='name']").value;
  const location = document.querySelector("input[name='location']").value.trim();
  const comment = document.querySelector("textarea[name='comment']").value.trim();
  const category = document.getElementById("feedbackCategory")?.value.trim() || "General";
  const anonymous = document.getElementById("anonymous")?.checked || false;

  if (!location || !comment || !category) {
    alert("Please fill in all fields.");
    return;
  }

  const feedbackData = {
    name: anonymous ? "Anonymous" : name,
    location: location,
    comment: anonymous ? "Anonymous feedback submitted." : comment,
    category: category,
    timestamp: new Date().toLocaleString()
  };

  
  const oldData = JSON.parse(localStorage.getItem("feedback")) || [];
  oldData.push(feedbackData);
  localStorage.setItem("feedback", JSON.stringify(oldData));

  
  fetch('/submit_feedback', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(feedbackData)
  })
    .then(res => res.json())
    .then(data => {
      alert("✅ Feedback submitted!");
      displayFeedback(data.feedbacks || []);
      document.querySelector("form").reset();
    });
}

function displayFeedback(feedbacks) {
  const fbDiv = document.getElementById("feedbacks");
  fbDiv.innerHTML = "";
  feedbacks.forEach(item => {
    if (item.location && item.category && item.comment) {
      fbDiv.innerHTML += `
        <div class="feedback-item">
          <p><strong>${item.location}</strong> (${item.category}): ${item.comment}</p>
        </div>`;
    }
  });
}

window.onload = () => {
  fetch('/get_feedbacks')
    .then(res => res.json())
    .then(data => displayFeedback(data));
};

const feedbackForm = document.querySelector("form[action='/submit_feedback']");
if (feedbackForm) {
  feedbackForm.addEventListener("submit", function (e) {
    e.preventDefault();
    submitFeedback();
  });
}


function startVoiceSearch() {
  const recognition = new (window.SpeechRecognition || window.webkitSpeechRecognition)();
  recognition.lang = "en-US";
  recognition.start();

  recognition.onresult = function (event) {
    const transcript = event.results[0][0].transcript;
    document.getElementById("searchBox").value = transcript;
    document.getElementById("searchBox").dispatchEvent(new Event("input"));
  };
}
