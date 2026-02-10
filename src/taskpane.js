const HOLIDAYS = [
  "01-01", "05-01", "05-08", "07-14", "08-15",
  "11-01", "11-11", "12-25"
];

function isHoliday(date) {
  const d = ("0" + date.getDate()).slice(-2);
  const m = ("0" + (date.getMonth() + 1)).slice(-2);
  return HOLIDAYS.includes(`${m}-${d}`);
}

function loadSettings() {
  const s = Office.context.roamingSettings;
  document.getElementById("enableAddon").checked = s.get("enabled") !== false;
  document.getElementById("startTime").value = s.get("startTime") || "07:30";
  document.getElementById("endTime").value = s.get("endTime") || "19:30";
}

function saveSettings() {
  const s = Office.context.roamingSettings;
  s.set("enabled", document.getElementById("enableAddon").checked);
  s.set("startTime", document.getElementById("startTime").value);
  s.set("endTime", document.getElementById("endTime").value);
  s.saveAsync();
  alert("Paramètres sauvegardés.");
}

function isOutsideWorkingHours() {
  const s = Office.context.roamingSettings;
  if (s.get("enabled") === false) return false;

  const now = new Date();
  if (isHoliday(now)) return true;

  const [sh, sm] = (s.get("startTime") || "07:30").split(":");
  const [eh, em] = (s.get("endTime") || "19:30").split(":");

  const start = new Date(now);
  start.setHours(sh, sm, 0);

  const end = new Date(now);
  end.setHours(eh, em, 0);

  return now < start || now > end;
}

Office.onReady(() => {
  loadSettings();

  document.getElementById("saveSettings").onclick = saveSettings;

  if (!isOutsideWorkingHours()) {
    document.getElementById("alertBox").style.display = "none";
  }

  document.getElementById("btnOui").onclick = () => {
    alert("L’envoi différé sera positionné par l’utilisateur.");
    Office.context.ui.closeContainer();
  };

  document.getElementById("btnNon").onclick = () => {
    Office.context.ui.closeContainer();
  };

  document.getElementById("btnAnnuler").onclick = () => {
    Office.context.ui.closeContainer();
  };
});
