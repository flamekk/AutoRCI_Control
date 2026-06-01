document.addEventListener("DOMContentLoaded", () => {
  const runForm = document.getElementById("runForm");
  const runButton = document.getElementById("runButton");
  if (runForm && runButton) {
    runForm.addEventListener("submit", () => {
      runButton.disabled = true;
      runButton.textContent = "Traitement en cours...";
    });
  }
});
