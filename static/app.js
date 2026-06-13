// ── Mobile sidebar toggle ──────────────────────
const sidebar = document.getElementById('sidebar');
const overlay = document.getElementById('sidebarOverlay');
const menuBtn = document.getElementById('menuToggle');

if (menuBtn) {
  menuBtn.addEventListener('click', () => {
    sidebar.classList.toggle('open');
    overlay.classList.toggle('show');
  });
}
if (overlay) {
  overlay.addEventListener('click', () => {
    sidebar.classList.remove('open');
    overlay.classList.remove('show');
  });
}

// ── Auto-dismiss alerts after 4 seconds ────────
document.querySelectorAll('.alert.fade.show').forEach(el => {
  setTimeout(() => {
    const bsAlert = bootstrap.Alert.getOrCreateInstance(el);
    if (bsAlert) bsAlert.close();
  }, 4000);
});

// ── Checkbox card visual feedback ─────────────
document.querySelectorAll('.checkbox-card input[type=checkbox]').forEach(cb => {
  cb.addEventListener('change', function () {
    this.closest('.checkbox-card').classList.toggle('checked', this.checked);
  });
});
