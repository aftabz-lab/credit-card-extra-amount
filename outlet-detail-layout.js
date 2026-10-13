// Scroll behavior is scoped to Outlet Detail and does not read or modify dashboard data.
const view = document.getElementById('outlets-view');
const panel = document.getElementById('outlet-detail-panel');
const heading = panel?.querySelector('.panel-heading');
const table = document.getElementById('outlet-table');
const scrollbar = document.getElementById('outlet-detail-scrollbar');
const track = document.getElementById('outlet-detail-scroll-track');

if (view && panel && heading && table && scrollbar && track) {
  let frame = 0;
  function update() {
    frame = 0;
    if (view.hidden || !table.clientWidth) {
      scrollbar.hidden = true;
      table.style.setProperty('--outlet-detail-header-offset', '0px');
      return;
    }
    const panelRect = panel.getBoundingClientRect();
    const headingRect = heading.getBoundingClientRect();
    const tableRect = table.getBoundingClientRect();
    const offset = Math.max(0, headingRect.bottom - tableRect.top);
    table.style.setProperty('--outlet-detail-header-offset', offset + 'px');
    const frozen = panelRect.top <= 0.5 && headingRect.top >= -0.5 && headingRect.top <= 0.5;
    const overflow = table.scrollWidth > table.clientWidth + 1;
    scrollbar.hidden = !frozen || !overflow || tableRect.bottom <= headingRect.bottom;
    if (scrollbar.hidden) return;
    const left = Math.max(0, tableRect.left);
    const right = Math.min(window.innerWidth, tableRect.left + table.clientWidth);
    scrollbar.style.left = left + 'px';
    scrollbar.style.width = Math.max(0, right - left) + 'px';
    track.style.width = Math.max(0, table.scrollWidth - 2) + 'px';
    if (Math.abs(scrollbar.scrollLeft - table.scrollLeft) > 0.5) scrollbar.scrollLeft = table.scrollLeft;
  }
  function schedule() { if (!frame) frame = window.requestAnimationFrame(update); }
  table.addEventListener('scroll', () => {
    if (!scrollbar.hidden && Math.abs(scrollbar.scrollLeft - table.scrollLeft) > 0.5) scrollbar.scrollLeft = table.scrollLeft;
    schedule();
  }, { passive: true });
  scrollbar.addEventListener('scroll', () => {
    if (!scrollbar.hidden && Math.abs(table.scrollLeft - scrollbar.scrollLeft) > 0.5) table.scrollLeft = scrollbar.scrollLeft;
  }, { passive: true });
  window.addEventListener('scroll', schedule, { passive: true });
  window.addEventListener('resize', schedule, { passive: true });
  document.addEventListener('visibilitychange', schedule);
  new ResizeObserver(schedule).observe(table);
  new ResizeObserver(schedule).observe(heading);
  new MutationObserver(schedule).observe(table, { childList: true });
  new MutationObserver(schedule).observe(view, { attributes: true, attributeFilter: ['hidden'] });
  if (document.fonts?.ready) document.fonts.ready.then(schedule);
  update();
}
