async function loadPublications() {
  const container = document.getElementById('publications-content');
  const statsContainer = document.getElementById('publication-stats');
  const controlsContainer = document.getElementById('publication-controls');
  if (!container || !statsContainer || !controlsContainer) {
    return;
  }

  container.innerHTML = '<p class="publication-loading">Loading publications…</p>';

  try {
    const response = await fetch('publications-data.json', { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`Could not load publications data (${response.status})`);
    }
    const data = await response.json();
    renderPublications(data, container, statsContainer, controlsContainer);
  } catch (error) {
    console.error(error);
    container.innerHTML = '<p class="publication-error">Unable to load the Google Scholar publication data right now. Please refresh later.</p>';
  }
}

function renderPublications(data, container, statsContainer, controlsContainer) {
  const items = Array.isArray(data.items) ? data.items : [];
  const stats = Array.isArray(data.stats) ? data.stats : [];
  const years = Array.isArray(data.years) ? data.years : [];
  const availableYears = years.map(group => group.year).filter(Boolean);
  const selectedYear = getSelectedYear(availableYears);
  const filteredItems = selectedYear === 'all'
    ? items
    : items.filter(item => item.year === Number(selectedYear));

  controlsContainer.innerHTML = '';

  const select = document.createElement('select');
  select.id = 'year-filter';
  select.innerHTML = `
    <option value="all" ${selectedYear === 'all' ? 'selected' : ''}>All years</option>
    ${availableYears.map(year => `<option value="${year}" ${String(selectedYear) === String(year) ? 'selected' : ''}>${year}</option>`).join('')}
  `;
  select.addEventListener('change', (event) => {
    const nextYear = event.target.value;
    setSelectedYear(nextYear);
    renderPublications(data, container, statsContainer, controlsContainer);
  });

  const controlsMarkup = `
    <div class="publication-controls-row">
      <div class="publication-count">Showing <strong>${filteredItems.length}</strong> of <strong>${items.length}</strong> publications</div>
    </div>
  `;
  controlsContainer.innerHTML = controlsMarkup;
  const filterWrap = document.createElement('div');
  filterWrap.className = 'publication-filter-wrap';
  filterWrap.innerHTML = `
    <label class="publication-filter" for="year-filter">
      <span>Filter by year</span>
    </label>
  `;
  filterWrap.appendChild(select);
  controlsContainer.appendChild(filterWrap);

  const effectiveStats = stats.filter(stat => !(stat.journal || '').includes('Nanyang Technological University'));
  const totalCount = effectiveStats.reduce((sum, stat) => sum + (stat.count || 0), 0);
  const statsMarkup = `
    <table class="card-table publication-stats-table">
      <thead>
        <tr>
          <th>Journal / venue</th>
          <th>Count</th>
        </tr>
      </thead>
      <tbody>
        ${effectiveStats.map(stat => `
          <tr>
            <td>${escapeHtml(`${stat.journal || 'Unknown'} ${stat.impactDisplay || ''}`.trim())}</td>
            <td>${stat.count || 0}</td>
          </tr>
        `).join('')}
        <tr class="publication-total-row">
          <td><strong>Total</strong></td>
          <td><strong>${totalCount}</strong></td>
        </tr>
      </tbody>
    </table>
  `;

  const visibleGroups = selectedYear === 'all'
    ? years
    : years.filter(group => String(group.year) === String(selectedYear));

  const yearMarkup = visibleGroups.map(group => `
    <section class="publication-year-group">
      <h2 class="publication-year-heading">${group.year}</h2>
      <div class="publication-list">
        ${group.items.filter(item => selectedYear === 'all' || item.year === Number(selectedYear)).map((item, index) => `
          <div class="publication-entry">[${index + 1}] ${escapeHtml(formatPublication(item))}</div>
        `).join('')}
      </div>
    </section>
  `).join('');

  statsContainer.innerHTML = statsMarkup;
  container.innerHTML = yearMarkup || '<p class="publication-error">No publications found for the selected year.</p>';
}

function formatPublication(item) {
  const parts = [];
  if (item.title) {
    parts.push(item.title);
  }
  if (item.authors) {
    parts.push(item.authors);
  }
  if (item.year) {
    parts.push(`(${item.year})`);
  }
  const journalText = item.journal ? `${item.journal}${item.impactDisplay ? ` ${item.impactDisplay}` : ''}` : '';
  if (journalText) {
    parts.push(journalText);
  }
  return parts.join('. ');
}

function getSelectedYear(availableYears) {
  const stored = window.sessionStorage.getItem('publication-year-filter');
  if (stored && (stored === 'all' || availableYears.includes(Number(stored)))) {
    return stored;
  }
  return 'all';
}

function setSelectedYear(value) {
  window.sessionStorage.setItem('publication-year-filter', value);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/[\"']/g, '&quot;')
    .replace(/'/g, '&#39;');
}

document.addEventListener('DOMContentLoaded', loadPublications);
