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
  const stats = Array.isArray(data.stats) ? data.stats : [];
  const years = Array.isArray(data.years) ? data.years : [];
  const availableYears = years.map(group => group.year).filter(Boolean);
  const selectedYear = getSelectedYear(availableYears);

  // Citation numbering: entries are grouped newest year first, but the numbers
  // follow the citation order of the publication list (oldest paper = [1]), so
  // they count down from the top of the page and stay stable when filtering.
  const orderedItems = years.reduce(
    (all, group) => all.concat(Array.isArray(group.items) ? group.items : []),
    []
  );
  const citationNumbers = new Map();
  orderedItems.forEach((item, index) => {
    citationNumbers.set(item, orderedItems.length - index);
  });

  const items = orderedItems.length
    ? orderedItems
    : (Array.isArray(data.items) ? data.items : []);
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

  // Which journals appear here (and their IF/quartile) is decided in the
  // Journals sheet of publications-review.xlsx and baked into publications-data.json.
  const effectiveStats = stats.filter(stat => (stat.count || 0) > 0);
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
        ${group.items.filter(item => selectedYear === 'all' || item.year === Number(selectedYear)).map(item => `
          <div class="publication-entry">[${citationNumbers.get(item)}] ${formatPublicationHtml(item)}</div>
        `).join('')}
      </div>
    </section>
  `).join('');

  statsContainer.innerHTML = statsMarkup;
  container.innerHTML = yearMarkup || '<p class="publication-error">No publications found for the selected year.</p>';
}

// Surname particles that must stay part of the family name (van der Berg, de Silva...).
const NAME_PARTICLES = ['van', 'von', 'de', 'del', 'della', 'der', 'di', 'da', 'dos', 'du', 'la', 'le', 'bin', 'binti', 'al', 'ter', 'ten'];

// The group leader's name is shown in bold wherever it appears in an author list.
const SELF_AUTHORS = ['Jiang, K.', 'Jiang, Ke'];

// APA 7 reference as HTML: the group author is bold, the journal is underlined, e.g.
// Wang, Y., ... & Hai, L. (2025). Title of the paper. Thin-Walled Structures, 212, 113190.
function formatPublicationHtml(item) {
  const authors = highlightSelf(escapeHtml(formatAuthors(item.authors)));
  const year = `(${item.year ? escapeHtml(item.year) : 'n.d.'})`;
  const title = escapeHtml(withPeriod(item.title));
  const source = sourceInfo(item);

  const parts = [];
  if (authors) {
    parts.push(`${authors} ${year}.`);
    if (title) {
      parts.push(title);
    }
  } else if (title) {
    parts.push(`${title} ${year}.`);
  } else {
    parts.push(`${year}.`);
  }

  if (source.journal || source.extras.length) {
    const journal = source.journal
      ? `<span class="publication-journal">${escapeHtml(source.journal)}</span>`
      : '';
    const tail = source.extras.map(escapeHtml).join(', ');
    parts.push(withPeriod([journal, tail].filter(Boolean).join(', ')));
  }
  return parts.join(' ');
}

function highlightSelf(escapedAuthors) {
  return SELF_AUTHORS.reduce((markup, name) => {
    const needle = escapeHtml(name);
    return markup.split(needle).join(`<span class="publication-self">${needle}</span>`);
  }, escapedAuthors);
}

function formatAuthors(rawAuthors) {
  const pieces = String(rawAuthors || '')
    .split(',')
    .map(piece => piece.trim())
    .filter(Boolean);

  let truncated = false;
  const names = [];
  pieces.forEach(piece => {
    if (/^(\.{2,}|…)$/.test(piece)) {
      truncated = true;
      return;
    }
    names.push(formatAuthorName(piece));
  });

  if (!names.length) {
    return '';
  }
  if (truncated) {
    return `${names.join(', ')}, et al.`;
  }
  if (names.length === 1) {
    return names[0];
  }
  return `${names.slice(0, -1).join(', ')}, & ${names[names.length - 1]}`;
}

// Google Scholar writes "Y Wang" / "BY Lee"; APA needs "Wang, Y." / "Lee, B. Y.".
function formatAuthorName(rawName) {
  const tokens = String(rawName || '').split(/\s+/).filter(Boolean);
  if (tokens.length < 2) {
    return tokens.join(' ');
  }
  const initials = formatInitials(tokens[0]);
  const rest = tokens.slice(1);
  const surname = rest
    .filter((token, index) => index === rest.length - 1
      || NAME_PARTICLES.indexOf(token.toLowerCase()) !== -1
      || token.charAt(0) === token.charAt(0).toUpperCase())
    .join(' ');
  if (!surname) {
    return tokens.join(' ');
  }
  return initials ? `${surname}, ${initials}` : surname;
}

function formatInitials(token) {
  const cleaned = String(token || '').replace(/\./g, '');
  if (!cleaned) {
    return '';
  }
  if (/[a-z]/.test(cleaned)) {
    // A spelled-out given name such as "Ke" becomes "K."
    return `${cleaned.charAt(0).toUpperCase()}.`;
  }
  return cleaned
    .split(/(-)/)
    .map(part => (part === '-'
      ? '-'
      : part.split('').map(letter => `${letter.toUpperCase()}.`).join(' ')))
    .join('');
}

// Journal + volume/issue/pages. Corrected columns from publications-review.xlsx win;
// otherwise the raw Scholar venue ("Thin-Walled Structures 212, 113190 , 2025") is parsed.
function sourceInfo(item) {
  const canonical = String(item.journal || '').trim();
  const volume = String(item.volume || '').trim();
  const issue = String(item.issue || '').trim();
  const pages = String(item.pages || '').trim();
  if (canonical && (volume || pages)) {
    const extras = [];
    if (volume) {
      extras.push(issue ? `${volume}(${issue})` : volume);
    }
    if (pages) {
      extras.push(pages);
    }
    return { journal: canonical, extras };
  }

  let text = String(item.venue || '').trim().replace(/\s*,\s*(?:19|20)\d{2}\s*$/, '').trim();
  if (!text) {
    return { journal: canonical, extras: [] };
  }

  const extras = [];
  let parsedPages = '';
  const pagesMatch = text.match(/^(.*),\s*([^,]+)$/);
  if (pagesMatch && /^[\dA-Za-z]+(?:\s*[-–]\s*[\dA-Za-z]+)?$/.test(pagesMatch[2].trim())) {
    text = pagesMatch[1].trim();
    parsedPages = pagesMatch[2].trim();
  }

  let journal = text;
  const volumeMatch = text.match(/^(.*?)\s+(\d+)\s*(?:\(([^)]+)\))?$/);
  if (volumeMatch) {
    journal = volumeMatch[1].trim();
    const parsedIssue = (volumeMatch[3] || '').trim();
    extras.push(parsedIssue ? `${volumeMatch[2]}(${parsedIssue})` : volumeMatch[2]);
  }
  if (parsedPages) {
    extras.push(parsedPages);
  }
  if (canonical && canonical.toLowerCase() === journal.toLowerCase()) {
    journal = canonical;
  }
  return { journal, extras };
}

function withPeriod(text) {
  const value = String(text || '').trim();
  if (!value) {
    return '';
  }
  return /[.!?]$/.test(value) ? value : `${value}.`;
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
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

document.addEventListener('DOMContentLoaded', loadPublications);
