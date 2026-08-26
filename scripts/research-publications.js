// Fills the publication list of every research topic card (R-*.html) from
// publications-data.json, using the "category" column of publications-review.xlsx.
// A card keeps its hand-written content when the data has no paper for that topic.
async function loadResearchPublications() {
  const cards = Array.from(document.querySelectorAll('.paper-card .box-title'));
  if (!cards.length || typeof formatPublicationHtml !== 'function') {
    return;
  }

  let data;
  try {
    const response = await fetch('publications-data.json', { cache: 'no-store' });
    if (!response.ok) {
      throw new Error(`Could not load publications data (${response.status})`);
    }
    data = await response.json();
  } catch (error) {
    console.error(error);
    return;
  }

  const items = Array.isArray(data.items) ? data.items : [];
  const heading = document.querySelector('.content1 h1');
  const area = heading ? heading.textContent.trim() : '';

  cards.forEach(card => {
    const topicHeading = card.querySelector('h2');
    if (!topicHeading) {
      return;
    }
    const topic = topicHeading.textContent.trim();
    const matches = items
      .filter(item => sameText(item.area, area) && sameText(item.topic, topic))
      .sort((a, b) => (b.year || 0) - (a.year || 0));
    if (!matches.length) {
      return;
    }

    card.querySelectorAll('p').forEach(paragraph => paragraph.remove());
    matches.forEach((item, index) => {
      const paragraph = document.createElement('p');
      paragraph.innerHTML = `[${matches.length - index}] ${formatPublicationHtml(item)}`;
      card.appendChild(paragraph);
    });
  });
}

function sameText(left, right) {
  return String(left || '').trim().toLowerCase() === String(right || '').trim().toLowerCase();
}

document.addEventListener('DOMContentLoaded', loadResearchPublications);
