import {
  buildBlock,
  decorateIcons,
  decorateSections,
  decorateBlocks,
  decorateTemplateAndTheme,
  waitForFirstImage,
  loadSection,
  loadSections,
  loadCSS,
} from './aem.js';

/**
 * Builds hero block and prepends to main in a new section.
 * @param {Element} main The container element
 */
function buildHeroBlock(main) {
  const h1 = main.querySelector('h1');
  const picture = main.querySelector('picture');
  // eslint-disable-next-line no-bitwise
  if (h1 && picture && (h1.compareDocumentPosition(picture) & Node.DOCUMENT_POSITION_PRECEDING)) {
    // Check if h1 or picture is already inside a hero block
    if (h1.closest('.hero') || picture.closest('.hero')) {
      return; // Don't create a duplicate hero block
    }
    const section = document.createElement('div');
    section.append(buildBlock('hero', { elems: [picture, h1] }));
    main.prepend(section);
  }
}

/**
 * load fonts.css and set a session storage flag
 */
async function loadFonts() {
  await loadCSS(`${window.hlx.codeBasePath}/styles/fonts.css`);
  try {
    if (!window.location.hostname.includes('localhost')) sessionStorage.setItem('fonts-loaded', 'true');
  } catch (e) {
    // do nothing
  }
}

/**
 * Builds all synthetic blocks in a container element.
 * @param {Element} main The container element
 */
function buildAutoBlocks(main) {
  try {
    // auto load `*/fragments/*` references
    const fragments = [...main.querySelectorAll('a[href*="/fragments/"]')].filter((f) => !f.closest('.fragment'));
    if (fragments.length > 0) {
      // eslint-disable-next-line import/no-cycle
      import('../blocks/fragment/fragment.js').then(({ loadFragment }) => {
        fragments.forEach(async (fragment) => {
          try {
            const { pathname } = new URL(fragment.href);
            const frag = await loadFragment(pathname);
            fragment.parentElement.replaceWith(...frag.children);
          } catch (error) {
            // eslint-disable-next-line no-console
            console.error('Fragment loading failed', error);
          }
        });
      });
    }

    buildHeroBlock(main);
  } catch (error) {
    // eslint-disable-next-line no-console
    console.error('Auto Blocking failed', error);
  }
}

function getRamBlockName(wrapper) {
  if (!wrapper) return '';
  const block = wrapper.querySelector(':scope > [class*="ram-"]');
  if (!block) return '';
  return [...block.classList].find((name) => name.startsWith('ram-') && name !== 'ram') || '';
}

function getSectionClass(blockNames) {
  if (blockNames.includes('ram-footer')) {
    return 'section ram-homepage-page-footer ram-footer-container footer-container';
  }

  if (blockNames.includes('ram-header') && blockNames.includes('ram-hero')) {
    return 'section ram-header-container ram-hero-container';
  }

  if (blockNames.length === 1) {
    return `section ram-homepage-page-content ${blockNames[0]}-container`;
  }

  return `section ram-homepage-page-content ${blockNames.map((name) => `${name}-container`).join(' ')}`;
}

function normalizeRamHomepageStructure(main) {
  // In UE canvas, keep authored section nodes intact so section selection/add works reliably.
  if (main.querySelector('[data-aue-resource]')) return;

  const sections = [...main.querySelectorAll(':scope > .section')];
  if (!sections.length) return;

  const getWrappers = (section) => [...section.querySelectorAll(':scope > div')];
  const getRamWrappers = (section) => getWrappers(section)
    .filter((wrapper) => getRamBlockName(wrapper));
  const getRamBlockNames = (section) => getRamWrappers(section)
    .map((wrapper) => getRamBlockName(wrapper));
  const splitSectionByWrappers = (section) => {
    const wrappers = getWrappers(section);
    if (wrappers.length <= 1) return;

    const chunks = [];
    for (let i = 0; i < wrappers.length; i += 1) {
      const current = wrappers[i];
      const currentBlock = getRamBlockName(current);
      const next = wrappers[i + 1];
      const nextBlock = getRamBlockName(next);

      if (!currentBlock) {
        chunks.push({ className: 'section', wrappers: [current] });
        continue;
      }

      if (currentBlock === 'ram-header' && nextBlock === 'ram-hero') {
        chunks.push({
          className: getSectionClass(['ram-header', 'ram-hero']),
          wrappers: [current, next],
        });
        i += 1;
      } else {
        chunks.push({
          className: getSectionClass([currentBlock]),
          wrappers: [current],
        });
      }
    }

    if (chunks.length <= 1) return;

    chunks.forEach(({ className, wrappers: sectionWrappers }) => {
      const newSection = document.createElement('div');
      newSection.className = className;
      sectionWrappers.forEach((wrapper) => {
        newSection.appendChild(wrapper);
      });
      section.parentNode.insertBefore(newSection, section);
    });

    section.remove();
  };

  const allRamBlockNames = sections.flatMap((section) => getRamBlockNames(section));
  const hasHeader = allRamBlockNames.includes('ram-header');
  const hasHero = allRamBlockNames.includes('ram-hero');
  if (!hasHeader || !hasHero) return;

  const currentSections = [...main.querySelectorAll(':scope > .section')];
  for (let i = 0; i < currentSections.length - 1; i += 1) {
    const section = currentSections[i];
    const nextSection = currentSections[i + 1];
    const currentBlockNames = getRamBlockNames(section);
    const nextBlockNames = getRamBlockNames(nextSection);

    if (currentBlockNames.length === 1
      && nextBlockNames.length === 1
      && currentBlockNames[0] === 'ram-header'
      && nextBlockNames[0] === 'ram-hero') {
      const heroWrapper = getRamWrappers(nextSection)[0];
      if (heroWrapper) {
        section.appendChild(heroWrapper);
        nextSection.remove();
      }
    }
  }

  [...main.querySelectorAll(':scope > .section')].forEach((section) => {
    if (getRamWrappers(section).length > 1) {
      splitSectionByWrappers(section);
    }
  });

  [...main.querySelectorAll(':scope > .section')].forEach((section) => {
    const blockNames = getRamBlockNames(section);
    if (!blockNames.length) return;
    section.className = getSectionClass(blockNames);
  });
}

/**
 * Decorates formatted links to style them as buttons.
 * @param {HTMLElement} main The main container element
 */
function decorateButtons(main) {
  main.querySelectorAll('p a[href]').forEach((a) => {
    a.title = a.title || a.textContent;
    const p = a.closest('p');
    const text = a.textContent.trim();

    // quick structural checks
    if (a.querySelector('img') || p.textContent.trim() !== text) return;

    // skip URL display links
    try {
      if (new URL(a.href).href === new URL(text, window.location).href) return;
    } catch { /* continue */ }

    // require authored formatting for buttonization
    const strong = a.closest('strong');
    const em = a.closest('em');
    if (!strong && !em) return;

    p.className = 'button-wrapper';
    a.className = 'button';
    if (strong && em) { // high-impact call-to-action
      a.classList.add('accent');
      const outer = strong.contains(em) ? strong : em;
      outer.replaceWith(a);
    } else if (strong) {
      a.classList.add('primary');
      strong.replaceWith(a);
    } else {
      a.classList.add('secondary');
      em.replaceWith(a);
    }
  });
}

/**
 * Moves Universal Editor instrumentation attributes from one element to another.
 * @param {Element} source The source element containing authored attributes
 * @param {Element} target The target element that should keep instrumentation
 */
export function moveInstrumentation(source, target) {
  if (!source || !target || !source.attributes) return;

  [...source.attributes].forEach(({ name, value }) => {
    if (name.startsWith('data-aue-') || name.startsWith('data-richtext-')) {
      target.setAttribute(name, value);
    }
  });
}

/**
 * Decorates the main element.
 * @param {Element} main The main element
 */
// eslint-disable-next-line import/prefer-default-export
export function decorateMain(main) {
  decorateIcons(main);
  buildAutoBlocks(main);
  decorateSections(main);
  normalizeRamHomepageStructure(main);
  decorateBlocks(main);
  decorateButtons(main);
}

/**
 * Loads everything needed to get to LCP.
 * @param {Element} doc The container element
 */
async function loadEager(doc) {
  document.documentElement.lang = 'en';
  decorateTemplateAndTheme();
  const main = doc.querySelector('main');
  if (main) {
    decorateMain(main);
    document.body.classList.add('appear');
    const hasRamHeader = !!main.querySelector('.ram-header');
    const hasRamHero = !!main.querySelector('.ram-hero');
    document.body.classList.toggle('ram-homepage-page', hasRamHeader && hasRamHero);
    await loadSection(main.querySelector('.section'), waitForFirstImage);
  }

  try {
    /* if desktop (proxy for fast connection) or fonts already loaded, load fonts.css */
    if (window.innerWidth >= 900 || sessionStorage.getItem('fonts-loaded')) {
      loadFonts();
    }
  } catch (e) {
    // do nothing
  }
}

/**
 * Loads everything that doesn't need to be delayed.
 * @param {Element} doc The container element
 */
async function loadLazy(doc) {
  //loadHeader(doc.querySelector('header'));

  const main = doc.querySelector('main');
  await loadSections(main);

  const { hash } = window.location;
  const element = hash ? doc.getElementById(hash.substring(1)) : false;
  if (hash && element) element.scrollIntoView();

  //loadFooter(doc.querySelector('footer'));

  loadCSS(`${window.hlx.codeBasePath}/styles/lazy-styles.css`);
  loadFonts();
}

/**
 * Loads everything that happens a lot later,
 * without impacting the user experience.
 */
function loadDelayed() {
  // eslint-disable-next-line import/no-cycle
  window.setTimeout(() => import('./delayed.js'), 3000);
  // load anything that can be postponed to the latest here
}

async function loadPage() {
  await loadEager(document);
  await loadLazy(document);
  loadDelayed();
}

loadPage();
