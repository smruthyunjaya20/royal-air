/* eslint-disable no-console */
const path = require('path');
const PptxGenJS = require('pptxgenjs');

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Royal Air Engineering Team';
pptx.company = 'Royal Air';
pptx.subject = 'Royal Air EDS + AEM + Universal Editor Architecture';
pptx.title = 'Royal Air Architecture - Reference Style';
pptx.lang = 'en-US';

const C = {
  bg: 'E8E8E8',
  black: '000000',
  bluePanel: '156082',
  grayBox: 'A3A3A3',
  link: '365F73',
};

const FONT_TITLE = 'Aptos Display';
const FONT_BODY = 'Aptos';

function baseSlide() {
  const slide = pptx.addSlide();
  slide.background = { color: C.bg };
  return slide;
}

function addTitle(slide, title, centered = false) {
  slide.addText(title, {
    x: centered ? 0.6 : 1.0,
    y: centered ? 0.55 : 0.35,
    w: centered ? 12.1 : 10.5,
    h: centered ? 1.2 : 0.8,
    align: centered ? 'center' : 'left',
    fontFace: FONT_TITLE,
    fontSize: centered ? 41 : 31,
    bold: false,
    color: C.black,
    valign: 'mid',
  });
}

function addBulletList(slide, items, opts = {}) {
  const runs = items.map((t) => ({
    text: t,
    options: { bullet: { indent: 18 } },
  }));

  slide.addText(runs, {
    x: opts.x || 1.0,
    y: opts.y || 1.9,
    w: opts.w || 11.2,
    h: opts.h || 4.8,
    fontFace: FONT_BODY,
    fontSize: opts.fontSize || 24,
    color: C.black,
    paraSpaceAfterPt: opts.spacing || 12,
    breakLine: true,
    valign: 'top',
  });
}

function addParagraph(slide, text, opts = {}) {
  slide.addText(text, {
    x: opts.x || 1.0,
    y: opts.y || 1.9,
    w: opts.w || 11.2,
    h: opts.h || 1.0,
    fontFace: FONT_BODY,
    fontSize: opts.fontSize || 22,
    color: C.black,
    valign: 'top',
    breakLine: true,
  });
}

function addHeading(slide, text, opts = {}) {
  slide.addText(text, {
    x: opts.x || 1.0,
    y: opts.y || 1.4,
    w: opts.w || 6.0,
    h: opts.h || 0.45,
    fontFace: FONT_BODY,
    fontSize: opts.fontSize || 22,
    bold: true,
    color: C.black,
  });
}

function addSlide1() {
  const slide = baseSlide();

  addTitle(slide, 'Royal Air Web Platform\nArchitecture', true);

  addHeading(slide, 'Problem Statement', { x: 1.8, y: 2.75, fontSize: 20 });
  addParagraph(
    slide,
    'Royal Air requires a scalable and maintainable web architecture where business teams can publish content quickly without waiting for frequent frontend code releases.',
    { x: 1.8, y: 3.15, w: 9.8, h: 0.95, fontSize: 19 },
  );

  addParagraph(
    slide,
    'As the number of campaigns, routes, and customer journeys grows, a block-based architecture with strict authoring governance is required to preserve speed and consistency.',
    { x: 1.8, y: 4.1, w: 9.8, h: 1.0, fontSize: 19 },
  );

  addHeading(slide, 'Challenges', { x: 1.8, y: 5.15, fontSize: 20 });
  addBulletList(
    slide,
    [
      'Manual dependency on frontend teams for many content updates',
      'Need for reusable UI patterns across multiple pages and campaigns',
      'Risk of inconsistent authoring without component-level governance',
      'Performance pressure for first-load and mobile experience',
    ],
    { x: 1.75, y: 5.52, w: 10.2, h: 1.8, fontSize: 17, spacing: 6 },
  );
}

function addSlide2() {
  const slide = baseSlide();

  addTitle(slide, 'Business Use Case');

  addBulletList(
    slide,
    [
      'Royal Air operates a hybrid web model: static marketing pages and dynamic transactional pages.',
      'Static pages (offers, destination content, service information) are created with EDS blocks for high-speed delivery and easy content iteration.',
      'Dynamic journeys (login, flight information, booking) are implemented as a React SPA rendered on AEM Pages.',
      'Header and footer are maintained as shared Content Fragments (CF) to ensure brand and navigation consistency across both runtimes.',
    ],
    { x: 1.0, y: 1.75, w: 11.4, h: 2.85, fontSize: 21, spacing: 7 },
  );

  addHeading(slide, 'Benefits', { x: 1.0, y: 4.25, fontSize: 26 });
  addBulletList(
    slide,
    [
      'Fast publishing for static pages without full application deployment',
      'React SPA flexibility for complex stateful transactional flows',
      'Single source for shared header/footer through Content Fragments',
      'Clear separation of concerns: content velocity vs application logic',
      'Consistent UX and governance across EDS and AEM Pages',
    ],
    { x: 1.0, y: 4.72, w: 11.4, h: 2.4, fontSize: 22, spacing: 8 },
  );
}

function addSlide3() {
  const slide = baseSlide();

  addTitle(slide, 'Solution');

  addParagraph(
    slide,
    'Royal Air uses a dual-rendering architecture where the page type determines the runtime path.',
    { x: 1.0, y: 1.55, w: 11.2, h: 0.75, fontSize: 23 },
  );

  addParagraph(
    slide,
    'Static pages are delivered through EDS using reusable blocks and progressive loading. This path is optimized for speed and editorial agility.',
    { x: 1.0, y: 2.55, w: 11.2, h: 1.0, fontSize: 23 },
  );

  addParagraph(
    slide,
    'Dynamic pages such as login, flight information, and booking are served by AEM Pages and rendered by a React SPA for rich interaction and state management.',
    { x: 1.0, y: 3.9, w: 11.2, h: 1.0, fontSize: 23 },
  );

  addParagraph(
    slide,
    'Shared header and footer Content Fragments are consumed by both channels to keep navigation, branding, and global messaging synchronized.',
    { x: 1.0, y: 5.2, w: 11.2, h: 0.95, fontSize: 23 },
  );
}

function addRoundedPanel(slide, cfg) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: cfg.x,
    y: cfg.y,
    w: cfg.w,
    h: cfg.h,
    radius: 0.1,
    fill: { color: C.bluePanel },
    line: { color: '0E2841', pt: 1.2 },
  });

  slide.addText(cfg.title, {
    x: cfg.x + 0.18,
    y: cfg.titleY || (cfg.y + 0.25),
    w: cfg.w - 0.36,
    h: cfg.titleH || 0.85,
    align: 'center',
    valign: 'mid',
    fontFace: FONT_BODY,
    fontSize: cfg.fontSize || 21,
    color: cfg.whiteText ? 'FFFFFF' : C.black,
    bold: false,
    breakLine: true,
  });
}

function addSlide4Architecture() {
  const slide = baseSlide();

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.2,
    y: 0.12,
    w: 2.95,
    h: 0.28,
    fill: { color: C.black },
    line: { color: C.black },
  });
  slide.addText('Architecture overview', {
    x: 0.45,
    y: 0.16,
    w: 2.45,
    h: 0.18,
    fontFace: FONT_BODY,
    fontSize: 15,
    color: 'FFFFFF',
  });

  addRoundedPanel(slide, {
    x: 0.55,
    y: 1.02,
    w: 2.75,
    h: 5.7,
    title: 'Authoring\n(AEM + UE)',
    titleH: 1.2,
    fontSize: 18,
  });

  addRoundedPanel(slide, {
    x: 3.65,
    y: 1.02,
    w: 2.35,
    h: 2.35,
    title: 'Static Pages\n(EDS Blocks)',
    whiteText: true,
    titleH: 1.1,
    fontSize: 17,
  });

  addRoundedPanel(slide, {
    x: 3.65,
    y: 4.35,
    w: 2.35,
    h: 2.35,
    title: 'Dynamic Pages\n(AEM Pages +\nReact SPA)',
    whiteText: true,
    titleH: 1.4,
    fontSize: 16,
  });

  addRoundedPanel(slide, {
    x: 6.35,
    y: 1.02,
    w: 2.55,
    h: 5.7,
    title: 'Shared Content\nFragments\n(Header/Footer)',
    whiteText: true,
    titleH: 1.4,
    fontSize: 17,
  });

  addRoundedPanel(slide, {
    x: 9.2,
    y: 1.02,
    w: 3.4,
    h: 2.75,
    title: 'Delivery Runtime\nEDS Live\n(.aem.page / .aem.live)',
    whiteText: true,
    titleH: 1.4,
    fontSize: 17,
  });

  addRoundedPanel(slide, {
    x: 9.2,
    y: 3.95,
    w: 3.4,
    h: 2.75,
    title: 'Delivery Runtime\nAEM Pages +\nReact SPA',
    whiteText: true,
    titleH: 1.4,
    fontSize: 17,
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 6.95,
    y: 3.28,
    w: 1.35,
    h: 1.2,
    fill: { color: C.grayBox },
    line: { color: C.black, pt: 1 },
  });
  slide.addText('CF\nAPI', {
    x: 7.05,
    y: 3.58,
    w: 1.15,
    h: 0.7,
    align: 'center',
    fontFace: FONT_BODY,
    fontSize: 16,
    color: C.black,
    breakLine: true,
  });

  slide.addShape(pptx.ShapeType.line, {
    x: 3.3,
    y: 2.1,
    w: 0.35,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 3.3,
    y: 5.45,
    w: 0.35,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });

  slide.addShape(pptx.ShapeType.line, {
    x: 6.02,
    y: 2.2,
    w: 0.33,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 6.02,
    y: 5.55,
    w: 0.33,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });

  slide.addShape(pptx.ShapeType.line, {
    x: 8.9,
    y: 2.45,
    w: 0.3,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 8.9,
    y: 5.35,
    w: 0.3,
    h: 0,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });

  slide.addShape(pptx.ShapeType.line, {
    x: 8.3,
    y: 3.9,
    w: 0.9,
    h: -1.1,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });
  slide.addShape(pptx.ShapeType.line, {
    x: 8.3,
    y: 3.9,
    w: 0.9,
    h: 1.15,
    line: { color: C.black, pt: 1.2, endArrowType: 'triangle' },
  });
}

function addUrlText(slide, y, label, url, urlHeight = 0.8) {
  slide.addText(label, {
    x: 1.0,
    y,
    w: 11.2,
    h: 0.35,
    fontFace: FONT_BODY,
    fontSize: 24,
    bold: true,
    color: C.black,
  });
  slide.addText(url, {
    x: 1.0,
    y: y + 0.55,
    w: 11.2,
    h: urlHeight,
    fontFace: FONT_BODY,
    fontSize: 22,
    color: C.link,
    underline: { color: C.link, pt: 1 },
    breakLine: true,
  });
}

function addSlide5DemoUrls() {
  const slide = baseSlide();

  addTitle(slide, 'Demo URLs');

  addUrlText(
    slide,
    1.55,
    'Live Site:',
    'https://main--royal-air--<owner>.aem.live/',
    0.52,
  );

  addUrlText(
    slide,
    2.55,
    'Feature Preview:',
    'https://<branch>--royal-air--<owner>.aem.page/',
    0.52,
  );

  addUrlText(
    slide,
    3.55,
    'Repository:',
    'https://github.com/<owner>/royal-air',
    0.52,
  );

  addUrlText(
    slide,
    4.55,
    'Universal Editor Governance Files:',
    'component-definition.json\ncomponent-models.json\ncomponent-filters.json',
    1.7,
  );
}

function addSlide6Achievements() {
  const slide = baseSlide();

  addTitle(slide, 'Achievements');

  addBulletList(
    slide,
    [
      'Implemented hybrid architecture: EDS for static pages and AEM React SPA for dynamic pages',
      'Defined page-type routing model for login, flight information, and booking flows',
      'Established shared CF-based header/footer used consistently across both runtimes',
      'Enabled high-velocity editorial updates on static content through EDS blocks',
      'Preserved application flexibility for transactional journeys via React SPA on AEM Pages',
    ],
    { x: 1.0, y: 1.85, w: 11.1, h: 4.2, fontSize: 24, spacing: 12 },
  );

  addHeading(slide, 'Next Step', { x: 1.0, y: 6.25, fontSize: 24 });
  addParagraph(
    slide,
    'Finalize CF contract versioning and define runtime ownership SLAs for static and dynamic channels.',
    { x: 1.0, y: 6.62, w: 11.2, h: 0.45, fontSize: 22 },
  );
}

function buildDeck() {
  addSlide1();
  addSlide2();
  addSlide3();
  addSlide4Architecture();
  addSlide5DemoUrls();
  addSlide6Achievements();
}

buildDeck();

const output = path.join(process.cwd(), 'Royal-Air-EDS-AEM-Architecture-Ref-Format.pptx');

pptx.writeFile({ fileName: output })
  .then(() => {
    console.log(`Presentation created: ${output}`);
  })
  .catch((err) => {
    console.error('Failed to generate presentation', err);
    process.exitCode = 1;
  });
