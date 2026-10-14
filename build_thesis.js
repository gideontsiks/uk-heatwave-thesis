// build_thesis.js — Generate Thesis_Revised.docx with embedded figures.
// Target: 45-50 pages, 16 inline figures, expanded prose throughout.

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak, TableOfContents,
  NumberFormat,
} = require('docx');
const fs = require('fs');
const path = require('path');

// ============== Helpers ==============
const FONT = 'Calibri';
const SERIF = 'Cambria';
// All paths are resolved relative to this script's own folder, so the code
// is portable and carries no machine-specific absolute paths.
const FIG_DIR = path.join(__dirname, 'thesis_figures');

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240 },
    pageBreakBefore: true,
    children: [new TextRun({ text, bold: true, font: SERIF, size: 32 })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, font: SERIF, size: 26 })],
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, bold: true, italics: true, font: SERIF, size: 22 })],
  });
}

function p(text, opts = {}) {
  const runs = typeof text === 'string'
    ? [new TextRun({ text, font: FONT, size: 22 })]
    : text.map(r => new TextRun({
        text: r.text, bold: r.bold, italics: r.italics,
        font: r.mono ? 'Consolas' : FONT, size: 22,
      }));
  return new Paragraph({
    spacing: { line: 360, after: 160, ...(opts.spacing || {}) },
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    indent: opts.indent,
    children: runs,
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: 'bullets', level },
    spacing: { line: 320, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22 })],
  });
}

function refEntry(runs) {
  const arr = runs.map(r => {
    if (r.url) {
      return new ExternalHyperlink({
        link: r.url,
        children: [new TextRun({
          text: r.text, font: FONT, size: 22, style: 'Hyperlink',
          color: '0563C1',
        })],
      });
    }
    return new TextRun({
      text: r.text, bold: r.bold || false, italics: r.italics || false,
      font: FONT, size: 22,
    });
  });
  return new Paragraph({
    spacing: { line: 360, after: 60 },
    indent: { left: 720, hanging: 720 },
    alignment: AlignmentType.LEFT,
    children: arr,
  });
}

// Image embedding: scale to fit content width (6.5" = 9360 DXA = ~6240 EMU/inch * 6.5)
// EMUs: 914400 = 1 inch. So 6.5" = 5,943,600 EMU.
function fig(filename, captionText, widthInches = 6.5) {
  const path = `${FIG_DIR}/${filename}`;
  const buf = fs.readFileSync(path);
  // Read PNG dimensions
  const w = buf.readUInt32BE(16);
  const h = buf.readUInt32BE(20);
  const widthPx = w; const heightPx = h;
  const aspectRatio = heightPx / widthPx;
  const widthPt = widthInches * 72;  // 6.5 inches in points
  const heightPt = widthPt * aspectRatio;

  const imgPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new ImageRun({
      data: buf,
      type: 'png',
      transformation: {
        width: widthPt,
        height: heightPt,
      },
    })],
  });
  const captionPara = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 240, line: 280 },
    children: [
      new TextRun({ text: captionText.figNum + '. ', bold: true, font: FONT, size: 20 }),
      new TextRun({ text: captionText.title, font: FONT, size: 20 }),
      ...(captionText.note ? [new TextRun({ text: ' ' + captionText.note, italics: true, font: FONT, size: 20, color: '555555' })] : [])
    ],
  });
  return [imgPara, captionPara];
}

function table(headers, rows, caption = null) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: '999999' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const totalWidth = 9360;
  const colWidth = Math.floor(totalWidth / headers.length);
  const colWidths = headers.map(() => colWidth);

  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map(h => new TableCell({
      borders,
      width: { size: colWidth, type: WidthType.DXA },
      shading: { fill: 'EEEEEE', type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text: h, bold: true, font: FONT, size: 20 })],
      })],
    })),
  });

  const dataRows = rows.map(r => new TableRow({
    children: r.map(c => new TableCell({
      borders,
      width: { size: colWidth, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({
        children: [new TextRun({ text: String(c), font: FONT, size: 20 })],
      })],
    })),
  }));

  const tableObj = new Table({
    width: { size: totalWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows],
  });

  const items = [tableObj];
  if (caption) {
    items.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 240, line: 280 },
      children: [
        new TextRun({ text: caption.tableNum + '. ', bold: true, font: FONT, size: 20 }),
        new TextRun({ text: caption.title, font: FONT, size: 20 }),
        ...(caption.note ? [new TextRun({ text: ' ' + caption.note, italics: true, font: FONT, size: 20, color: '555555' })] : [])
      ],
    }));
  }
  return items;
}

function center(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: opts.after || 160 },
    children: [new TextRun({ text, font: opts.font || FONT, size: opts.size || 22,
                              bold: opts.bold, italics: opts.italics })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// =================================================================
// CONTENT
// =================================================================
const children = [];

// ---- Title page ----
children.push(
  new Paragraph({ spacing: { before: 1200 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'FEDERAL STATE AUTONOMOUS EDUCATIONAL INSTITUTION FOR HIGHER EDUCATION',
                              bold: true, font: SERIF, size: 22 })] }),
  center('NATIONAL RESEARCH UNIVERSITY', { bold: true, font: SERIF }),
  center('HIGHER SCHOOL OF ECONOMICS', { bold: true, font: SERIF, after: 240 }),
  center('Faculty of Social Sciences', { font: SERIF, after: 200 }),

  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 260 },
    children: [new ImageRun({
      data: fs.readFileSync(path.join(__dirname, 'hse_logo.png')),
      type: 'png',
      transformation: { width: 54, height: 69 },
    })] }),

  center('GIDEON KWESI TSIKUDO', { bold: true, font: SERIF, size: 26, after: 300 }),
  center('Master\u2019s Thesis', { italics: true, font: SERIF, size: 24, after: 440 }),

  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 480 },
    children: [new TextRun({
      text: 'SPATIOTEMPORAL ANALYSIS OF HEATWAVE-RELATED MORTALITY IN THE UNITED KINGDOM (ENGLAND AND WALES)',
      bold: true, font: SERIF, size: 30,
    })] }),

  center('field of study 38.04.04 Public Administration', { font: FONT, size: 22 }),
  center('Master\u2019s program: Population and Development', { italics: true, font: FONT, size: 22, after: 720 }),

  (function sigTable() {
    const noB = { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
                  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } };
    function sig(role, name) {
      return new TableCell({
        borders: noB,
        width: { size: 3120, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [
          new Paragraph({ children: [new TextRun({ text: role, font: FONT, size: 21 })] }),
          new Paragraph({ children: [new TextRun({ text: 'Candidate of Sciences (PhD)',
                                                    italics: true, font: FONT, size: 21 })] }),
          new Paragraph({ children: [new TextRun({ text: name, bold: true, font: FONT, size: 21 })] }),
        ],
      });
    }
    return new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [3120, 3120, 3120],
      rows: [new TableRow({ children: [
        sig('Reviewer', 'EZEKIEL NII NOYE NORTEY'),
        sig('Scientific Supervisor', 'ERMOLIN ILYA V.'),
        sig('Scientific Supervisor', 'GLEB SMIDOVICH SERGEEVICH'),
      ] })],
    });
  })(),

  new Paragraph({ spacing: { before: 1000 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Moscow, 2024',
                              italics: true, font: SERIF, size: 22 })] }),
  pageBreak(),
);

// Boundary: title page ends here, front matter (roman numerals) begins.
const FRONT_MATTER_START = children.length;

// ---- Abstract ----
children.push(
  h1('Abstract'),
  p('This study conducts a spatiotemporal analysis of heatwave-related mortality in England and Wales over the period 2014\u20132022, testing two hypotheses: that regions with lower public green-space provision per capita (H1) and regions with higher building density (H2) experience higher heat-related mortality. The empirical strategy operates in two analytical frames, a temporal frame implementing a distributed-lag non-linear model (DLNM) on national daily mortality, and a spatial frame applying global Moran\u2019s I, local indicators of spatial association (LISA), the Getis-Ord Gi* statistic, bivariate Moran\u2019s I, and maximum-likelihood spatial-lag and spatial-error regression on a panel of ten NUTS-1 regions.'),
  p('The harmonised analysis panel comprises 58,420 region-by-sex-by-day observations spanning 4.33 million deaths. Temperature exposure is measured by the Hadley Centre Central England Temperature v2.1.1.0 daily-mean series (Met Office, 2025), replacing a jittered temperature column in the source mortality dataset that was identified as the legacy unadjusted CET series biased by exactly the documented post-1974 urban-warming adjustment. Green space is operationalised through two open-data substitutes for the proprietary Fields in Trust Green Space Index: square metres of public green space per capita, computed from OS Open Greenspace polygon areas under Fields-in-Trust typology exclusions and divided by Office for National Statistics mid-year population estimates; and average distance to the nearest park, from the ONS April 2020 reference tables. Building density is operationalised through the proportion of regional land area classified as built-up (OS Open Built Up Areas, 2024), with dwelling density (Census 2021 TS041), and population density as solid measures.'),
  p('The temporal frame supports a heat\u2013mortality effect. The cumulative relative risk over a 0\u201314 day lag at the 99th percentile of daily-mean Central England Temperature relative to the minimum-mortality temperature (18.3 \u00b0C) is 1.045, with 95% confidence interval [1.031, 1.058]. The cold burden is substantially larger, with a corresponding cumulative relative risk of 1.385 [1.344, 1.428], consistent with published estimates for temperate climates (Gasparrini et al., 2015). The spatial frame supports the existence of geographic clustering of excess heat-mortality in heatwave years: annual global Moran\u2019s I is positive and significant in 2017 (I = +0.36, p = 0.013) and 2018 (I = +0.30, p = 0.023). A High\u2013High cluster centred on the West Midlands is identified by Local Moran\u2019s I (p = 0.040), with East Midlands as a Getis-Ord G* hot spot at the 95% level.'),
  p('However, the two substantive hypotheses (H1 and H2) are not supported at NUTS-1 resolution. Bivariate Moran\u2019s I between green-space provision and excess mortality is +0.04 (p = 0.83), the opposite sign to that predicted by H1, and the bivariate I between building density and excess mortality is +0.03 (p = 0.90). Spatial-regression Lagrange Multiplier tests fail to reject the null of no residual spatial dependence (LM-lag p = 0.67; LM-error p = 0.95). The reversal of H1 and the null result on H2 are attributable to age-structure and urbanicity confounding operating at the regional scale, specifically, the simultaneous coincidence in London of high density, low per-capita green space, and the lowest crude mortality in the panel (a consequence of London\u2019s younger demographic structure). The paper recommends sub-regional (Lower Layer Super Output Area) replication with deprivation and age adjustment as the next research step.'),
  p([
    { text: 'Keywords: ', bold: true },
    { text: 'heatwave; mortality; distributed lag non-linear model; spatial autocorrelation; Central England Temperature; Local Moran\u2019s I; Fields in Trust Green Space Index; OS Open Greenspace; London paradox.' },
  ]),
  pageBreak(),
);

// ---- TOC ----
children.push(
  h1('Table of Contents'),
  new TableOfContents('Contents', {
    hyperlink: true,
    headingStyleRange: '1-3',
  }),
  pageBreak(),
);

// ---- List of Figures ----
children.push(
  h1('List of Figures'),
  p('Figure 1. Data pipeline: from external sources to harmonised analysis panel.'),
  p('Figure 2. Temperature audit: HadCET v2.1.1.0 monthly mean and the constant \u22120.21 \u00b0C bias in the supplied CSV.'),
  p('Figure 3. NUTS-1 reference map: the ten analysis regions.'),
  p('Figure 4. Spatial weights network: Queen contiguity edges between regional centroids.'),
  p('Figure 5. Daily mortality and Central England Temperature, 2014\u20132022.'),
  p('Figure 6. Heat-day exposure profile by day of year and analysis year.'),
  p('Figure 7. Monthly mortality heatmap by year.'),
  p('Figure 8. Annual mortality by region.'),
  p('Figure 9. Distributed-lag non-linear model: cumulative exposure\u2013response curve.'),
  p('Figure 10. Distributed-lag non-linear model: lag\u2013response surface.'),
  p('Figure 11. Sensitivity of the heat-day mortality ratio to specification.'),
  p('Figure 12. Choropleth of pooled excess heat-mortality per 100,000.'),
  p('Figure 13. Local Moran\u2019s I (LISA) cluster map of pooled excess heat-mortality.'),
  p('Figure 14. Getis-Ord G* statistic for pooled excess heat-mortality.'),
  p('Figure 15. Annual global Moran\u2019s I, 2014\u20132022.'),
  p('Figure 16. Bivariate scatter: H1 and H2 against pooled excess heat-mortality.'),
  pageBreak(),
);

// ---- List of Tables ----
children.push(
  h1('List of Tables'),
  p('Table 3.1. Annual global Moran\u2019s I on excess heat-mortality (Queen weights, 999 conditional permutations).'),
  p('Table 3.2. Bivariate Moran\u2019s I: spatial tests of H1 and H2 (Queen weights, 999 permutations).'),
  p('Table 3.3. Sensitivity of the heat-day mortality ratio to threshold and stratification.'),
  p('Table 3.4. Regional covariates and pooled excess heat-mortality.'),
  pageBreak(),
);

// ---- List of Acronyms ----
const ACRONYMS = [
  ['AIC', 'Akaike Information Criterion'],
  ['BGC', 'Boundary Generalised Clipped (ONS boundary release format)'],
  ['CEDA', 'Centre for Environmental Data Analysis (NERC)'],
  ['CET', 'Central England Temperature'],
  ['CI', 'Confidence Interval'],
  ['COVID-19', 'Coronavirus Disease 2019'],
  ['CSV', 'Comma-Separated Values'],
  ['DESNZ', 'Department for Energy Security and Net Zero (UK)'],
  ['DHSC', 'Department of Health and Social Care (UK)'],
  ['DLNM', 'Distributed Lag Non-linear Model'],
  ['DOI', 'Digital Object Identifier'],
  ['EPSG', 'European Petroleum Survey Group (coordinate reference codes)'],
  ['FiT', 'Fields in Trust'],
  ['GLM', 'Generalised Linear Model'],
  ['GSI', 'Green Space Index'],
  ['H1, H2', 'Hypothesis 1 (green space), Hypothesis 2 (building density)'],
  ['HadCET', 'Hadley Centre Central England Temperature'],
  ['HH, LL, HL, LH', 'LISA quadrants: High\u2013High, Low\u2013Low, High\u2013Low, Low\u2013High'],
  ['HSE', 'Higher School of Economics (National Research University, Moscow)'],
  ['IMD', 'Index of Multiple Deprivation (English government measure of small-area deprivation)'],
  ['IPCC', 'Intergovernmental Panel on Climate Change'],
  ['IRLS', 'Iteratively Reweighted Least Squares'],
  ['ITL', 'International Territorial Levels (UK successor to NUTS classification)'],
  ['JSON', 'JavaScript Object Notation'],
  ['k-NN', 'k-Nearest Neighbours (spatial weights specification)'],
  ['LAD', 'Local Authority District (UK administrative geography)'],
  ['LISA', 'Local Indicators of Spatial Association'],
  ['LM', 'Lagrange Multiplier (spatial-regression specification test)'],
  ['LSOA', 'Lower Layer Super Output Area (UK small-area census geography)'],
  ['LST', 'Land Surface Temperature'],
  ['ML', 'Maximum Likelihood'],
  ['MMT', 'Minimum-Mortality Temperature'],
  ['MODIS', 'Moderate Resolution Imaging Spectroradiometer (NASA satellite instrument)'],
  ['NUTS', 'Nomenclature of Territorial Units for Statistics (European regional classification)'],
  ['OLS', 'Ordinary Least Squares'],
  ['ONS', 'Office for National Statistics (UK)'],
  ['OS', 'Ordnance Survey (UK national mapping agency)'],
  ['PHE', 'Public Health England (predecessor to UKHSA)'],
  ['RR', 'Relative Risk'],
  ['SAR', 'Spatial Autoregressive (model); also Spatial Lag Model'],
  ['SE', 'Standard Error'],
  ['SEM', 'Spatial Error Model'],
  ['SSP', 'Shared Socioeconomic Pathway (IPCC emissions trajectory)'],
  ['TS', 'Topic Summary (ONS Census 2021 table series, e.g.\u00a0TS041)'],
  ['UHI', 'Urban Heat Island'],
  ['UKCP18', 'UK Climate Projections 2018'],
  ['UKHSA', 'UK Health Security Agency'],
  ['WGS84', 'World Geodetic System 1984 (geographic coordinate reference system)'],
  ['WKB', 'Well-Known Binary (OGC geometry encoding)'],
];

function acronymTable(rows) {
  const border = { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' };
  const borders = { top: border, bottom: border, left: border, right: border };
  const wAc = 1800; const wEx = 7560;
  const tableRows = rows.map(([ac, ex]) => new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: wAc, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: ac, bold: true, font: FONT, size: 20 })],
        })],
      }),
      new TableCell({
        borders,
        width: { size: wEx, type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: ex, font: FONT, size: 20 })],
        })],
      }),
    ],
  }));
  return new Table({
    width: { size: wAc + wEx, type: WidthType.DXA },
    columnWidths: [wAc, wEx],
    rows: tableRows,
  });
}

children.push(
  h1('List of Acronyms'),
  p('The following acronyms and abbreviations are used throughout the thesis. Standard statistical abbreviations are listed alphabetically alongside organisation names, data-product identifiers, and geographic-classification codes. Context-specific clarifications are given in parentheses where the acronym would otherwise be ambiguous.', { run: { italics: true } }),
  acronymTable(ACRONYMS),
  pageBreak(),
);

// Boundary: front matter ends here, body (arabic numerals) begins.
const BODY_START = children.length;

// =================================================================
// CHAPTER 1 — Introduction (expanded)
// =================================================================
children.push(
  h1('Chapter 1. Introduction'),

  h2('1.1 Background and motivation'),
  p('Climate change is reshaping the distribution of temperature extremes across northwestern Europe, and the United Kingdom, a country whose mortality climatology has historically been dominated by cold-related deaths, is increasingly experiencing heat episodes of a magnitude previously unfamiliar to its population, its built environment, and its public-health infrastructure. The 19 July 2022 Met Office record of a daily-maximum temperature of 40.3 \u00b0C at Coningsby, Lincolnshire, exceeded by 1.6 \u00b0C the previous national record set in 2019 and represented the first observed occurrence of 40 \u00b0C in the British Isles in the instrumental record (Met Office, 2022). The corresponding daily-mean Central England Temperature of 28.1 \u00b0C on the same day is, in this thesis\u2019s 2014\u20132022 analysis panel, the single highest daily-mean temperature observed. The 2022 episode produced approximately 2,803 excess deaths registered in the week-of-occurrence and the following two weeks combined, with the largest excesses concentrated in the over-65 age bands and the South East and London regions (Office for National Statistics, 2023).'),
  p('The 2022 event is not isolated. The 2003 European heatwave produced an estimated 2,091 excess deaths in England and Wales over the August 2\u201318 period alone (Johnson et al., 2005); the August 2018 heatwave generated approximately 863 excess deaths in the over-65 age bands; and the consecutive 2019 and 2020 episodes each produced excess mortality of similar order. The Met Office\u2019s probabilistic climate projections under the UKCP18 framework indicate that heatwave frequency will more than double, and intensity will rise by approximately 2 \u00b0C in regional-mean summer maximum, under all credible emissions trajectories through mid-century (Lowe et al., 2018). The implication is that the heat-mortality burden of the past decade, which has hovered between 0.5 and 2 percent of summer all-cause mortality, will grow non-linearly under realistic warming.'),
  p('At the same time, two structural features of the urban environment are routinely advanced as potential moderators of heat-mortality risk. The first is the cooling effect of public green space: urban parks, gardens, and similar vegetated areas reduce ambient air temperature through evapotranspirative cooling, provide direct radiative shading, and decrease the diurnal storage of heat in adjacent impervious surfaces (Bowler et al., 2010; Aram et al., 2019). The second is the urban heat-island effect itself: dense built-up areas retain heat through reduced sky-view factor, increased thermal mass, lower albedo, and reduced ventilation, generating intra-urban temperature differentials that can reach 5 \u201310 \u00b0C on calm summer nights in major British cities (Heaviside et al., 2016; Wilby, 2003). If green-space provision moderates exposure favourably, and if building density amplifies it adversely, then the distribution of heat-related mortality across English and Welsh regions should reflect these moderators, a testable proposition that this thesis evaluates.'),

  h2('1.2 Heatwave-related mortality in the United Kingdom'),
  p('Three features of the United Kingdom heat-mortality literature are worth establishing at the outset. First, the dose\u2013response between temperature and daily mortality is non-linear and broadly U-shaped, with a minimum-mortality temperature (MMT) typically estimated in the 17\u201320 \u00b0C range for British populations and rising relative risks above and below that point (Armstrong et al., 2011; Hajat et al., 2014). The asymmetry of the curve is pronounced: the cold-attributable fraction of all-cause mortality is approximately one order of magnitude larger than the heat-attributable fraction across the full calendar year (Gasparrini et al., 2015), but the heat fraction concentrates in a small number of days per year and constitutes a more concentrated and forecastable public-health hazard.'),
  p('Second, the lag structure of the heat effect is short. Distributed-lag non-linear models consistently identify cumulative effects of acute heat that peak at lag 0\u20133 days and are largely complete within seven days, with some evidence of mortality displacement (\"harvesting\") in the second week post-exposure (Gasparrini, 2014). This contrasts sharply with the cold side of the dose\u2013response, where effects accumulate gradually over 14\u201328 days and harvesting is less evident. The implication for analytical design is that a 0\u201314 day lag window is sufficient to capture the heat signal and the immediate aftermath, while leaving the cold-side lag structure partially censored, a limitation this thesis acknowledges in the Discussion.'),
  p('Third, vulnerability is socially and demographically structured. Heat-mortality risk is consistently elevated among older adults, particularly the over-75 age band; among women in the oldest age bands; among individuals living alone; among residents of dense urban housing without natural ventilation; and among populations of small areas with high deprivation scores (Hajat et al., 2007; Reid et al., 2009; Vandentorren et al., 2006; Masselot et al., 2023). The thesis\u2019s NUTS-1 / ITL-1 panel cannot directly observe any of these moderators, a constraint that the Discussion (Section 3.8) will treat as central rather than incidental.'),

  h2('1.3 The urban heat island and the moderator hypothesis'),
  p('The urban heat-island effect (UHI) is the systematic temperature elevation of urbanised areas relative to their rural surroundings, arising from the combined effect of reduced evaporative cooling (less vegetation and water bodies), greater absorption and storage of solar radiation by built surfaces, anthropogenic waste heat from buildings and vehicles, and reduced wind speeds caused by urban morphology (Oke, 1982; Heaviside et al., 2016). In the United Kingdom, the UHI has been most extensively documented for Greater London, where night-time temperature differentials of 7 \u00b0C between central and rural surroundings have been recorded during the 2003 and 2018 heatwaves (Macintyre et al., 2018). Smaller magnitudes have been recorded in Birmingham, Manchester, Greater Glasgow, and other major conurbations.'),
  p('The substantive premise of the present thesis is that, if the UHI mechanism operates as the microclimatic literature documents, regional differences in the structural features that drive it, green-space scarcity and density of the built fabric, should be reflected in regional differences in heat-related mortality. The two hypotheses of this thesis are explicit operationalisations of this premise at the NUTS-1 / ITL-1 scale, on the available eight-year panel and the available open-data measures.'),

  h2('1.4 Research question and hypotheses'),
  p('The principal research question is: do public green-space provision and building density, measured at the NUTS-1 / ITL-1 regional scale, moderate heat-related mortality in England and Wales over the period 2014\u20132022? This question is operationalised through two hypotheses:'),
  bullet('H1. Regions with lower per-capita public green-space provision experience higher heat-related mortality.'),
  bullet('H2. Regions with higher building density experience higher heat-related mortality.'),
  p('Each hypothesis is tested in two analytical frames. In the temporal frame, the heat\u2013mortality dose\u2013response is estimated from national daily data, and the role of green space and density is assessed indirectly through the cross-regional ranking of excess deaths. In the spatial frame, the geographic clustering of excess mortality is quantified, and its covariance with green space and density is tested through bivariate Moran\u2019s I and maximum-likelihood spatial regression. Both frames are applied to the same harmonised panel.'),

  h2('1.5 Statement of problem'),
  p('Existing UK-focused analyses of heat-mortality either operate on the national time series without explicit spatial structure (e.g., Hajat et al., 2014; Armstrong et al., 2011), or operate on small-area cross-sections without the temporal DLNM machinery now standard in the international literature (Vicedo-Cabrera et al., 2021; Masselot et al., 2023). The present thesis attempts to bridge these by applying a distributed-lag specification on national daily mortality (the temporal frame) and a complete spatial-autocorrelation suite on regional excess mortality (the spatial frame), using the same harmonised data panel for both. A secondary but consequential methodological contribution is the substitution of a proprietary Green Space Index (Fields in Trust) with two open-data measures that can be reproduced end-to-end from publicly available sources, addressing the increasingly important reproducibility expectation of environmental-epidemiology research.'),
  p('A third problem motivates the present revision. The original (unrevised) thesis of this title (Tsikudo, 2024) reported a region-level ordinary least squares regression with standardised coefficients of +5.66 and \u22125.66 on the green-space and building-density variables respectively, with R\u00b2 = 0.017. The symmetry of the coefficients raised, in pre-revision audit, the possibility that the building-density variable had been mechanically constructed from the green-space variable (in which case the regression would not constitute a meaningful test of either hypothesis). The revision establishes, via independent open-data construction of both variables, that this is not the case: the two open-data measures correlate at r = \u22120.78, not at \u00b11.00. The symmetry of the original coefficients is therefore an artefact of two separate single-predictor standardised regressions on the same response, not of mechanical construction, and is reported in Section 3.6 as a transparency reference rather than as an analytical result.'),

  h2('1.6 Objectives'),
  p('The study has five objectives:'),
  bullet('To assemble a region-by-sex-by-day mortality panel for England and Wales covering 2014\u20132022, harmonised against a clean Central England Temperature series.'),
  bullet('To estimate the dose\u2013response between daily temperature and mortality using a distributed-lag non-linear model with cumulative effects over a 0\u201314 day lag.'),
  bullet('To quantify and locate spatial clustering of excess heat-mortality at the NUTS-1 / ITL-1 regional scale using global and local indicators of spatial association.'),
  bullet('To test the two substantive hypotheses (H1, H2) using bivariate Moran\u2019s I and maximum-likelihood spatial regression with Lagrange Multiplier specification testing.'),
  bullet('To recommend a concrete next research step to address the limitations of NUTS-1 resolution.'),

  h2('1.7 Methodological approach (summary)'),
  p('Mortality data are obtained from supplied region-by-sex-by-day CSV files for 2014\u20132019, 2021, and 2022 (eight analysis years; 2020 is excluded because of the surging covid-19 deaths at the time). The CSV files are audited (see Section 3.1) and harmonised: the Welsh placeholder code W99999999 is remapped to the canonical W92000004; phantom 29-February rows in non-leap years are dropped; and the temperature data is obtained from the Hadley Centre Central England Temperature daily-mean series version 2.1.1.0 (Met Office, 2025). Heatwave windows are recomputed from the clean temperature series using the Met Office operational definition of three or more consecutive days with daily mean temperature at or above a specified threshold (primary specification: 20 \u00b0C).'),
  p('External regional covariates are sourced from open data (see Section 3.2). The H1 predictor is the two open measures: square metres of public green space per capita (computed from OS Open Greenspace polygons, divided by ONS mid-year population estimates) and average distance to the nearest park (Office for National Statistics, 2020). Building density (H2) is measured by the proportion of regional land area classified as built-up (OS Open Built Up Areas, 2024), with dwelling density (Census 2021 Table TS041), and population density as solid measures.'),
  p('The full pipeline is shown schematically in Figure 1, which traces the flow from six external open-data sources to the harmonised analysis panel, and from the panel to the two analytical frames.'),
  ...fig('01_pipeline.png', { figNum: 'Figure 1', title: 'Data pipeline: from external sources to harmonised analysis panel.', note: 'Six open-data sources are harmonised into a 58,420-row panel (region \u00d7 sex \u00d7 day, 2014\u20132022, eight years \u00d7 ten regions \u00d7 two sexes), which feeds both temporal and spatial frames.' }),
  p('The temporal frame implements a quasi-Poisson generalised linear model with a natural-cubic-spline cross-basis for temperature and lag (Gasparrini, 2011), with seasonal and long-term trend control via a natural spline on continuous time at eight degrees of freedom per year, day-of-week fixed effects, and a population offset. The spatial frame applies global Moran\u2019s I (pooled and per year), Local Moran\u2019s I with conditional 999-permutation inference, Getis-Ord G*, bivariate Moran\u2019s I, and maximum-likelihood spatial-lag and spatial-error regression with Lagrange Multiplier specification diagnostics (Anselin, 1988, 1995).'),

  h2('1.8 Relevance and contribution'),
  p('The Met Office\u2019s 2022 dataset, which records the first observed 40 \u00b0C daily-maximum temperature in the United Kingdom, has heightened policy interest in regionally targeted heat-resilience adaptation. A defensible regional decomposition of heat-mortality risk, including both confirmation of where it exists and honest disclosure of where it cannot be identified at current resolution, directly informs the United Kingdom Health Security Agency Heatwave Plan and analogous regional public-health planning. The methodological contribution of the thesis is the demonstration that the proprietary Fields in Trust Green Space Index can be substituted by an open-data measure with identical regional ordering on the four-region public comparison (Pearson r = 1.00), supporting end-to-end reproducibility from publicly available sources; this matters for the broader field because environmental-epidemiology research is increasingly expected to be replicable from open data.'),
  p('The substantive contribution is the explicit identification (in Sections 3.7 and 3.8) of why H1 and H2 cannot be supported at NUTS-1 resolution, and in the concrete LSOA-level next-research-step specification (Section 4.3). The negative finding is not a failure: it is an informative result about the resolution limits of NUTS-1 regional analysis for moderator hypotheses dependent on within-region demographic structure, and it indicates the direction in which the field should move to obtain identification.'),
);

// =================================================================
// CHAPTER 2 — Literature Review (expanded)
// =================================================================
children.push(
  h1('Chapter 2. Literature Review'),

  h2('2.1 Quantifying heatwave mortality: the international evidence'),
  p('The empirical study of heat-related mortality has matured significantly since the watershed events of the 2003 European summer heatwave, which produced approximately 70,000 excess deaths across the affected countries (Robine et al., 2008). That episode prompted a methodological transformation that shifted the field from descriptive single-event analyses to formal statistical estimation of temperature\u2013mortality dose\u2013response curves using long observational records. The principal vehicle for this transformation has been the distributed-lag non-linear model (DLNM) framework developed by Gasparrini and colleagues over the period 2010\u20132014 (Gasparrini et al., 2010; Gasparrini, 2011, 2014). The framework permits the joint estimation of the non-linear temperature\u2013mortality relationship and the lag structure of the response, making it possible to characterise both the shape of the dose\u2013response curve and the temporal distribution of its effects, information that aggregate summary statistics cannot capture.'),
  p('Applied internationally, the DLNM framework has produced a body of comparable estimates across some 384 cities in 13 countries (Gasparrini et al., 2015). The principal substantive findings have proven remarkably consistent: the temperature\u2013mortality relationship is U-shaped in nearly all settings, with minimum-mortality temperatures (MMTs) typically located between the 70th and 90th percentiles of the local temperature distribution. The cold-attributable fraction of mortality substantially exceeds the heat-attributable fraction in temperate climates, by a factor of approximately ten in the United Kingdom, although in subtropical and tropical settings, the heat fraction approaches or exceeds the cold fraction. The relative-risk magnitudes at moderately hot temperatures (90th\u201395th percentile vs MMT) are typically in the range 1.05\u20131.20, with the upper extreme (99th percentile and above) producing relative risks of 1.15\u20131.40 in more vulnerable populations.'),
  p('The Vicedo-Cabrera et al. (2021) multi-country attribution analysis extended the framework to climate-change attribution and identified that approximately 37 percent of heat-related deaths in the recent decades can be attributed to anthropogenic warming. For the United Kingdom specifically, this figure corresponded to approximately 82 heat-attributable deaths per year over the 1991\u20132018 period that would not have occurred under counterfactual non-warming conditions. The headline implication is that the heat-mortality burden is not only present and quantifiable but is partially attributable to causes that ongoing emissions trajectories will continue to amplify.'),

  h2('2.2 UK-specific estimates and trends'),
  p('The methodological reference point for heat-mortality estimation in the United Kingdom is the work of Armstrong, Hajat and colleagues, who applied DLNM and related distributed-lag specifications to English and Welsh mortality data over the period 1993\u20132006 (Armstrong et al., 2011) and produced the first systematic UK estimate of an MMT (approximately 17\u201318 \u00b0C in regional daily-mean exposure) and a cumulative heat RR at the 99th percentile of approximately 1.04\u20131.07. These values have been broadly reproduced in subsequent analyses with extended time series, including Hajat et al. (2014) and Gasparrini et al. (2015), both of which place the UK heat RR at the 99th percentile in a narrow band around 1.04\u20131.06, a range in which the present thesis\u2019s estimate of 1.045 sits comfortably.'),
  p('More recent UK-specific work has focused on two refinements. The first is downward extension of resolution: Masselot et al. (2023) report heat-attributable mortality risks at Lower Layer Super Output Area (LSOA) level for England and Wales, with approximately 34,000 small areas in the analysis. Their finding of substantial gradients across deprivation deciles, with heat RRs in the most deprived decile some 30\u201350 percent larger than in the least deprived decile, establishes a benchmark against which a NUTS-1 analysis can be located in the resolution hierarchy. The second refinement is the integration of long-term temperature records: Bennett et al. (2014) and successor studies have used the multi-century Central England Temperature series to characterise climatic baselines and to anchor heat-attributable mortality estimates against pre-industrial conditions.'),
  p('Importantly, the temporal trend in the United Kingdom is not unidirectional. Comparing the 2003 European heatwave with subsequent comparable-temperature episodes, the case-fatality ratio (deaths per degree-day above threshold) appears to have declined modestly over the 2003\u20132018 period (de\u2019 Donato et al., 2015). This adaptation effect is attributed to a combination of factors: the formalisation of the United Kingdom Heatwave Plan in 2004 with operational alert levels and intervention protocols (Department of Health, 2004; subsequent versions through 2018); incremental improvement of the housing stock through Building Regulations Part L (Conservation of Fuel and Power) updates; and behavioural awareness from public-health communications. The adaptation effect is empirically real but is dwarfed in magnitude by the underlying climatic trend.'),

  h2('2.3 The DLNM framework and its applications'),
  p('The distributed-lag non-linear model addresses two technical features of the temperature\u2013mortality relationship that simpler models cannot. The first is the non-linearity of the relationship at any single lag: the relative risk does not vary monotonically with temperature but exhibits a minimum near the MMT and rises in both directions. The second is the lag-distributed nature of the relationship: the relative risk for a given temperature deviation depends not only on the temperature on the day of exposure but also on the temperatures of the preceding days. The DLNM captures both features simultaneously by specifying a \"cross-basis\" matrix that is the tensor product of a basis on temperature and a basis on lag, and by using this cross-basis as a covariate set in a generalised linear or generalised additive model framework.'),
  p('The choice of basis functions is consequential. Natural cubic splines are the canonical choice in the DLNM literature, with four to six degrees of freedom on the temperature axis (sufficient to capture the U-shape without overfitting) and three to five degrees of freedom on the lag axis (sufficient to capture the rise-and-decay shape of the lag distribution). The cross-basis approach has the substantial advantage that the marginal exposure\u2013response curve at any single lag, the marginal lag\u2013response curve at any single temperature, and the cumulative exposure\u2013response curve summed over the full lag window can all be derived from a single fitted model. The present thesis adopts this approach, with four degrees of freedom on temperature and four on lag, over a 0\u201314 day lag window.'),
  p('The control for seasonal and long-term confounding in DLNM specifications is critical. The standard convention, established by Gasparrini and colleagues, is a natural spline on continuous time at seven to eight degrees of freedom per year, plus day-of-week fixed effects. This level of seasonal control is sufficient to remove the strong winter\u2013summer mortality cycle and the long-term mortality trend, leaving the residual variation in mortality for the temperature cross-basis to explain. Inadequate seasonal control, for example, three or four degrees of freedom per year, leaves residual seasonal confounding in the temperature term and biases the heat coefficient downward, sometimes to the point of producing relative risks below unity at warm temperatures. The present thesis encountered this exact phenomenon during model development (see Section 3.4.2) and resolved it by adopting the canonical eight-degrees-of-freedom-per-year specification.'),

  h2('2.4 Spatial analysis of mortality and the urban heat-island effect'),
  p('The application of formal spatial-autocorrelation statistics to heat-mortality has lagged behind the temporal frame, in part because the relevant outcome (excess heat-mortality) is sparse in time and noisy at fine spatial resolution. Nonetheless, two methodological reference points have established themselves. The first is the global Moran\u2019s I statistic (Moran, 1948), which quantifies the average correlation between an observation\u2019s value and the values of its spatial neighbours, with significance assessed by permutation under a null of spatial independence. The second is the local indicators of spatial association (LISA) framework developed by Anselin (1995), which decomposes the global Moran\u2019s I into observation-specific contributions and identifies local clusters of high-high, low-low, high-low, and low-high covariation. For hot-spot identification specifically, the Getis-Ord Gi* statistic (Getis & Ord, 1992) provides a complementary indicator that differs from LISA in including the focal observation in the local sum and in testing the null of no local concentration rather than the null of local independence.'),
  p('The application of these statistics to UK heat-mortality outcomes is rare. Heaviside, Macintyre and Vardoulakis (2016) provide the principal English review of the urban heat-island literature, emphasising that the absolute UHI magnitude is largest for Greater London but that secondary cities (Birmingham, Manchester, Sheffield, Leeds, Liverpool) also exhibit measurable warming relative to their rural surroundings. The mortality consequences of the UHI have been quantified in city-specific studies (e.g., Macintyre et al., 2018, for the West Midlands), but a UK-wide systematic spatial-autocorrelation analysis of heat-mortality at the NUTS-1 regional scale has not, to my knowledge, been previously reported.'),
  p('The methodological approach in the present thesis follows Anselin (1988) and the empirical practice systematised in Anselin and Rey (2014). The principal weights specification is Queen contiguity, in which any two regions sharing at least a common point are treated as neighbours; this is supplemented by Rook contiguity (sharing a common edge) and a k-nearest-neighbours specification with k = 3, with the latter serving as a solid check. The spatial-regression specification is the standard Anselin\u2013Bera\u2013Florax framework, in which Lagrange Multiplier tests on the residuals of an ordinary least squares regression are used to discriminate between the spatial-lag (SAR) and spatial-error (SEM) maximum-likelihood alternatives. With N = 10 NUTS-1 regions, the framework is applied with explicit awareness that statistical power is limited and that pseudo-p-values from permutation inference have coarse resolution.'),

  h2('2.5 Socioeconomic factors, age, and vulnerability'),
  p('Vulnerability to heat-related mortality is consistently elevated in specific subpopulations across virtually all settings studied. The principal moderators are: chronological age, with risk rising sharply above 65 and again above 85 years (Vandentorren et al., 2006; Reid et al., 2009); sex, with women in the oldest age bands at elevated risk relative to age-matched men, although this pattern is partially confounded with the higher female longevity; living arrangements, with individuals living alone at substantially elevated risk relative to those in multi-occupancy households; pre-existing chronic disease, particularly cardiovascular and respiratory conditions, dementia, and chronic kidney disease; medication regimens that interfere with thermoregulation, including diuretics, anticholinergics, and psychotropics; and small-area socioeconomic deprivation, indexed in the United Kingdom by the Index of Multiple Deprivation 2019 deciles (Public Health England, 2019).'),
  p('Most of these moderators are unobservable at the NUTS-1 / ITL-1 scale and constitute the principal confounders for the regional-scale tests in this thesis. The thesis observes no age stratification (the acquired mortality data are aggregated to region by sex by day with no age structure); no individual-level living arrangements or chronic-disease prevalence; no medication exposure; and no within-region IMD variation. Regional variation in the age structure of the population, with London substantially younger in median age than the rural English regions and Wales (Office for National Statistics, 2023), generates a regional confound between density, green-space, and demographic vulnerability that the thesis is, on its own data, unable to disentangle. Section 3.8 develops this point as the central interpretive issue of the negative result on H1 and H2.'),

  h2('2.6 Temporal adaptation and future risk under climate change'),
  p('Two opposing forces shape the future heat-mortality burden. The first is adaptation, which empirically reduces the case-fatality ratio at any given temperature over time through housing-stock improvement, behavioural change, healthcare-system response, and the rolling diffusion of mechanical cooling (air conditioning). For the United Kingdom, this effect is small in absolute magnitude (Hajat et al., 2014 estimate that adaptation will offset approximately 15\u201320 percent of the climate-attributable rise in heat-mortality through 2050) and the saturation point of adaptation is poorly characterised. The second force is climatic intensification: the IPCC Sixth Assessment Report (IPCC, 2021) considers further increases in heat extremes a high-confidence outcome under all evaluated emissions trajectories, with the United Kingdom\u2019s summer maximum temperature distribution shifting upward by approximately 2 \u00b0C by 2050 under SSP2-4.5 and 4 \u00b0C by 2100 under SSP5-8.5.'),
  p('The implication is that the absolute heat-attributable mortality burden in the United Kingdom is likely to rise substantially over the coming decades, even as the per-degree case-fatality ratio falls modestly through adaptation. Recent projections (e.g., Hajat et al., 2014; Vardoulakis et al., 2014) place the heat-attributable mortality burden in the 2050s at approximately 5,000\u20137,000 excess deaths per year for England and Wales under moderate emissions, rising to potentially 12,000\u201315,000 per year under high-emissions trajectories, several times the present-day burden of approximately 1,000\u20132,000 per year. The Romanello et al. (2023) Lancet Countdown report places the European 2022 summer heat-attributable mortality at approximately 61,000 deaths, of which the United Kingdom share was approximately 3,300, indicating that the magnitude of the burden is substantial and growing.'),

  h2('2.7 Research gaps and the present contribution'),
  p('The principal gap addressed by this thesis is the absence of a single, reproducible UK-focused study that simultaneously applies the DLNM framework to recent post-2014 mortality data, executes the full spatial-autocorrelation suite on the same panel, and tests the green-space and density hypotheses using fully open-data measures. The methodological contribution lies in the open-data and the application of the spatial-statistics suite at the NUTS-1 scale (Section 3.5). The substantive contribution lies in the explicit identification (in Section 3.8) of why H1 and H2 cannot be supported at NUTS-1 resolution, and in the concrete LSOA-level next-research-step specification (Section 4.3).'),
  p('A secondary contribution worth noting is the analytical pipeline\u2019s transparency. Every estimator used in the thesis was implemented in pure numpy/scipy. The implementations are validated against published formulae and against the known reference cases (clustered patterns producing positive Moran\u2019s I, checkerboard patterns producing negative Moran\u2019s I, simulated Poisson data with known coefficients), and the canonical libraries are documented and can be provided upon request; for reproduction on a standard machine. This explicit transparency of statistical implementation is, in my view, valuable for didactic purposes and for the methodological literacy of subsequent researchers.'),
);

// =================================================================
// CHAPTER 3 — Analysis and Results (substantially expanded)
// =================================================================
children.push(
  h1('Chapter 3. Analysis and Results'),

  h2('3.1 Data audit and harmonisation'),
  p('The supplied source dataset comprises eight comma-separated value files, one per analysis year (2014\u20132019, 2021, 2022), each containing 7,300 or 7,320 region-by-sex-by-day records. A pre-analysis audit identified several issues, of which four were material and were resolved as follows.'),
  p([
    { text: 'First, the temperature column.', bold: true },
    { text: ' The daily temperature column in the supplied CSVs exhibited a stable mechanical jitter pattern across all years, specifically, increments of +0.10 \u00b0C and +0.01 \u00b0C consistent with a Microsoft Excel auto-fill operation, rather than with a real meteorological series. Comparing the daily mean of the 20 jittered values (one per region-sex cell) per day against the Hadley Centre Central England Temperature daily-mean version 2.1.1.0 (Met Office, 2025) revealed a stable \u22120.21 \u00b0C bias every year, equal to the documented post-1974 urban-warming-bias adjustment applied in the v2.x series (Parker et al., 1992; Legg et al., 2025). The CSV temperature column was therefore replaced wholesale with HadCET v2.1.1.0 daily mean, broadcast across all 20 region-by-sex rows of each calendar day. Figure 2 visualises the audit result.' },
  ]),
  ...fig('02_temp_audit.png', { figNum: 'Figure 2', title: 'Temperature audit.', note: '(a) HadCET v2.1.1.0 monthly-mean series for the 2014\u20132022 analysis period, showing the expected annual cycle and the 2018 and 2022 summer peaks. The 2020 gap reflects the exclusion of 2020/covid-19 from the mortality panel. (b) The bias between the supplied CSV temperature column and HadCET v2.1.1.0 is a constant \u22120.21 \u00b0C in every year, exactly matching the documented post-1974 urban-warming-bias adjustment applied to HadCET v2.x.' }),
  p([
    { text: 'Second, the Welsh region code.', bold: true },
    { text: ' Wales carried the placeholder code W99999999, which is the Office for National Statistics placeholder for the value \u201cnot applicable\u201d. This was remapped to the canonical Wales country code W92000004 throughout, consistent with the December 2022 geography release.' },
  ]),
  p([
    { text: 'Third, phantom 29 February rows.', bold: true },
    { text: ' Each non-leap year contained 20 phantom 29-February records (one per region-sex cell) with null dates and zero death counts; these were dropped from the analysis panel.' },
  ]),
  p([
    { text: 'Fourth, the heatwave-period column.', bold: true },
    { text: ' The supplied Heat.period column did not reproduce the Met Office operational heatwave definition consistently across years; heatwave windows were therefore recomputed from the clean HadCET series as periods of three or more consecutive days at or above a specified threshold, with the primary threshold set at 20 \u00b0C daily mean.' },
  ]),
  p('The resulting harmonised panel contains 58,420 region-by-sex-by-day records over eight years, ten regions and two sexes, totalling 4,329,012 deaths. England plus Wales mid-2022 population, aggregated from the ONS Mid-Year Estimates (MYEB1) by Local Authority District and remapped to NUTS-1 / ITL-1 regions, totals 60.28 million, agreeing with the published ONS estimate. Figure 3 shows the ten analysis regions with their canonical labels.'),
  ...fig('03_region_map.png', { figNum: 'Figure 3', title: 'NUTS-1 reference map.', note: 'The ten analysis regions: nine English Regions (December 2022 BGC release) plus Wales (Countries, December 2022 BGC release). Coloured for visual distinction only.' }),

  h2('3.2 External covariates and the open data'),
  p('The original thesis reported the Fields in Trust Green Space Index for four regions (Fields in Trust, 2020) as the H1 predictor. The Fields in Trust data are not available as open data (a constraint stated explicitly on the Fields in Trust website), and the present revision therefore replaces the proprietary index with two open measures. The primary H1 measure is the per-capita area of public green space, computed from OS Open Greenspace (Ordnance Survey, 2024) polygon areas under the same typology exclusions used by Fields in Trust (allotments, religious grounds, cemeteries, and golf courses excluded), divided by mid-2020 regional population. The four originally-reported Fields in Trust values rank the four regions in exactly the same order as the computed per-capita measure (Pearson r = 1.00 on the four-point comparison), providing direct validation that the open substitute preserves the substantive ordering of regions.'),
  p('The secondary H1 measure is the average distance from a population-weighted point to the nearest park, taken directly from the Office for National Statistics Region Parks reference table (Office for National Statistics, 2020). This is a measure of accessibility rather than per-capita provision, and the correlation between the two measures across the ten regions is +0.80, indicating that they capture genuinely various aspects of green-space supply. London ranks lowest on per-capita provision (22.8 m\u00b2 per person, FiT-comparable typology) but ranks first on accessibility (387 m to the nearest park), a pattern that bears directly on the London paradox discussed in Section 3.8.'),
  p('For H2, the primary measure is the proportion of regional land area classified as built-up, computed from OS Open Built Up Areas (Ordnance Survey, 2024) via centroid-in-polygon spatial join against the NUTS-1 boundaries. Dwelling density (households per km\u00b2 from Census 2021 Table TS041 divided by region land area) and population density (mid-2020 population divided by region land area) are reported as solid measures. The three H2 measures correlate above 0.99 across the ten regions, indicating that they capture the same underlying urbanisation gradient.'),
  ...table(
    ['Region', 'GS m\u00b2/cap', 'Dist-to-park (m)', '% built-up', 'Dwellings/km\u00b2', 'Excess/100k'],
    [
      ['North East', '34.6', '827', '8.6%', '147', '14.7'],
      ['North West', '32.8', '700', '13.5%', '534', '18.4'],
      ['Yorkshire & Humber', '30.4', '720', '10.1%', '361', '19.0'],
      ['East Midlands', '38.7', '1,037', '9.4%', '316', '28.7'],
      ['West Midlands', '36.4', '780', '12.2%', '465', '26.2'],
      ['East of England', '42.3', '1,349', '10.2%', '335', '20.5'],
      ['London', '22.8', '387', '74.6%', '2,184', '15.6'],
      ['South East', '37.9', '991', '14.6%', '494', '23.9'],
      ['South West', '33.7', '1,346', '7.1%', '243', '20.8'],
      ['Wales', '41.6', '1,180', '4.5%', '147', '21.4'],
    ],
    { tableNum: 'Table 3.4', title: 'Regional covariates and pooled excess heat-mortality.', note: 'GS m\u00b2/cap = green space per capita under FiT typology exclusions (primary H1). Dist-to-park = ONS Region Parks April 2020 reference. % built-up = OS Open Built Up Areas / land area. Dwellings/km\u00b2 = Census 2021 TS041. Excess/100k = pooled excess deaths per 100,000 population on heat days over 2014\u20132022 (Met Office definition, 20 \u00b0C, 3 consecutive days).' },
  ),

  h2('3.3 Spatial layers and weights'),
  p('Region polygons are taken from the ONS Open Geography Portal: nine English NUTS-1 regions (Regions, December 2022, BGC, generalised, clipped) and the Wales country boundary (Countries, December 2022, UK BGC). Region land areas computed from these boundary polygons by spherical-area integration match published ONS standard areas within 0.3% (negligible bias from coastline generalisation). Three spatial-weights matrices are constructed: a Queen contiguity matrix and a Rook contiguity matrix (both yielding 18 edges and an average of 3.60 neighbours, with no isolated regions), and a k-nearest-neighbours matrix at k = 3. All matrices are row-standardised. The Queen matrix is the primary weights specification; Rook and k-NN are reported as robustness. Figure 4 visualises the Queen contiguity network on the spatial layer.'),
  ...fig('04_weights_network.png', { figNum: 'Figure 4', title: 'Spatial weights network.', note: 'Queen contiguity edges between regional centroids, with row-standardised weights matrix W (10\u00d710) derived from the December 2022 NUTS-1 boundaries. Every region has between two and five neighbours; no region is isolated. Wales\u2013South East share a boundary point via the Bristol Channel under the Queen but not the Rook specification; this is the only edge that differs between the two contiguity definitions, and the resulting Moran\u2019s I values are identical for both.' }),
);

// Chapter 3 continued — temporal models
children.push(
  h2('3.4 Temporal models'),
  p('The temporal-frame analysis applies five complementary specifications to the harmonised panel. The principal estimator is the distributed-lag non-linear model (Section 3.4.2). The other four, a baseline ordinary least squares (Section 3.6), a quasi-Poisson generalised linear model with natural splines (Section 3.4.1), a rolling-baseline excess-mortality calculation (Section 3.4.3), and a sensitivity analysis on heat-day mortality ratios (Section 3.4.4), provide methodological context and resilient. Figure 5 displays the underlying daily national mortality and temperature series that all temporal specifications operate on.'),
  ...fig('05_daily_series.png', { figNum: 'Figure 5', title: 'Daily mortality and Central England Temperature, 2014\u20132022.', note: 'Daily national death totals (grey, left axis) overlaid on the HadCET v2.1.1.0 daily mean (red, right axis). The dominant cycle is annual, with winter peaks of approximately 2,500 deaths/day and summer troughs of approximately 1,500 deaths/day. Summer 2018 and 2022 heatwave peaks are annotated.' }),
  p('A subsidiary question concerns the timing of heatwave exposure. Figure 6 plots the heat-day exposure profile by day of year and analysis year under the primary 20 \u00b0C / 3-day Met Office definition. The 2018 cluster runs from late June through early August; the 2019 cluster is concentrated in late July; the 2022 cluster is similarly late July; and the 2021 cluster is a brief late-July episode. There are no heat days satisfying the 20 \u00b0C / 3-day rule in 2014, 2015, 2016, or in early summer of any year.'),
  ...fig('06_heat_days.png', { figNum: 'Figure 6', title: 'Heat-day exposure profile.', note: 'Heat days by day of year and analysis year, coloured by daily-mean Central England Temperature. The 2018 and 2022 clusters are the principal heat-attributable mortality exposures in the panel.' }),

  h3('3.4.1 Quasi-Poisson generalised linear model with natural splines'),
  p('A quasi-Poisson generalised linear model is fitted to daily deaths with log link, population offset, and the following terms: a four-degree-of-freedom natural cubic spline on daily-mean Central England Temperature, an eight-degree-of-freedom natural cubic spline on day-of-year, year fixed effects, region fixed effects, and a female indicator. The estimated dispersion parameter is 1.52, confirming over-dispersion relative to a strict Poisson model and justifying the quasi-Poisson specification. The relative risk at the 99th versus the 50th percentile of daily-mean temperature, controlling for the other model terms, is 1.122 with 95 percent confidence interval [1.113, 1.131]. The female-versus-male coefficient on the log scale is \u22120.020 (standard error 0.0012), implying a relative risk of 0.980 for females, consistent with the established small male excess in daily mortality rates.'),

  h3('3.4.2 Distributed-lag non-linear model'),
  p('The DLNM specification follows Gasparrini (2011): a cross-basis matrix is constructed as the tensor product of a four-degree-of-freedom natural cubic spline on temperature with a four-degree-of-freedom natural cubic spline on lag, evaluated over a 0\u201314 day lag window on national daily total deaths (regions summed). Seasonal and long-term trend control is provided by a natural spline on continuous time at eight degrees of freedom per year (sixty-four total), together with day-of-week fixed effects. Standard quasi-Poisson estimation with Pearson dispersion is used.'),
  p('The minimum-mortality temperature (MMT) is estimated at 18.3 \u00b0C, in the upper third of the temperature distribution as expected for a temperate climate. The cumulative relative risk over the full 0\u201314 day lag window at the 99th percentile of daily-mean temperature (21.2 \u00b0C) relative to the MMT is 1.045, with 95 percent confidence interval [1.031, 1.058]. The cold side of the curve is substantially steeper: the cumulative relative risk at the 1st percentile (0.0 \u00b0C) relative to the MMT is 1.385 [1.344, 1.428]. This pattern, a U-shaped cumulative dose\u2013response with a cold burden several times larger than the heat burden, is consistent with the published UK literature (Gasparrini et al., 2015; Vicedo-Cabrera et al., 2021). Figure 9 plots the full cumulative exposure-response curve.'),
  ...fig('09_dlnm_curve.png', { figNum: 'Figure 9', title: 'Distributed-lag non-linear model: cumulative exposure\u2013response curve.', note: 'Cumulative relative risk over 0\u201314 day lag, relative to the minimum-mortality temperature (MMT) of 18.3 \u00b0C. Cold side: RR 1.385 at 0 \u00b0C [95% CI 1.344, 1.428]. Heat side: RR 1.045 at the 99th percentile (21.2 \u00b0C) [95% CI 1.031, 1.058]. The asymmetry of the curve is the headline temporal-frame result: the cold burden in temperate England is several times larger than the heat burden.' }),
  p('A methodological note is warranted. An earlier version of the cumulative-RR contrast in this specification compared the 99th percentile of daily-mean temperature against the 50th percentile (10.5 \u00b0C). Because the 50th percentile lies on the cold side of the MMT (18.3 \u00b0C), the resulting contrast captured the difference between two points that both carry elevated risk relative to the MMT, a relatively warm point and a relatively cold point, and produced a cumulative RR below 1. The 99th-versus-MMT contrast is the correct quantity for a hypothesis about the effect of heat (as opposed to the joint effect of temperature in either tail) and is the contrast reported above. The mistake is preserved in this note for didactic purposes; the corrected estimate is the one reported throughout the thesis.'),
  p('The lag-response surface, obtained by evaluating the cross-basis at each combination of temperature and lag, is shown in Figure 10. The principal features are: (i) the rapid onset of the heat effect at lag 0\u20133, with peak relative risk near 1.05 at the hottest temperatures; (ii) the gradual decay of the lag-distributed heat effect through lag 7\u201310; and (iii) the much larger and more persistent cold effect at low temperatures and long lags.'),
  ...fig('10_lag_surface.png', { figNum: 'Figure 10', title: 'Distributed-lag non-linear model: lag\u2013response surface.', note: 'Relative risk surface over the full temperature \u00d7 lag grid, with the MMT reference line at 18.3 \u00b0C. The cold side (below MMT) shows the larger and more persistent effect; the heat side (above MMT) shows a smaller but more concentrated lag distribution. The displayed surface uses a kernel approximation distributing the cumulative RR across lags; a per-lag GLM re-fit produces a qualitatively identical pattern.' }),

  h3('3.4.3 Excess mortality with a rolling baseline'),
  p('Excess mortality on heat days is computed by subtracting from observed deaths an expected baseline equal to the mean of deaths in the corresponding region-by-sex-by-week-of-year cell across all years in the panel (a five-year-equivalent rolling baseline). At pooled level, the East Midlands records the highest excess deaths per 100,000 population on heat days (28.7), followed by the West Midlands (26.2) and the South East (23.9). London records 15.6, materially below several less-dense regions. This pattern, with London below several rural regions on the per-capita measure, is the proximate evidence for the age-structure confounding discussed in Section 3.8. Figure 7 shows the regional mortality pattern by month and year, identifying the months of highest excess.'),
  ...fig('07_month_year.png', { figNum: 'Figure 7', title: 'Monthly mortality heatmap.', note: 'National monthly death totals (regions summed) by analysis year. January 2015, 2017, 2018, and the December 2022 cell each exhibit elevated values; July/August 2018 and 2022, the principal heatwave months, are visible as warm cells against an otherwise cool summer pattern.' }),
  p('Annual mortality by region (Figure 8) shows the strong cross-regional variation that is dominated by population size: the North West and South East have substantially larger annual death counts than the smaller Wales and the North East, in proportion to their populations. Per-capita rates (not shown) compress these differences but reveal the regional ranking that drives the spatial-frame analysis.'),
  ...fig('08_regional_annual.png', { figNum: 'Figure 8', title: 'Annual mortality by region.', note: 'Total deaths per year by region, 2014\u20132022. Mean annual values are noted on each panel. The pattern is dominated by population size; per-capita rates display the regional variation more clearly and are reported in Table 3.4.' }),

  h3('3.4.4 Sensitivity analyses'),
  p('Three sensitivity analyses are reported. First, the heat-day mortality ratio (mortality on heat days divided by mortality on non-heat summer days) is computed at heatwave thresholds of 18 \u00b0C, 20 \u00b0C and 22 \u00b0C: 1.069, 1.106 and 1.152 respectively, increasing monotonically with the threshold as expected. The 25 \u00b0C threshold yields zero qualifying days under the three-consecutive-days rule (daily mean Central England Temperature does not reach 25 \u00b0C in this period) and is reported as degenerate. Second, excluding 2021 reduces the heat-day mortality ratio only marginally, from 1.106 to 1.099, indicating that the 2021 episodes are not the principal driver of the headline ratio. Third, sex-stratified analysis returns a ratio of 1.111 for males and 1.100 for females, consistent with the small female-protective effect observed in Section 3.4.1. Figure 11 visualises the sensitivity sweep.'),
  ...fig('11_sensitivity.png', { figNum: 'Figure 11', title: 'Sensitivity of the heat-day mortality ratio to specification.', note: 'Seven specifications: heatwave thresholds 18/20/22 \u00b0C, exclusion of 2021, sex stratification, and the all-data reference. The reference horizontal line at ratio = 1 corresponds to no heat-day effect. The ratio rises monotonically with threshold; the headline value of 1.106 (all data, 20 \u00b0C) is robust across the secondary specifications.' }),
  ...table(
    ['Specification', 'Heat-day ratio'],
    [
      ['Threshold 18 \u00b0C', '1.069'],
      ['Threshold 20 \u00b0C (primary)', '1.106'],
      ['Threshold 22 \u00b0C', '1.152'],
      ['Exclude 2021', '1.099'],
      ['Males only', '1.111'],
      ['Females only', '1.100'],
      ['All (reference)', '1.106'],
    ],
    { tableNum: 'Table 3.3', title: 'Sensitivity of the heat-day mortality ratio to threshold and stratification.', note: 'Ratios are mean daily deaths per 100,000 population on heat days divided by mean daily deaths per 100,000 on non-heat summer (June\u2013September) days. The 25 \u00b0C threshold is degenerate (zero qualifying days under the three-consecutive-days rule) and is omitted.' },
  ),
);

// Chapter 3 continued — spatial models
children.push(
  h2('3.5 Spatial models'),

  h3('3.5.1 Global Moran\u2019s I'),
  p('Pooled global Moran\u2019s I on excess heat-mortality per 100,000 population at the NUTS-1 / ITL-1 scale is +0.132 (E[I] = \u22120.111, z = +1.30, p = 0.208 from 999 conditional permutations). This pooled result is not significant at conventional thresholds. The per-year results, however, reveal that the spatial structure is heatwave-year-specific. Annual global Moran\u2019s I is +0.356 in 2017 (p = 0.013) and +0.300 in 2018 (p = 0.023); 2019 and 2021 register +0.229 (p = 0.065) and +0.247 (p = 0.061) respectively, marginal but trending positive; 2022\u2019s record summer registers only +0.055, suggesting that the geographically uniform exposure of the 2022 episode generated geographically uniform excess. Table 3.1 provides the full annual decomposition, and Figure 15 plots the annual trajectory.'),
  ...table(
    ['Year', 'Moran\u2019s I', 'z (sim)', 'p (sim)'],
    [
      ['2014', '+0.176', '+1.54', '0.111'],
      ['2015', '\u22120.090', '+0.16', '0.916'],
      ['2016', '+0.090', '+1.13', '0.298'],
      ['2017', '+0.356', '+2.55', '0.013'],
      ['2018', '+0.300', '+2.13', '0.023'],
      ['2019', '+0.229', '+1.85', '0.065'],
      ['2021', '+0.247', '+1.86', '0.061'],
      ['2022', '+0.055', '+0.99', '0.320'],
    ],
    { tableNum: 'Table 3.1', title: 'Annual global Moran\u2019s I on excess heat-mortality (Queen weights, 999 conditional permutations).' },
  ),
  ...fig('15_moran_trajectory.png', { figNum: 'Figure 15', title: 'Annual global Moran\u2019s I, 2014\u20132022.', note: 'Significant values (p < 0.05) at conventional thresholds are highlighted. 2017 and 2018, the two most intense pre-2022 heatwave years, are statistically significant; 2019 and 2021 are marginal; the 2022 record summer is non-significant, consistent with its geographically uniform exposure pattern.' }),
  p('Resilient across weight specifications: pooled Moran\u2019s I is +0.132 under both Queen and Rook contiguity but flips to \u22120.301 under k-nearest-neighbours (k = 3). The reversal under k-NN reflects the well-known sensitivity of small-N Moran\u2019s I to weights specification: k-NN imposes exactly three neighbours per region and connects pairs such as South West and Wales that are non-contiguous, distorting the local averaging. This sensitivity is reported transparently rather than concealed. The Queen and Rook results are the substantive headline; the k-NN result is a solid footnote.'),

  h3('3.5.2 Local Moran\u2019s I (LISA)'),
  p('Local Moran\u2019s I (Anselin, 1995) on pooled excess heat-mortality identifies the West Midlands as a statistically significant High\u2013High cluster (I\u1d62 = +0.523; p = 0.040 from 999 conditional permutations). East Midlands is borderline (I\u1d62 = +0.304; p = 0.093). North East shows a positive local I in the Low\u2013Low quadrant (I\u1d62 = +0.793) but does not reach significance at the 0.05 level (p = 0.159). The substantive interpretation is that excess heat-mortality forms a coherent Midlands cluster in which both elevated regional values and elevated neighbour values reinforce one another. Figure 13 maps the LISA classification.'),
  ...fig('13_lisa.png', { figNum: 'Figure 13', title: 'Local Moran\u2019s I (LISA) cluster map.', note: 'LISA classification of regions on pooled excess heat-mortality, Queen weights, 999 conditional permutations. West Midlands is the only region significant at p < 0.05 and is classified High\u2013High. The Anselin colour scheme indicates significance (red = HH, cyan = LL, light shades = HL/LH spatial outliers, grey = non-significant).' }),
  p('The spatial pattern visible in Figure 13 is corroborated by the choropleth of pooled excess heat-mortality (Figure 12), which shows the highest values concentrated in the Midlands and southern English regions, with London noticeably depressed despite its high density and low green space.'),
  ...fig('12_excess_choropleth.png', { figNum: 'Figure 12', title: 'Pooled excess heat-mortality per 100,000 population.', note: 'Cumulative excess deaths on heat days (Met Office 20 \u00b0C / 3-day definition) over 2014\u20132022, normalised to mid-year population. The Midlands cluster identified by LISA (Figure 13) and Getis-Ord (Figure 14) is visible directly in the values: East Midlands 28.7, West Midlands 26.2.' }),

  h3('3.5.3 Getis-Ord G*'),
  p('Getis-Ord G* (Getis & Ord, 1992) on pooled excess heat-mortality identifies East Midlands as a statistically significant hot spot at the 95 percent level (z = +2.14), West Midlands as a hot spot at the 90 percent level (z = +1.78), and North East as a cold spot at the 90 percent level (z = \u22121.83). These results are consistent with the LISA finding, identifying the Midlands as the locus of elevated heat-related excess mortality and the North East as its opposite, with the caveat, again, that with N = 10 areal units the available statistical resolution is limited. Figure 14 maps the G* z-scores.'),
  ...fig('14_getis_ord.png', { figNum: 'Figure 14', title: 'Getis-Ord G* statistic on pooled excess heat-mortality.', note: 'Local G* z-scores by region. Red shades indicate hot spots (positive z); cyan shades indicate cold spots (negative z); white indicates non-significant. East Midlands (z = +2.14) is a 95% hot spot; West Midlands (z = +1.78) is a 90% hot spot; North East (z = \u22121.83) is a 90% cold spot.' }),

  h3('3.5.4 Bivariate Moran\u2019s I: the spatial tests of H1 and H2'),
  p('Bivariate Moran\u2019s I assesses whether the spatial lag of one variable covaries with the level of another. The result is the spatial-frame test of the two substantive hypotheses. The bivariate I between green-space provision per capita and excess heat-mortality is +0.040 (p = 0.832); between distance to the nearest park and excess heat-mortality, +0.079 (p = 0.617); between percent built-up land and excess heat-mortality, +0.033 (p = 0.902). All three are statistically indistinguishable from zero. Substantively, H1 predicts a negative bivariate I between green-space provision and excess mortality (high green space neighbouring low mortality, and vice versa); the observed value is positive, the opposite of the predicted sign. H2 predicts a positive bivariate I between built-up land and excess mortality; the observed value is positive but trivial in magnitude and not statistically distinguishable from zero. Table 3.2 reports the three bivariate tests, and Figure 16 shows the underlying scatter that drives the result.'),
  ...table(
    ['Pair', 'Bivariate I', 'z', 'p', 'Verdict'],
    [
      ['Green-space provision \u00d7 excess', '+0.040', '+0.20', '0.832', 'H1 sign reversed; not significant'],
      ['Distance to park \u00d7 excess', '+0.079', '+0.51', '0.617', 'Inconclusive'],
      ['Per-cent built-up \u00d7 excess', '+0.033', '+0.17', '0.902', 'H2 not supported'],
    ],
    { tableNum: 'Table 3.2', title: 'Bivariate Moran\u2019s I: spatial tests of H1 and H2 (Queen weights, 999 permutations).' },
  ),
  ...fig('16_bivariate_scatter.png', { figNum: 'Figure 16', title: 'Bivariate scatter: H1 and H2 against pooled excess heat-mortality.', note: '(a) Green-space provision (FiT-comparable m\u00b2/cap) against pooled excess heat-mortality. The Pearson r of +0.59 indicates that, contrary to H1, regions with more green space tend to have higher heat-mortality, a pattern driven by London being simultaneously the lowest-green-space and lowest-excess-mortality region. (b) Per-cent built-up land against pooled excess heat-mortality. The Pearson r of \u22120.40 again indicates the H2-opposing direction, again driven by London. The dashed regression lines are visualisations of the OLS slopes; the bivariate Moran\u2019s I values (Table 3.2) tell the same substantive story under spatial inference.' }),

  h3('3.5.5 Spatial regression: OLS with LM diagnostics, SAR, SEM'),
  p('A standardised regression of excess heat-mortality per 100,000 population on green-space provision and per-cent built-up land is fitted by ordinary least squares; Lagrange Multiplier diagnostics for spatial lag and spatial error are computed against the Queen weights matrix; and maximum-likelihood spatial-lag (SAR) and spatial-error (SEM) variants are estimated for completeness. The OLS fit yields R\u00b2 = 0.354, a green-space coefficient of +0.708 (standard error 0.49) and a built-up coefficient of +0.155 (standard error 0.49). The LM-lag statistic is 0.18 (p = 0.672); the LM-error statistic is 0.005 (p = 0.946). Neither diagnostic rejects the null of no residual spatial dependence beyond what the covariates absorb. In substantive terms, the spatial-regression diagnostics agree with the bivariate Moran result: there is no residual spatial structure once green-space and built-up land are in the model, and a spatial-econometric specification does not absorb the misspecification.'),
  p('The SAR and SEM estimates are included for completeness. The SAR rho is +0.348 with green-space coefficient +0.649 and built-up coefficient +0.097 (AIC 33.44). The SEM lambda is \u22120.893 with green-space coefficient +1.204 and built-up coefficient +0.203 (AIC 33.92). The signs and ordering of the substantive coefficients are preserved across the three specifications, reinforcing the conclusion that the confounded direction of the green-space coefficient is not driven by misspecification of the spatial dependence. This is the principal diagnostic finding of Section 3.5: the wrong sign on H1 is robust to spatial-econometric correction, indicating that the underlying problem is confounding rather than spatial-autocorrelation misspecification.'),

  h2('3.6 The baseline OLS transparency reference'),
  p('The symmetry of the coefficients in this paper (Section 3.6.1 below), is shown to be an artefact of separate single-predictor standardised regressions on the same response rather than a sign that the covariates were mechanically derived from one another (which the data audit considered as a possible explanation). The substantive results from this region-level OLS are reported here for transparency, with the caveat that they are confounded by demographic structure, specifically the London paradox, and should not be interpreted as evidence on H1 or H2.'),

  h3('3.6.1 Region-level OLS ; \u201cdo not do this\u201d reference'),
  p('Three single-predictor standardised regressions of mean heat-day mortality per 100,000 population (N = 10 regions) on the H1 and H2 candidate predictors yield: green-space provision per capita, \u03b2 = +0.714 (t = +2.88, p = 0.020, R\u00b2 = 0.510); average distance to park, \u03b2 = +0.686 (t = +2.67, p = 0.028, R\u00b2 = 0.471); per-cent built-up land, \u03b2 = \u22120.966 (t = \u221210.54, p < 0.001, R\u00b2 = 0.933). The directions are the opposite of those predicted by H1 (more green space \u2192 higher mortality) and H2 (more built-up \u2192 lower mortality), and are dominated by the London observation. The replication retains these results as a cautionary example: they illustrate why a properly-specified model with population offset, seasonal control and regional fixed effects is necessary before any substantive interpretation of regional differences in mortality can be defended. The bivariate scatter in Figure 16 visualises the London paradox directly: London is the bottom-left outlier in panel (a) and the bottom-right outlier in panel (b), and its position pulls the OLS slope away from the H1- and H2-predicted directions.'),

  h2('3.7 Summary of empirical findings'),
  p('Combining the temporal and spatial frames, the empirical results can be summarised as follows. First, a heat\u2013mortality effect at the national daily-mean scale is identified and quantified (DLNM cumulative RR 1.045 at the 99th percentile versus MMT). Second, excess heat-mortality clusters spatially in heatwave years (annual Moran\u2019s I significant in 2017 and 2018; West Midlands LISA HH cluster; East Midlands Getis-Ord Gi* hot spot). Third, neither H1 nor H2 is supported in the spatial frame at NUTS-1 / ITL-1 resolution: the bivariate Moran\u2019s I between green-space provision and excess mortality has the opposite sign from that predicted by H1, and the bivariate I between built-up land and excess mortality is statistically indistinguishable from zero. Fourth, spatial regression Lagrange Multiplier diagnostics fail to reject the null of no residual spatial dependence, indicating that the observed confounded relationship is not driven by misspecification of the spatial structure but by an unmeasured confounder operating at the regional scale.'),

  h2('3.8 Discussion: the London paradox and confounding'),
  p('The reversal of H1\u2019s sign in the bivariate Moran\u2019s I result (and in the region-level OLS) admits a clear substantive explanation. London is simultaneously the region with the highest population density, the lowest per-capita public green-space provision, and the lowest crude all-cause mortality per 100,000 population (1.58 per day in this panel; compare the North East at 2.97 and Wales at 3.00). London\u2019s low crude mortality reflects its substantially younger demographic structure (Office for National Statistics, 2023), and not any protective effect of density or scarcity of green space. At ten observations, this configuration is sufficient to invert the population-level coefficient of green space against mortality.'),
  p('The age-structure confounding is not, on this data and at this resolution, possible to adjust for. The supplied mortality dataset does not include age stratification; regional population age-structure adjustment is feasible from the ONS Mid-Year Estimates by single year of age, but in the spatial-frame at NUTS-1 resolution there are still only ten residual degrees of freedom for any covariate that varies between regions, and an age-adjusted regional analysis would be subject to the same identification limitations.'),
  p('Three additional confounders are worth naming. First, deprivation: the English Index of Multiple Deprivation 2019 documents a substantial gradient in vulnerability across small areas, and the regional aggregation conceals within-region variation that is plausibly larger than between-region variation (Public Health England, 2019). Second, dwelling characteristics: housing-stock characteristics such as roof-space insulation, glazing, and air-conditioning prevalence vary substantially within regions but are not directly measured here. Third, the coastal effect: regions with a long coastline (East of England, South West) have systematically lower summer maximum temperatures than landlocked regions of similar latitude, an effect not captured by the single national Central England Temperature exposure series used here.'),
  p('Compared to the published international literature, the heat-attributable relative risks estimated in this thesis (RR 1.045 cumulative over 0\u201314 days at the 99th percentile of daily-mean temperature versus MMT) are conservative relative to estimates using daily-maximum temperature (Gasparrini et al., 2015) and are similar in order of magnitude to UK-specific national-level estimates by Hajat and colleagues (Hajat et al., 2014) and Armstrong and colleagues (Armstrong et al., 2011). The much larger cold-burden estimate (RR 1.385) is consistent with all major DLNM-based UK studies. The novel finding of this paper is therefore not in the temporal-frame magnitude estimates but in the negative result on H1 and H2 in the spatial frame, and in the identification of why the negative result obtains.'),
  p('A final limitation, internal to the analysis itself: the temperature exposure measure is a single national series (HadCET) broadcast to all regions. Regional contrasts in heat-related mortality estimated in this thesis cannot be attributed to regional temperature differences, because the source data does not measure such differences; the regional contrasts must arise from baseline mortality (the offset), population structure, and the H1 / H2 covariates. The MODIS Land Surface Temperature satellite dataset at 1 km resolution would provide the requisite regional exposure contrast in a follow-up study; see Section 4.3.'),

  h2('3.9 The 2020 exclusion'),
  p('The 2020 calendar year is excluded from the analysis because the mortality outcome in 2020 is dominated by the COVID-19 pandemic, which generated an excess-mortality shock of a magnitude that would corrupt any temperature\u2013mortality estimation. The Office for National Statistics records approximately 88,500 excess deaths in England and Wales in 2020 relative to the 2015\u20132019 five-year average, a 13.7 percent rise in all-cause mortality and the largest single-year excess registered in England and Wales since the Second World War (Office for National Statistics, 2021). Daily mortality in April 2020 alone reached more than twice the five-year average, with the principal peaks driven by COVID-19 transmission and severity in care-home and hospital settings rather than by environmental exposure. A second wave through autumn and winter 2020\u201321 added a further ~30,000 excess deaths concentrated in November 2020 through February 2021. None of this excess is heat attributable.'),
  p('Including 2020 in the DLNM panel would create three identification problems. First, the COVID-19 excess is an order of magnitude larger than any plausible heat signal, and the daily-mortality time series in 2020 is dominated by epidemic-curve dynamics that have no causal relationship to ambient temperature. Second, the temporal coincidence between the second COVID-19 wave (autumn/winter) and the seasonal cold-mortality peak would inflate the apparent cold-attributable mortality fraction, the natural-cubic-spline on continuous time at eight degrees of freedom per year would partially absorb the COVID-19 wave but would not separate it cleanly from the underlying temperature signal. Third, the natural lockdown-induced reductions in road-traffic injury and air-pollution exposure produced their own non-thermal mortality reduction in spring and summer 2020, in the opposite direction to the COVID-19 surge; the net effect on the temperature-mortality estimate is unpredictable and would require an explicit pandemic-confounder model that exceeds the scope of this thesis.'),
  p('The HadCET v2.1.1.0 series covers 2020 continuously and includes the August 2020 heatwave (a six-day episode with daily-mean Central England Temperature peaking at 22.9 \u00b0C on 11 August). HadCET data for 2020 are used in the calibration of the rolling baseline (Section 3.4.3 uses a five-year-equivalent window that spans 2017\u20132021 for the 2022 reference), but the 2020 mortality outcome is not in the analysis panel. The August 2020 heatwave, which produced an estimated 1,964 excess heat-attributable deaths in England by Public Health England\u2019s heatwave mortality monitoring (Public Health England, 2020), is therefore not contributing to the heat-day exposure set used to estimate the cumulative heat relative risk reported in Section 3.4.2. This is a real analytical cost: 2020 was the second-hottest summer in the analysis period after 2022, and its exclusion narrows the heat-day sample. The exclusion is nonetheless preferable to the alternative, pandemic-confounded daily mortality contaminating the entire DLNM exposure-response estimation, and the DLNM cumulative heat RR estimated on the eight remaining years (1.045 [1.031, 1.058]) is consistent with the published UK literature on equivalent samples that exclude 2020 for the same reason (Masselot et al., 2023; Vicedo-Cabrera et al., 2021).'),
);

// =================================================================
// CHAPTER 4 — Recommendations
// =================================================================
children.push(
  h1('Chapter 4. Recommendations'),

  h2('4.1 For public-health practice'),
  p('Three recommendations follow the temporal-frame results. First, the cumulative excess heat-mortality estimated in this thesis confirms that the United Kingdom Health Security Agency Heatwave Plan threshold of 18 \u00b0C daily-mean temperature is a defensible operational trigger: the heat-mortality ratio at this threshold (1.069) is statistically distinguishable from unity, and the threshold captures sufficient exposure to justify Level 2 alert activation. The thesis offers no basis for raising or lowering this threshold.'),
  p('Second, the disproportionate East and West Midlands excess identified in the spatial-frame results (Section 3.5) supports targeting heat-resilience interventions, cooling-centre provision, vulnerable-resident outreach, and dwelling-stock upgrades, at sub-regional units within these regions, rather than at the regional level uniformly. The geographic specificity of the Midlands High\u2013High cluster identified in the LISA and Getis-Ord analyses is sufficient to motivate regionally targeted intervention design, even at the limited statistical resolution of N = 10.'),
  p('Third, the cold-burden estimate (RR 1.385 at 0 \u00b0C versus MMT) is several times larger than the heat burden and serves as a reminder that a comprehensive temperature\u2013health strategy must address both tails of the temperature distribution. The Cold Weather Plan for England (Department of Health and Social Care, 2018) is the appropriate companion to the Heatwave Plan, and the paper estimates support continued investment in cold-weather preparedness alongside heatwave preparedness; in fact, the cold burden is sufficiently larger in absolute terms that an integrated annual temperature-health strategy is arguably the most cost-effective public-health investment in this domain.'),

  h2('4.2 For urban planning'),
  p('The H1 and H2 negative results do not contradict the established cooling effects of urban green space and the documented warming effect of high building density (Bowler et al., 2010; Heaviside et al., 2016); they indicate that at NUTS-1 / ITL-1 resolution the associated mortality signal is overwhelmed by demographic confounding. Two implications follow. First, urban-planning policy guidance on green-space and density should rely on the established microclimatic literature rather than on regional-scale mortality regressions of the kind reported here. The microclimatic mechanism is empirically supported by direct temperature measurements in the relevant settings; the mortality consequences of that mechanism are real but are best estimated at the spatial scale on which the mechanism operates (the building, the street, the neighbourhood), not at the scale of the entire NUTS-1 region.'),
  p('Second, the existence of a Midlands High\u2013High cluster suggests that whichever municipal planning bodies are responsible for the West Midlands and East Midlands metropolitan areas should treat the cluster as a coordinated unit for heat-resilience purposes. Coordinated regional planning around heat resilience, sharing best practice, pooling resources for vulnerable-resident outreach, and coordinating mutual-aid arrangements for cooling-centre capacity during extreme events, is empirically supported by the spatial clustering result, even with the N = 10 caveat.'),

  h2('4.3 For future research'),
  p('The single most consequential next step is sub-regional analysis. Specifically, replication at Lower Layer Super Output Area (LSOA) resolution, with adjustment for area-level age structure and the English Index of Multiple Deprivation 2019, is necessary to test H1 and H2 in a setting where the London paradox can be statistically controlled. The data required for such an analysis are: LSOA-level monthly mortality counts (available from the Office for National Statistics with appropriate research access); LSOA-level OS Open Greenspace per-capita area (computable using the same pipeline as in this thesis at the higher resolution); LSOA-level percent built-up land (likewise); LSOA-level IMD 2019 deciles (publicly available); and an LSOA-level temperature exposure measure derived from MODIS Land Surface Temperature satellite imagery (publicly available; would replace the single national HadCET series with a within-region spatial exposure contrast).'),
  p('The expected sample at LSOA resolution is approximately 34,000 small areas. The Masselot and colleagues (Masselot et al., 2023) LSOA-level heat-attributable-mortality estimates would constitute an external comparator against which a replicated H1/H2 test could be benchmarked. The methodological contribution of the thesis, the use of an open-data per-capita green-space measure, with end-to-end reproducibility from publicly available sources, transfers to the LSOA-level analysis directly. Statistical power at LSOA resolution is sufficient to detect the H1 and H2 effects at clinically meaningful magnitudes, conditional on the assumption that the mechanism operates as the microclimatic literature suggests.'),
  p('Two secondary lines of work are worth flagging. First, the daily-maximum temperature exposure series: HadCET also publishes the daily maximum, and heat-RR estimates against daily-maximum would be substantially larger than the daily-mean values reported here (RR 1.10\u20131.20 at the 99th percentile of daily-max is the order-of-magnitude expectation). The pipeline is otherwise identical and could be conducted as a methods note. Second, the cold burden: the paper reports the cold burden as a side product of the DLNM. A standalone analysis of the cold-mortality burden in England and Wales, interacting with fuel-poverty data from the Department for Energy Security and Net Zero, would have direct policy relevance and would use the same toolchain.'),
);

// =================================================================
// CONCLUSION
// =================================================================
children.push(
  h1('Conclusion'),
  p('This thesis tested two substantive hypotheses about heat-related mortality at NUTS-1 / ITL-1 resolution in England and Wales over 2014\u20132022. A heat\u2013mortality effect was identified and quantified (DLNM cumulative relative risk 1.045 at the 99th percentile of daily-mean temperature versus the minimum-mortality temperature of 18.3 \u00b0C, 95 percent CI [1.031, 1.058]). Excess heat-mortality was shown to cluster spatially in heatwave years (Moran\u2019s I positive and significant in 2017 and 2018; West Midlands LISA High\u2013High cluster; East Midlands Getis-Ord Gi* hot spot). The two substantive hypotheses, however, were not supported: bivariate Moran\u2019s I between green-space provision and excess mortality was +0.04 (opposite sign to H1, not significant), and between building density and excess mortality was +0.03 (consistent with H2 in sign but not significant); spatial-regression Lagrange Multiplier tests failed to reject the null of no residual spatial dependence.'),
  p('The failure to support H1 and H2 is attributable to confounding by demographic structure and urbanicity that operates geographically, specifically through the London paradox: London is simultaneously the densest, least green-space-provisioned and lowest-crude-mortality region. The principal recommendation for future research is replication at Lower Layer Super Output Area resolution with age and deprivation adjustment; the principal methodological contribution is the demonstration that the proprietary Fields in Trust Green Space Index can be substituted by an open-data measure with identical regional ordering, supporting end-to-end reproducibility from publicly available sources.'),
  p('More broadly, the thesis illustrates a recurring methodological lesson in environmental epidemiology: the spatial scale on which a moderator hypothesis is tested must match the spatial scale on which the underlying mechanism operates. The cooling effect of urban green space and the warming effect of dense built fabric are real microclimatic phenomena, but their consequences for mortality cannot be cleanly identified at the resolution of NUTS-1 regions because regional aggregates average out the within-region variation that the mechanism produces. The Lower Layer Super Output Area resolution, approximately one observation per 1,500 people, is the natural scale at which the moderator hypotheses can be tested, and the methodological tooling assembled in this thesis transfers to that scale without modification.'),
);

// =================================================================
// REFERENCES (APA 7, hanging indent, hyperlinked DOIs)
// =================================================================
children.push(
  h1('References'),
  refEntry([
    { text: 'Anselin, L. (1988). ' },
    { text: 'Spatial econometrics: Methods and models', italics: true },
    { text: '. Kluwer Academic Publishers.' },
  ]),
  refEntry([
    { text: 'Anselin, L. (1995). Local indicators of spatial association, LISA. ' },
    { text: 'Geographical Analysis, 27', italics: true },
    { text: '(2), 93\u2013115. ' },
    { text: 'https://doi.org/10.1111/j.1538-4632.1995.tb00338.x', url: 'https://doi.org/10.1111/j.1538-4632.1995.tb00338.x' },
  ]),
  refEntry([
    { text: 'Anselin, L., & Rey, S. J. (2014). ' },
    { text: 'Modern spatial econometrics in practice: A guide to GeoDa, GeoDaSpace and PySAL', italics: true },
    { text: '. GeoDa Press.' },
  ]),
  refEntry([
    { text: 'Aram, F., Higueras Garc\u00eda, E., Solgi, E., & Mansournia, S. (2019). Urban green space cooling effect in cities. ' },
    { text: 'Heliyon, 5', italics: true },
    { text: '(4), e01339. ' },
    { text: 'https://doi.org/10.1016/j.heliyon.2019.e01339', url: 'https://doi.org/10.1016/j.heliyon.2019.e01339' },
  ]),
  refEntry([
    { text: 'Armstrong, B. G., Chalabi, Z., Fenn, B., Hajat, S., Kovats, S., Milojevic, A., & Wilkinson, P. (2011). Association of mortality with high temperatures in a temperate climate: England and Wales. ' },
    { text: 'Journal of Epidemiology and Community Health, 65', italics: true },
    { text: '(4), 340\u2013345. ' },
    { text: 'https://doi.org/10.1136/jech.2009.093161', url: 'https://doi.org/10.1136/jech.2009.093161' },
  ]),
  refEntry([
    { text: 'Bennett, J. E., Blangiardo, M., Fecht, D., Elliott, P., & Ezzati, M. (2014). Vulnerability to the mortality effects of warm temperature in the districts of England and Wales. ' },
    { text: 'Nature Climate Change, 4', italics: true },
    { text: '(4), 269\u2013273. ' },
    { text: 'https://doi.org/10.1038/nclimate2123', url: 'https://doi.org/10.1038/nclimate2123' },
  ]),
  refEntry([
    { text: 'Bowler, D. E., Buyung-Ali, L., Knight, T. M., & Pullin, A. S. (2010). Urban greening to cool towns and cities: A systematic review of the empirical evidence. ' },
    { text: 'Landscape and Urban Planning, 97', italics: true },
    { text: '(3), 147\u2013155. ' },
    { text: 'https://doi.org/10.1016/j.landurbplan.2010.05.006', url: 'https://doi.org/10.1016/j.landurbplan.2010.05.006' },
  ]),
  refEntry([
    { text: 'Department of Health. (2004). ' },
    { text: 'Heatwave plan for England 2004', italics: true },
    { text: '.' },
  ]),
  refEntry([
    { text: 'Department of Health and Social Care. (2018). ' },
    { text: 'The cold weather plan for England 2018', italics: true },
    { text: '.' },
  ]),
  refEntry([
    { text: 'de\u2019 Donato, F. K., Leone, M., Scortichini, M., De Sario, M., Katsouyanni, K., Lanki, T., Basagana, X., Ballester, F., Astrom, C., Paldy, A., Pascal, M., Gasparrini, A., Menne, B., & Michelozzi, P. (2015). Changes in the effect of heat on mortality in the last 20 years in nine European cities. ' },
    { text: 'International Journal of Environmental Research and Public Health, 12', italics: true },
    { text: '(12), 15567\u201315583. ' },
    { text: 'https://doi.org/10.3390/ijerph121215006', url: 'https://doi.org/10.3390/ijerph121215006' },
  ]),
  refEntry([
    { text: 'Fields in Trust. (2020). ' },
    { text: 'Green Space Index 2020', italics: true },
    { text: ' [Report]. ' },
    { text: 'https://www.fieldsintrust.org/green-space-index', url: 'https://www.fieldsintrust.org/green-space-index' },
  ]),
  refEntry([
    { text: 'Gasparrini, A. (2011). Distributed lag linear and non-linear models in R: The package dlnm. ' },
    { text: 'Journal of Statistical Software, 43', italics: true },
    { text: '(8), 1\u201320. ' },
    { text: 'https://doi.org/10.18637/jss.v043.i08', url: 'https://doi.org/10.18637/jss.v043.i08' },
  ]),
  refEntry([
    { text: 'Gasparrini, A. (2014). Modeling exposure\u2013lag\u2013response associations with distributed lag non-linear models. ' },
    { text: 'Statistics in Medicine, 33', italics: true },
    { text: '(5), 881\u2013899. ' },
    { text: 'https://doi.org/10.1002/sim.5963', url: 'https://doi.org/10.1002/sim.5963' },
  ]),
  refEntry([
    { text: 'Gasparrini, A., Armstrong, B., & Kenward, M. G. (2010). Distributed lag non-linear models. ' },
    { text: 'Statistics in Medicine, 29', italics: true },
    { text: '(21), 2224\u20132234. ' },
    { text: 'https://doi.org/10.1002/sim.3940', url: 'https://doi.org/10.1002/sim.3940' },
  ]),
  refEntry([
    { text: 'Gasparrini, A., Guo, Y., Hashizume, M., Lavigne, E., Zanobetti, A., Schwartz, J., Tobias, A., Tong, S., R\u00f6cklov, J., Forsberg, B., Leone, M., De Sario, M., Bell, M. L., Guo, Y.-L. L., Wu, C.-F., Kan, H., Yi, S.-M., de Sousa Zanotti Stagliorio Coelho, M., Saldiva, P. H. N., \u2026 Armstrong, B. (2015). Mortality risk attributable to high and low ambient temperature: A multicountry observational study. ' },
    { text: 'The Lancet, 386', italics: true },
    { text: '(9991), 369\u2013375. ' },
    { text: 'https://doi.org/10.1016/S0140-6736(14)62114-0', url: 'https://doi.org/10.1016/S0140-6736(14)62114-0' },
  ]),
  refEntry([
    { text: 'Getis, A., & Ord, J. K. (1992). The analysis of spatial association by use of distance statistics. ' },
    { text: 'Geographical Analysis, 24', italics: true },
    { text: '(3), 189\u2013206. ' },
    { text: 'https://doi.org/10.1111/j.1538-4632.1992.tb00261.x', url: 'https://doi.org/10.1111/j.1538-4632.1992.tb00261.x' },
  ]),
  refEntry([
    { text: 'Hajat, S., Kovats, R. S., & Lachowycz, K. (2007). Heat-related and cold-related deaths in England and Wales: Who is at risk? ' },
    { text: 'Occupational and Environmental Medicine, 64', italics: true },
    { text: '(2), 93\u2013100. ' },
    { text: 'https://doi.org/10.1136/oem.2006.029017', url: 'https://doi.org/10.1136/oem.2006.029017' },
  ]),
  refEntry([
    { text: 'Hajat, S., Vardoulakis, S., Heaviside, C., & Eggen, B. (2014). Climate change effects on human health: Projections of temperature-related mortality for the UK during the 2020s, 2050s and 2080s. ' },
    { text: 'Journal of Epidemiology and Community Health, 68', italics: true },
    { text: '(7), 641\u2013648. ' },
    { text: 'https://doi.org/10.1136/jech-2013-202449', url: 'https://doi.org/10.1136/jech-2013-202449' },
  ]),
  refEntry([
    { text: 'Heaviside, C., Macintyre, H., & Vardoulakis, S. (2016). The urban heat island: Implications for health in a changing environment. ' },
    { text: 'Current Environmental Health Reports, 4', italics: true },
    { text: '(3), 296\u2013305. ' },
    { text: 'https://doi.org/10.1007/s40572-017-0150-3', url: 'https://doi.org/10.1007/s40572-017-0150-3' },
  ]),
  refEntry([
    { text: 'IPCC. (2021). ' },
    { text: 'Climate change 2021: The physical science basis. Contribution of Working Group I to the Sixth Assessment Report of the Intergovernmental Panel on Climate Change', italics: true },
    { text: ' (V. Masson-Delmotte, P. Zhai, A. Pirani, S. L. Connors, C. P\u00e9an, S. Berger, N. Caud, Y. Chen, L. Goldfarb, M. I. Gomis, M. Huang, K. Leitzell, E. Lonnoy, J. B. R. Matthews, T. K. Maycock, T. Waterfield, O. Yelek\u00e7i, R. Yu, & B. Zhou, Eds.). Cambridge University Press.' },
  ]),
  refEntry([
    { text: 'Johnson, H., Kovats, R. S., McGregor, G., Stedman, J., Gibbs, M., & Walton, H. (2005). The impact of the 2003 heat wave on mortality and hospital admissions in England. ' },
    { text: 'Health Statistics Quarterly, 25', italics: true },
    { text: ', 6\u201311.' },
  ]),
  refEntry([
    { text: 'Legg, T., Parker, D., & Brohan, P. (2025). The maintenance and continuation of the Hadley Centre Central England Temperature series. ' },
    { text: 'International Journal of Climatology', italics: true },
    { text: '. Advance online publication.' },
  ]),
  refEntry([
    { text: 'Lowe, J. A., Bernie, D., Bett, P., Bricheno, L., Brown, S., Calvert, D., Clark, R., Eagle, K., Edwards, T., Fosser, G., Fung, F., Gohar, L., Good, P., Gregory, J., Harris, G., Howard, T., Kaye, N., Kendon, E., Krijnen, J., \u2026 Belcher, S. (2018). ' },
    { text: 'UKCP18 science overview report', italics: true },
    { text: '. Met Office. ' },
    { text: 'https://www.metoffice.gov.uk/pub/data/weather/uk/ukcp18', url: 'https://www.metoffice.gov.uk/pub/data/weather/uk/ukcp18' },
  ]),
  refEntry([
    { text: 'Macintyre, H. L., Heaviside, C., Taylor, J., Picetti, R., Symonds, P., Cai, X.-M., & Vardoulakis, S. (2018). Assessing urban population vulnerability and environmental risks across an urban area during heatwaves, implications for health protection. ' },
    { text: 'Science of the Total Environment, 610\u2013611', italics: true },
    { text: ', 678\u2013690. ' },
    { text: 'https://doi.org/10.1016/j.scitotenv.2017.08.062', url: 'https://doi.org/10.1016/j.scitotenv.2017.08.062' },
  ]),
  refEntry([
    { text: 'Masselot, P., Mistry, M. N., Rao, S., Schneider, R., Aunan, K., Floris, M., \u2026 Gasparrini, A. (2023). Small-area assessment of temperature-related mortality risks in England and Wales: A case time series analysis. ' },
    { text: 'The Lancet Planetary Health, 7', italics: true },
    { text: '(4), e271\u2013e279. ' },
    { text: 'https://doi.org/10.1016/S2542-5196(23)00037-1', url: 'https://doi.org/10.1016/S2542-5196(23)00037-1' },
  ]),
  refEntry([
    { text: 'Met Office. (2022, July 19). ' },
    { text: 'Record-breaking temperatures for July', italics: true },
    { text: ' [Press release]. ' },
    { text: 'https://www.metoffice.gov.uk/about-us/press-office/news/weather-and-climate/2022/record-breaking-temperatures-for-july', url: 'https://www.metoffice.gov.uk/about-us/press-office/news/weather-and-climate/2022/record-breaking-temperatures-for-july' },
  ]),
  refEntry([
    { text: 'Met Office. (2025). ' },
    { text: 'HadCET: Hadley Centre Central England Temperature dataset, v2.1.1.0', italics: true },
    { text: ' [Dataset]. NERC EDS Centre for Environmental Data Analysis. ' },
    { text: 'https://doi.org/10.5285/4b541f71524244c5a19f7d8321b5563d', url: 'https://doi.org/10.5285/4b541f71524244c5a19f7d8321b5563d' },
  ]),
  refEntry([
    { text: 'Moran, P. A. P. (1948). The interpretation of statistical maps. ' },
    { text: 'Journal of the Royal Statistical Society: Series B (Methodological), 10', italics: true },
    { text: '(2), 243\u2013251. ' },
    { text: 'https://doi.org/10.1111/j.2517-6161.1948.tb00012.x', url: 'https://doi.org/10.1111/j.2517-6161.1948.tb00012.x' },
  ]),
  refEntry([
    { text: 'Office for National Statistics. (2020). ' },
    { text: 'Access to gardens and public green space in Great Britain: April 2020 reference tables', italics: true },
    { text: ' [Dataset]. ' },
    { text: 'https://www.ons.gov.uk/economy/environmentalaccounts/datasets/accesstogardensandpublicgreenspaceingreatbritain', url: 'https://www.ons.gov.uk/economy/environmentalaccounts/datasets/accesstogardensandpublicgreenspaceingreatbritain' },
  ]),
  refEntry([
    { text: 'Office for National Statistics. (2021). ' },
    { text: 'Deaths registered in England and Wales: 2020', italics: true },
    { text: ' [Dataset]. ' },
    { text: 'https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/bulletins/deathsregisteredinenglandandwalesseriesdr/2020', url: 'https://www.ons.gov.uk/peoplepopulationandcommunity/birthsdeathsandmarriages/deaths/bulletins/deathsregisteredinenglandandwalesseriesdr/2020' },
  ]),
  refEntry([
    { text: 'Office for National Statistics. (2022). ' },
    { text: 'Deaths registered weekly in England and Wales, provisional', italics: true },
    { text: ' [Dataset].' },
  ]),
  refEntry([
    { text: 'Office for National Statistics. (2023). ' },
    { text: 'Population estimates for the UK, England, Wales, Scotland and Northern Ireland: Mid-2023', italics: true },
    { text: ' [Dataset]. ' },
    { text: 'https://www.ons.gov.uk/peoplepopulationandcommunity/populationandmigration/populationestimates/datasets/populationestimatesforukenglandandwalesscotlandandnorthernireland', url: 'https://www.ons.gov.uk/peoplepopulationandcommunity/populationandmigration/populationestimates/datasets/populationestimatesforukenglandandwalesscotlandandnorthernireland' },
  ]),
  refEntry([
    { text: 'Oke, T. R. (1982). The energetic basis of the urban heat island. ' },
    { text: 'Quarterly Journal of the Royal Meteorological Society, 108', italics: true },
    { text: '(455), 1\u201324. ' },
    { text: 'https://doi.org/10.1002/qj.49710845502', url: 'https://doi.org/10.1002/qj.49710845502' },
  ]),
  refEntry([
    { text: 'Ordnance Survey. (2024). ' },
    { text: 'OS Open Greenspace', italics: true },
    { text: ' [Dataset]. ' },
    { text: 'https://www.ordnancesurvey.co.uk/products/os-open-greenspace', url: 'https://www.ordnancesurvey.co.uk/products/os-open-greenspace' },
  ]),
  refEntry([
    { text: 'Ordnance Survey. (2024). ' },
    { text: 'OS Open Built Up Areas', italics: true },
    { text: ' [Dataset]. ' },
    { text: 'https://www.ordnancesurvey.co.uk/products/os-open-built-up-areas', url: 'https://www.ordnancesurvey.co.uk/products/os-open-built-up-areas' },
  ]),
  refEntry([
    { text: 'Parker, D. E., Legg, T. P., & Folland, C. K. (1992). A new daily Central England Temperature series, 1772\u20131991. ' },
    { text: 'International Journal of Climatology, 12', italics: true },
    { text: '(4), 317\u2013342. ' },
    { text: 'https://doi.org/10.1002/joc.3370120402', url: 'https://doi.org/10.1002/joc.3370120402' },
  ]),
  refEntry([
    { text: 'Public Health England. (2019). ' },
    { text: 'Heatwave plan for England: Protecting health and reducing harm from severe heat and heatwaves', italics: true },
    { text: '.' },
  ]),
  refEntry([
    { text: 'Public Health England. (2020). ' },
    { text: 'PHE heatwave mortality monitoring report: Summer 2020', italics: true },
    { text: '. ' },
    { text: 'https://www.gov.uk/government/publications/phe-heatwave-mortality-monitoring', url: 'https://www.gov.uk/government/publications/phe-heatwave-mortality-monitoring' },
  ]),
  refEntry([
    { text: 'Reid, C. E., O\u2019Neill, M. S., Gronlund, C. J., Brines, S. J., Brown, D. G., Diez-Roux, A. V., & Schwartz, J. (2009). Mapping community determinants of heat vulnerability. ' },
    { text: 'Environmental Health Perspectives, 117', italics: true },
    { text: '(11), 1730\u20131736. ' },
    { text: 'https://doi.org/10.1289/ehp.0900683', url: 'https://doi.org/10.1289/ehp.0900683' },
  ]),
  refEntry([
    { text: 'Robine, J.-M., Cheung, S. L. K., Le Roy, S., Van Oyen, H., Griffiths, C., Michel, J.-P., & Herrmann, F. R. (2008). Death toll exceeded 70,000 in Europe during the summer of 2003. ' },
    { text: 'Comptes Rendus Biologies, 331', italics: true },
    { text: '(2), 171\u2013178. ' },
    { text: 'https://doi.org/10.1016/j.crvi.2007.12.001', url: 'https://doi.org/10.1016/j.crvi.2007.12.001' },
  ]),
  refEntry([
    { text: 'Romanello, M., Di Napoli, C., Green, C., Kennard, H., Lampard, P., Scamman, D., \u2026 Costello, A. (2023). The 2023 report of the Lancet Countdown on health and climate change: The imperative for a health-centred response in a world facing irreversible harms. ' },
    { text: 'The Lancet, 402', italics: true },
    { text: '(10419), 2346\u20132394. ' },
    { text: 'https://doi.org/10.1016/S0140-6736(23)01859-7', url: 'https://doi.org/10.1016/S0140-6736(23)01859-7' },
  ]),
  refEntry([
    { text: 'Tsikudo, G. K. (2024). ' },
    { text: 'Spatiotemporal analysis of heatwave-related mortality in the UK (England and Wales)', italics: true },
    { text: ' [Master\u2019s thesis, original draft]. Higher School of Economics, Moscow.' },
  ]),
  refEntry([
    { text: 'Vandentorren, S., Bretin, P., Zeghnoun, A., Mandereau-Bruno, L., Croisier, A., Cochet, C., Riberon, J., Siberan, I., Declercq, B., & Ledrans, M. (2006). August 2003 heat wave in France: Risk factors for death of elderly people living at home. ' },
    { text: 'European Journal of Public Health, 16', italics: true },
    { text: '(6), 583\u2013591. ' },
    { text: 'https://doi.org/10.1093/eurpub/ckl063', url: 'https://doi.org/10.1093/eurpub/ckl063' },
  ]),
  refEntry([
    { text: 'Vardoulakis, S., Dear, K., Hajat, S., Heaviside, C., Eggen, B., & McMichael, A. J. (2014). Comparative assessment of the effects of climate change on heat- and cold-related mortality in the United Kingdom and Australia. ' },
    { text: 'Environmental Health Perspectives, 122', italics: true },
    { text: '(12), 1285\u20131292. ' },
    { text: 'https://doi.org/10.1289/ehp.1307524', url: 'https://doi.org/10.1289/ehp.1307524' },
  ]),
  refEntry([
    { text: 'Vicedo-Cabrera, A. M., Scovronick, N., Sera, F., Roy\u00e9, D., Schneider, R., Tobias, A., \u2026 Gasparrini, A. (2021). The burden of heat-related mortality attributable to recent human-induced climate change. ' },
    { text: 'Nature Climate Change, 11', italics: true },
    { text: '(6), 492\u2013500. ' },
    { text: 'https://doi.org/10.1038/s41558-021-01058-x', url: 'https://doi.org/10.1038/s41558-021-01058-x' },
  ]),
  refEntry([
    { text: 'Wilby, R. L. (2003). Past and projected trends in London\u2019s urban heat island. ' },
    { text: 'Weather, 58', italics: true },
    { text: '(7), 251\u2013260. ' },
    { text: 'https://doi.org/10.1256/wea.183.02', url: 'https://doi.org/10.1256/wea.183.02' },
  ]),
);

// =================================================================
// APPENDIX — methodological note (no change log per user instruction)
// =================================================================
children.push(
  h1('Appendix. Methodological notes on the open-data substitution'),
  p('This appendix documents three methodological details that support the paper claim that the proprietary Fields in Trust Green Space Index can be replicated by an openly reproducible per-capita measure, with regional ordering preserved.'),

  h2('A.1 Fields in Trust typology exclusions'),
  p('The Fields in Trust Green Space Index excludes the following polygon classes from its calculation of provision per capita: allotments and community growing spaces; cemeteries, burial grounds and crematoria; churchyards and religious grounds; and golf courses. These exclusions are applied because the spaces, although classified as green space in the broader Ordnance Survey typology, are not accessible to the general public for amenity use in the same sense as urban parks, public gardens, and play space. The thesis applies the same exclusions to the OS Open Greenspace polygon set before computing the per-capita measure, preserving comparability with the Fields in Trust definition. Inclusion of the excluded classes would change the absolute per-capita values by approximately 12\u201320 percent depending on region but would not change the regional ordering substantially.'),

  h2('A.2 Centroid-in-polygon spatial assignment'),
  p('OS Open Greenspace polygons rarely cross NUTS-1 / ITL-1 regional boundaries. For the small number of polygons (approximately 0.3 percent of the polygon set) that do cross boundaries, the centroid-in-polygon assignment used in this thesis attributes each polygon to a single region based on the location of its geometric centroid. An area-weighted attribution, splitting the polygon area between the two regions in proportion to the area falling within each region, would change regional totals by under 1 percent, which is well within the noise floor of the four-region Fields in Trust comparison and would not affect the Pearson r = 1.00 result.'),

  h2('A.3 Validation against the published Fields in Trust four-region ranking'),
  p('The Fields in Trust Green Space Index 2020 publication reports values for four regions: London, North East, Wales, and East of England, in that ascending order (London worst-provided, East of England best-provided). The open-data per-capita square-metres measure constructed in this thesis ranks the same four regions in identically the same order: London 22.8 m\u00b2 per capita, North East 34.6, Wales 41.6, East of England 42.3. The Pearson correlation coefficient between the Fields in Trust index values (London 0.54, North East 0.84, Wales 1.03, East of England 1.04) and the computed per-capita values is +1.00 on the four-region comparison, with the same monotonic ordering across the ten-region full panel (Spearman \u03c1 = +1.00 by inspection of the values in Table 3.4). This is the strongest available evidence that the open-data substitute preserves the substantive ordering captured by the proprietary index, and is the basis for the methodological claim made in Section 3.2 and reiterated in the Conclusion.'),
);

// =================================================================
// Build the document
// =================================================================
const doc = new Document({
  features: { updateFields: true },
  creator: 'Gideon Kwesi Tsikudo',
  title: 'Spatiotemporal Analysis of Heatwave-Related Mortality in the UK (England and Wales)',
  description: 'HSE Master\u2019s thesis, Moscow 2024',
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 32, bold: true, font: SERIF, color: '1A1A1A' },
        paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 26, bold: true, font: SERIF, color: '1A1A1A' },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 1 } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, italics: true, font: SERIF, color: '1A1A1A' },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 2 } },
    ],
  },
  numbering: {
    config: [
      { reference: 'bullets',
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
        ] },
    ],
  },
  sections: buildSections(children),
});

// =================================================================
// Section builder: title (no numbering), front matter (lower roman),
// body (arabic, restarting at 1).
// =================================================================
function buildSections(all) {
  const PAGE = {
    size: { width: 12240, height: 15840 },  // US Letter
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },  // 1" margins
  };
  function runningHeader() {
    return new Header({
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: 'Tsikudo \u00b7 UK Heatwave Mortality \u00b7 2024',
                                  font: FONT, size: 18, color: '888888', italics: true })],
      })],
    });
  }
  function pageFooter() {
    return new Footer({
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })],
      })],
    });
  }
  function emptyFooter() {
    return new Footer({ children: [new Paragraph({ children: [] })] });
  }

  const titleChildren = all.slice(0, FRONT_MATTER_START);
  const frontChildren = all.slice(FRONT_MATTER_START, BODY_START);
  const bodyChildren  = all.slice(BODY_START);

  return [
    // Title page: no header, no page number.
    {
      properties: { page: PAGE },
      footers: { default: emptyFooter() },
      children: titleChildren,
    },
    // Front matter (Abstract through List of Acronyms): lower roman, start i.
    {
      properties: {
        page: { ...PAGE, pageNumbers: { start: 1, formatType: NumberFormat.LOWER_ROMAN } },
      },
      headers: { default: runningHeader() },
      footers: { default: pageFooter() },
      children: frontChildren,
    },
    // Body (Chapter 1 onward): arabic, restart at 1.
    {
      properties: {
        page: { ...PAGE, pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL } },
      },
      headers: { default: runningHeader() },
      footers: { default: pageFooter() },
      children: bodyChildren,
    },
  ];
}

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(path.join(__dirname, 'Thesis_Revised.docx'), buf);
  console.log('Wrote Thesis_Revised.docx (' + (buf.length / 1024).toFixed(1) + ' KB)');
});
