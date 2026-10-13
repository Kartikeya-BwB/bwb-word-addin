// Example data for each chart type, to start from in the builder and to draw in tests. Figures are illustrative.

import type { ChartType } from './catalogue.ts'
import type { NumberFormat } from './format.ts'
import type { Grid } from './grid.ts'

export interface Sample {
  title: string
  source: string
  data: Grid
  format?: NumberFormat
  highlight?: string
}

export const SAMPLES: Record<ChartType, Sample> = {
  column: {
    title: 'Funding by year',
    source: 'BwB analysis (2026)',
    data: [['Year', 'Funding'], ['2022', 'USD 120m'], ['2023', 'USD 150m'], ['2024', 'USD 185m'], ['2025', 'USD 240m']],
  },
  stackedColumn: {
    title: 'Climate finance mobilised by instrument, USD bn',
    source: 'OECD (2026)',
    data: [
      ['Year', 'Grants', 'Loans', 'Equity', 'Guarantees'],
      ['2021', '18.2', '52.1', '4.3', '2.1'],
      ['2022', '21.4', '58.9', '5.1', '2.8'],
      ['2023', '24.0', '66.3', '6.2', '3.5'],
      ['2024', '25.7', '71.8', '7.9', '4.4'],
      ['2025', '27.3', '78.2', '9.1', '5.6'],
    ],
  },
  column100: {
    title: 'Climate finance by theme and region, 2025',
    source: 'OECD (2026)',
    data: [
      ['Region', 'Mitigation', 'Adaptation', 'Cross-cutting'],
      ['Sub-Saharan Africa', '42', '41', '17'],
      ['South Asia', '58', '29', '13'],
      ['Latin America', '61', '24', '15'],
      ['East Asia and Pacific', '67', '21', '12'],
    ],
  },
  bar: {
    title: 'Green bond issuance by country, 2025',
    source: 'Climate Bonds Initiative (2026)',
    data: [
      ['Country', 'Issuance'],
      ['Germany', 'USD 72.3bn'],
      ['France', 'USD 58.9bn'],
      ['China', 'USD 55.2bn'],
      ['United States', 'USD 44.1bn'],
      ['Netherlands', 'USD 21.7bn'],
      ['United Kingdom', 'USD 18.4bn'],
    ],
    highlight: 'United Kingdom',
  },
  stackedBar: {
    title: 'Portfolio by sector and instrument, USD m',
    source: 'BwB portfolio data (2026)',
    data: [
      ['Sector', 'Debt', 'Equity', 'Guarantees'],
      ['Renewable energy', '310', '95', '40'],
      ['Water and sanitation', '180', '20', '55'],
      ['Sustainable agriculture', '120', '45', '30'],
      ['Clean transport', '90', '60', '10'],
      ['Health', '45', '15', '25'],
    ],
  },
  bar100: {
    title: 'Survey: barriers to private investment',
    source: 'BwB investor survey, 120 responses (2026)',
    data: [
      ['Statement', 'Agree', 'Neutral', 'Disagree'],
      ['Currency risk deters investment', '68%', '19%', '13%'],
      ['Project pipelines are too thin', '54%', '28%', '18%'],
      ['Ticket sizes are too small', '47%', '31%', '22%'],
      ['Policy is unpredictable', '61%', '22%', '17%'],
    ],
  },
  line: {
    title: 'Sustainable bond issuance, USD bn',
    source: 'Climate Bonds Initiative (2026)',
    data: [
      ['Year', 'Green bonds', 'Social bonds', 'Sustainability bonds'],
      ['2019', '271', '20', '66'],
      ['2020', '297', '149', '157'],
      ['2021', '522', '223', '199'],
      ['2022', '487', '136', '157'],
      ['2023', '588', '161', '159'],
      ['2024', '672', '154', '172'],
      ['2025', '701', '162', '186'],
    ],
  },
  waterfall: {
    title: 'Programme budget, 2025 to 2026, USD m',
    source: 'BwB programme accounts (2026)',
    data: [
      ['Step', 'Value', 'Kind'],
      ['Budget 2025', '100', 'total'],
      ['Staff', '12', ''],
      ['Travel', '-5', ''],
      ['Consultants', '8', ''],
      ['Before savings', '', 'subtotal'],
      ['Efficiencies', '-6', ''],
      ['Budget 2026', '', 'total'],
    ],
  },
  mekko: {
    title: 'Climate finance need by region and sector, USD bn',
    source: 'BwB analysis of CPI data (2026)',
    data: [
      ['Region', 'Energy', 'Water', 'Agriculture'],
      ['Sub-Saharan Africa', '120', '45', '60'],
      ['South Asia', '160', '40', '45'],
      ['Latin America', '90', '30', '35'],
      ['Middle East and North Africa', '50', '35', '15'],
    ],
  },
  gantt: {
    title: 'Workplan',
    source: '',
    data: [
      ['Task', 'Start', 'End'],
      ['Inception', '', ''],
      ['Kick-off and document review', '2 November 2026', '27 November 2026'],
      ['Inception report', '30 November 2026', ''],
      ['Diagnostic', '', ''],
      ['Stakeholder interviews', '30 November 2026', '29 January 2027'],
      ['Market sizing', 'January 2027', 'February 2027'],
      ['Interim report', '26 February 2027', ''],
      ['Design', '', ''],
      ['Facility design options', 'March 2027', 'April 2027'],
      ['Validation workshop', '7 April 2027', ''],
      ['Final report', '30 April 2027', ''],
    ],
  },
  donut: {
    title: 'Portfolio by sector',
    source: 'BwB portfolio data (2026)',
    data: [['Sector', 'Commitments'], ['Energy', '38'], ['Water', '22'], ['Agriculture', '18'], ['Transport', '12'], ['Other', '10']],
    format: { currency: 'USD', scale: 'm' },
  },
  harvey: {
    title: 'Assessment of facility options',
    source: 'BwB analysis (2026)',
    data: [
      ['Option', 'Impact', 'Cost', 'Speed', 'Risk', 'Scalability'],
      ['Blended finance fund', '4', '2', '1', '3', '4'],
      ['Guarantee facility', '3', '3', '2', '2', '3'],
      ['Direct lending', '2', '1', '4', '1', '2'],
      ['Technical assistance only', '1', '4', '3', '4', '1'],
    ],
    highlight: 'Blended finance fund',
  },
}
