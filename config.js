// ═══════════════════════════════════════════════════════════════════
//  config.js – Beghou 2026 Firm Building Initiatives Dashboard
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  clientId:   'YOUR_AZURE_AD_CLIENT_ID',
  tenantId:   'YOUR_AZURE_AD_TENANT_ID',
  redirectUri: 'https://conroy-patrick.github.io/initiatives-2026/',
  graphEndpoint: 'https://graph.microsoft.com/v1.0',
  siteHostname:  'beghouconsultinginc-my.sharepoint.com',
  sitePath:      '/personal/aaron_mccracken_beghouconsulting_com',
  initiativeListName: '2026 Initiative Tracker',
  updatesListName:    'Initiative Updates',
  fiscalYearStart: new Date('2026-01-01'),
  fiscalYearEnd:   new Date('2026-12-31'),
  refreshIntervalMs: 5 * 60 * 1000
};

CONFIG.expectedProgress = (() => {
  const now = new Date(), start = CONFIG.fiscalYearStart, end = CONFIG.fiscalYearEnd;
  return Math.round((Math.max(0, Math.min(now - start, end - start)) / (end - start)) * 100);
})();

const SEED_INITIATIVES = [
  {
    id: 1, title: 'Bookings Sufficient to Achieve Target',
    category: 'Revenue & Operations',
    description: 'Drive new logo acquisition, pipeline discipline, and forward-looking metrics to achieve ACV booking targets. KPIs: new logos, pipeline management, # of RFPs received, ACV Book to Bill.',
    accountablePartner: 'Dennis', builders: 'Caden', execSponsor: 'MR',
    status: 'On Track', percentComplete: 30,
    q1Goals: 'Define bookings targets and ACV methodology; baseline current pipeline health',
    q2Goals: 'Pipeline management framework live; RFP tracking operational; new logo target set',
    q3Goals: 'Full KPI reporting live; mid-year review completed; H2 pipeline targets confirmed',
    latestComment: 'ACV methodology finalized and shared with leadership. Pipeline review cadence now bi-weekly. Three qualified new logo opportunities entered pipeline in Q1.',
    lastUpdatedBy: 'Caden R.', modified: '2026-04-08'
  },
  {
    id: 2, title: 'Commercial Infrastructure & Sustainable Growth',
    category: 'Revenue & Operations',
    description: 'Build sales infrastructure enabling consistent, scalable growth: sales ops, RFP excellence, solution architect requirements, sales training, large account management, account planning, and pricing discipline.',
    accountablePartner: 'Clare', builders: 'Mudit, Adriana, Luke', execSponsor: 'MR',
    status: 'At Risk', percentComplete: 18,
    q1Goals: 'Sales ops audit; RFP process assessment completed; solution architect role defined',
    q2Goals: 'Solution architect requirements defined; training program launched; pricing playbook v1',
    q3Goals: 'Full commercial infrastructure operational; pricing discipline reinforced; account plans live',
    latestComment: 'RFP assessment complete but solution architect role definition stalled pending headcount approval. Flagging At Risk — need exec decision by end of April to hit Q2 milestone.',
    lastUpdatedBy: 'Clare M.', modified: '2026-04-10'
  },
  {
    id: 3, title: 'Effective Workforce Planning & Deployment',
    category: 'Revenue & Operations',
    description: 'Redesign talent deployment across accounts, solutions, and geographies with clear standards and accountability. Includes pod model, staffing governance, performance management, and US-India collaboration model.',
    accountablePartner: 'Kevin', builders: 'Luke, Aaron, Dan K', execSponsor: 'MR',
    status: 'On Track', percentComplete: 45,
    q1Goals: 'Concept rollout and leadership alignment discussions completed',
    q2Goals: 'New staffing model finalized; all workstreams completed by 5/31',
    q3Goals: 'Pod model fully implemented; performance frameworks live; US-India model piloted',
    latestComment: 'Pod Model Design ~75% complete (V2 due 4/15). Staffing Governance ~50% complete. Manager meeting April 27th, All Hands May 5th.',
    lastUpdatedBy: 'Kevin F.', modified: '2026-04-07'
  },
  {
    id: 4, title: 'Deliverable Quality',
    category: 'Revenue & Operations',
    description: 'Establish firm-wide standards for deliverable quality through knowledge management systems, QC processes, training programs, solutions templates, and project documentation practices.',
    accountablePartner: 'TBD', builders: 'TBD', execSponsor: 'MR',
    status: 'Behind', percentComplete: 5,
    q1Goals: 'Identify accountable partner and core team; scope workstreams',
    q2Goals: 'QC framework defined; knowledge management platform selected',
    q3Goals: 'Templates, training materials, and QC processes launched firm-wide',
    latestComment: 'Accountable partner not yet assigned — Q1 goal not completed. Recommend escalating to MR to assign ownership by end of April.',
    lastUpdatedBy: 'MR', modified: '2026-04-09'
  },
  {
    id: 5, title: 'AI Strategy for Internal Use (L1\u2013L3)',
    category: 'Revenue & Operations',
    description: 'Define and execute Beghou\'s AI strategy for internal workflow automation. Clarify objectives and measurement; formalize risk guardrails; establish knowledge sharing and standardized AI toolkits across L1\u2013L3 workflows.',
    accountablePartner: 'Nicole', builders: 'Sharath', execSponsor: 'MR',
    status: 'On Track', percentComplete: 35,
    q1Goals: 'Clarify objectives and draft risk guardrail framework; inventory current AI tool usage',
    q2Goals: 'Knowledge sharing program launched; toolkit standards published; L1 workflows identified',
    q3Goals: 'Standardized AI toolkits deployed firm-wide; L2\u2013L3 workflows automated',
    latestComment: 'Risk guardrail framework v1 approved. Firm-wide AI survey: 67% of staff using at least one AI tool. Toolkit standardization underway.',
    lastUpdatedBy: 'Nicole B.', modified: '2026-04-06'
  },
  {
    id: 6, title: 'Future-Proofing Data Management Offering',
    category: 'Product & Tech',
    description: 'Define and execute a 2026 roadmap for Beghou\'s data management offering that ensures it remains competitive and supports in-year delivery commitments for clients.',
    accountablePartner: 'Piotr', builders: 'TBD', execSponsor: 'DC',
    status: 'On Track', percentComplete: 20,
    q1Goals: 'Roadmap definition and stakeholder alignment; competitive landscape review',
    q2Goals: 'Priority roadmap items in active development; client delivery commitments mapped',
    q3Goals: 'Key roadmap milestones delivered; competitive offering reviewed with clients',
    latestComment: 'Competitive landscape review complete. Three priorities: Databricks migration support, MDM modernization, AI-assisted data validation. Stakeholder alignment sessions scheduled for late April.',
    lastUpdatedBy: 'Piotr W.', modified: '2026-04-05'
  },
  {
    id: 7, title: 'Technology Strategy for Scale & Profit',
    category: 'Product & Tech',
    description: 'Prioritize product and reusable technology development efforts. Identify and evaluate roadmap candidates with ROI analysis to maximize firm profitability through scalable technology leverage.',
    accountablePartner: 'Jason', builders: 'TBD', execSponsor: 'DC',
    status: 'On Track', percentComplete: 25,
    q1Goals: 'Technology asset inventory; reuse opportunity assessment; ROI framework defined',
    q2Goals: 'Prioritized roadmap with ROI modeling published; top 3 candidates in scoping',
    q3Goals: 'Top-priority reusable technologies in development or deployed; profit impact modeled',
    latestComment: 'Technology asset inventory complete: 14 reusable components identified across Arc, Mainsail, and IC. ROI framework under review with DC. On track for Q2 roadmap publish.',
    lastUpdatedBy: 'Jason K.', modified: '2026-04-04'
  },
  {
    id: 8, title: 'Optimizing Arc Data Management Delivery',
    category: 'Product & Tech',
    description: 'Improve Arc data management delivery efficiency through structured training, strategic hiring, resource planning, and launch of a dedicated delivery center.',
    accountablePartner: 'Justin', builders: 'TBD', execSponsor: 'DC',
    status: 'On Track', percentComplete: 40,
    q1Goals: 'Training curriculum defined; hiring plan finalized; delivery center location scoped',
    q2Goals: 'Delivery center pilot launched; resource model operational; first cohort trained',
    q3Goals: 'Full delivery center operational; efficiency metrics on target; Q3 capacity plan live',
    latestComment: 'Training curriculum piloted with 8 associates — strong feedback. Hiring plan approved: 4 new DM roles in H1. Delivery center pilot kicks off May 1st.',
    lastUpdatedBy: 'Justin L.', modified: '2026-04-07'
  },
  {
    id: 9, title: 'Beghou Unified Systems Success',
    category: 'Systems & People',
    description: 'Drive successful implementation of Beghou\'s unified systems platform: dynamic capacity planning, automated revenue recognition, resource management, time capture, and finance-led invoicing.',
    accountablePartner: 'Yair (with Michael)', builders: 'Luke, Patrick, Vendors', execSponsor: 'MG',
    status: 'At Risk', percentComplete: 22,
    q1Goals: 'System requirements finalized; vendor selection complete; integration architecture approved',
    q2Goals: 'Integration development and testing underway; data migration plan confirmed',
    q3Goals: 'Go-live and adoption monitoring active; finance sign-off complete',
    latestComment: 'Vendor selected (Ruddr + Sage Intacct). Requirements 90% done but time capture module has uncovered complexity — estimating 3-week delay. At Risk due to timeline pressure.',
    lastUpdatedBy: 'Yair S.', modified: '2026-04-09'
  },
  {
    id: 10, title: 'Develop Our Talent Infrastructure',
    category: 'Systems & People',
    description: 'Build Beghou\'s talent infrastructure through a comprehensive competency model, structured skill development programs, and modernized performance management practices.',
    accountablePartner: 'Dan S (with Ali)', builders: 'HR team', execSponsor: 'AF',
    status: 'On Track', percentComplete: 28,
    q1Goals: 'Competency model framework drafted; HRBP model defined; current state assessment complete',
    q2Goals: 'Performance management process defined; skill development pilot launched',
    q3Goals: 'Full talent infrastructure launched; training programs active; competency model rolled out',
    latestComment: 'Competency model complete for Consultant through Manager levels. HRBP model approved for H2 implementation. Skill gap assessment survey launching next week.',
    lastUpdatedBy: 'Dan S.', modified: '2026-04-08'
  }
];

const SEED_UPDATES = [
  { initiativeTitle: 'Effective Workforce Planning & Deployment', updatedBy: 'Kevin F.', status: 'On Track', percentComplete: 45, comment: 'Pod Model Design ~75% complete (V2 due 4/15). Staffing Governance ~50%. Manager meeting April 27th, All Hands May 5th.', created: '2026-04-07T14:22:00Z' },
  { initiativeTitle: 'Effective Workforce Planning & Deployment', updatedBy: 'Aaron M.', status: 'On Track', percentComplete: 35, comment: 'Staffing governance workstream kicked off. Reviewed pod model V1 with Kevin — good alignment. US-India collaboration model draft in progress.', created: '2026-03-24T10:05:00Z' },
  { initiativeTitle: 'Effective Workforce Planning & Deployment', updatedBy: 'Kevin F.', status: 'On Track', percentComplete: 20, comment: 'Leadership alignment sessions completed. All partners briefed on pod model concept — positive reception. Workstream teams assembled.', created: '2026-02-28T09:15:00Z' },
  { initiativeTitle: 'Effective Workforce Planning & Deployment', updatedBy: 'Kevin F.', status: 'On Track', percentComplete: 10, comment: 'Initiative kicked off. Concept deck presented to leadership. Beginning workstream planning across pod model, staffing governance, performance mgmt, and US-India collaboration.', created: '2026-01-20T11:30:00Z' },

  { initiativeTitle: 'Bookings Sufficient to Achieve Target', updatedBy: 'Caden R.', status: 'On Track', percentComplete: 30, comment: 'ACV methodology finalized and shared with leadership. Pipeline review cadence now bi-weekly. Three qualified new logo opportunities in pipeline.', created: '2026-04-08T16:00:00Z' },
  { initiativeTitle: 'Bookings Sufficient to Achieve Target', updatedBy: 'Caden R.', status: 'On Track', percentComplete: 15, comment: 'ACV methodology draft under review. Pipeline health baseline established — coverage ratio at 2.1x target. RFP tracking template shared with BD team.', created: '2026-03-10T13:45:00Z' },
  { initiativeTitle: 'Bookings Sufficient to Achieve Target', updatedBy: 'Dennis P.', status: 'On Track', percentComplete: 5, comment: 'Initiative launched. First milestone: define ACV methodology and bookings targets with finance by end of February.', created: '2026-01-15T09:00:00Z' },

  { initiativeTitle: 'Commercial Infrastructure & Sustainable Growth', updatedBy: 'Clare M.', status: 'At Risk', percentComplete: 18, comment: 'RFP assessment complete. Solution architect role stalled pending headcount approval. Flagging At Risk — need exec decision by end of April.', created: '2026-04-10T08:30:00Z' },
  { initiativeTitle: 'Commercial Infrastructure & Sustainable Growth', updatedBy: 'Mudit S.', status: 'On Track', percentComplete: 15, comment: 'Sales ops audit complete. Key findings: inconsistent account planning, no standardized pricing playbook. RFP assessment underway.', created: '2026-03-28T15:20:00Z' },
  { initiativeTitle: 'Commercial Infrastructure & Sustainable Growth', updatedBy: 'Clare M.', status: 'On Track', percentComplete: 8, comment: 'Workstream teams assembled. Sales ops audit kicked off. RFP process interviews scheduled with 6 partners.', created: '2026-02-12T10:00:00Z' },

  { initiativeTitle: 'Deliverable Quality', updatedBy: 'MR', status: 'Behind', percentComplete: 5, comment: 'Accountable partner not yet assigned. Behind on Q1 goal. Recommend escalating to assign ownership by end of April.', created: '2026-04-09T11:00:00Z' },
  { initiativeTitle: 'Deliverable Quality', updatedBy: 'MR', status: 'Not Started', percentComplete: 0, comment: 'Initiative defined but partner assignment pending. Will align with leadership in March to identify the right owner.', created: '2026-02-01T09:00:00Z' },

  { initiativeTitle: 'AI Strategy for Internal Use (L1\u2013L3)', updatedBy: 'Nicole B.', status: 'On Track', percentComplete: 35, comment: 'Risk guardrail framework v1 approved. Firm-wide survey: 67% of staff using at least one AI tool. Toolkit standardization underway with Sharath.', created: '2026-04-06T14:00:00Z' },
  { initiativeTitle: 'AI Strategy for Internal Use (L1\u2013L3)', updatedBy: 'Sharath N.', status: 'On Track', percentComplete: 20, comment: 'AI tool inventory complete across practice areas — 11 tools in use, consolidating to 4 approved platforms. Risk guardrail draft circulating.', created: '2026-03-14T10:30:00Z' },
  { initiativeTitle: 'AI Strategy for Internal Use (L1\u2013L3)', updatedBy: 'Nicole B.', status: 'On Track', percentComplete: 10, comment: 'Objectives and scope confirmed with MR. Beginning AI tool inventory. Sharath engaged for toolkit standardization workstream.', created: '2026-01-28T11:00:00Z' },

  { initiativeTitle: 'Future-Proofing Data Management Offering', updatedBy: 'Piotr W.', status: 'On Track', percentComplete: 20, comment: 'Competitive landscape review complete. Three roadmap priorities: Databricks migration, MDM modernization, AI-assisted validation. Stakeholder alignment in late April.', created: '2026-04-05T09:45:00Z' },
  { initiativeTitle: 'Future-Proofing Data Management Offering', updatedBy: 'Piotr W.', status: 'On Track', percentComplete: 10, comment: 'Roadmap scoping kicked off. Competitor analysis underway. Internal client interviews scheduled with 4 accounts.', created: '2026-02-20T14:00:00Z' },

  { initiativeTitle: 'Technology Strategy for Scale & Profit', updatedBy: 'Jason K.', status: 'On Track', percentComplete: 25, comment: 'Technology asset inventory complete: 14 reusable components identified. ROI framework under review with DC. On track for Q2 roadmap.', created: '2026-04-04T16:15:00Z' },
  { initiativeTitle: 'Technology Strategy for Scale & Profit', updatedBy: 'Jason K.', status: 'On Track', percentComplete: 12, comment: 'Technology asset inventory in progress. ROI framework template drafted. 20+ reusable tech candidates being evaluated.', created: '2026-03-05T11:00:00Z' },

  { initiativeTitle: 'Optimizing Arc Data Management Delivery', updatedBy: 'Justin L.', status: 'On Track', percentComplete: 40, comment: 'Training curriculum piloted with 8 associates — strong feedback. Hiring plan approved: 4 new DM roles. Delivery center pilot starts May 1st.', created: '2026-04-07T13:30:00Z' },
  { initiativeTitle: 'Optimizing Arc Data Management Delivery', updatedBy: 'Justin L.', status: 'On Track', percentComplete: 25, comment: 'Training curriculum modules 1-3 drafted. Pilot group identified. Hiring plan under DC review — expecting approval by end of March.', created: '2026-03-18T10:00:00Z' },
  { initiativeTitle: 'Optimizing Arc Data Management Delivery', updatedBy: 'Justin L.', status: 'On Track', percentComplete: 10, comment: 'Initiative scoped. Delivery center location options being evaluated. Training curriculum outline drafted.', created: '2026-02-05T09:00:00Z' },

  { initiativeTitle: 'Beghou Unified Systems Success', updatedBy: 'Yair S.', status: 'At Risk', percentComplete: 22, comment: 'Vendor selected (Ruddr + Sage Intacct). Requirements 90% done but time capture complexity adding ~3 weeks. Flagging At Risk.', created: '2026-04-09T15:00:00Z' },
  { initiativeTitle: 'Beghou Unified Systems Success', updatedBy: 'Patrick C.', status: 'On Track', percentComplete: 18, comment: 'Integration architecture approved. HubSpot to Sage Intacct field mapping complete. Ruddr API scoping underway.', created: '2026-03-22T14:00:00Z' },
  { initiativeTitle: 'Beghou Unified Systems Success', updatedBy: 'Yair S.', status: 'On Track', percentComplete: 10, comment: 'Vendor shortlist down to two. Requirements workshops with finance and ops complete. Integration architecture draft under review.', created: '2026-02-25T10:30:00Z' },
  { initiativeTitle: 'Beghou Unified Systems Success', updatedBy: 'Michael G.', status: 'Not Started', percentComplete: 0, comment: 'Initiative launched. Yair leading day-to-day. Vendor selection process kicked off — targeting decision by end of February.', created: '2026-01-12T09:00:00Z' },

  { initiativeTitle: 'Develop Our Talent Infrastructure', updatedBy: 'Dan S.', status: 'On Track', percentComplete: 28, comment: 'Competency model complete through Manager level. HRBP model approved for H2. Skill gap assessment survey launching next week.', created: '2026-04-08T12:00:00Z' },
  { initiativeTitle: 'Develop Our Talent Infrastructure', updatedBy: 'Ali F.', status: 'On Track', percentComplete: 15, comment: 'Competency model draft reviewed with HR. Performance management benchmarking complete — reviewing Lattice and Leapsome. HRBP model in design.', created: '2026-03-15T11:00:00Z' },
  { initiativeTitle: 'Develop Our Talent Infrastructure', updatedBy: 'Dan S.', status: 'On Track', percentComplete: 8, comment: 'Initiative scoped with Ali. Current state talent assessment underway. Competency model framework kick-off scheduled for March.', created: '2026-02-10T10:00:00Z' }
];
