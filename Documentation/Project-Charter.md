# SB Schirmer Asset Management System

## Project Charter

**Version:** 1.0  
**Status:** Draft for Approval  
**Owner:** SB Schirmer Co.  
**Product Owner:** Sandy Schirmer  
**Software Engineer:** ChatGPT  
**Current Target:** Capability Maturity Level 2

---

## 1. Project Purpose

The SB Schirmer Asset Management System will create a reliable, searchable, and increasingly automated system for managing SB Schirmer Co. digital business assets.

The immediate focus is food and product images stored in Google Drive.

The system will inventory approved assets, analyze them with AI, and record structured metadata in Google Sheets without modifying the original files.

Over time, the system may support recipes, PDFs, Shopify content, Pinterest content, marketing assets, and other business media.

---

## 2. Business Problem

SB Schirmer Co. has accumulated a large number of images and related content across Google Drive and other systems.

Current challenges include:

- difficulty locating useful images,
- inconsistent folder organization,
- weak links between images and recipes,
- duplicated effort,
- limited visibility into asset quality,
- slow content creation,
- and missed opportunities to reuse existing marketing assets.

These problems increase time, cost, and frustration while limiting the company’s ability to use its existing content effectively.

---

## 3. Project Vision

Create a trusted digital asset system that enables SB Schirmer Co. to quickly understand:

- what assets exist,
- where they are stored,
- what they contain,
- which recipes they support,
- how useful they are,
- and where they can be used in marketing.

The long-term vision is a business capability that supports faster, safer, and more consistent content creation across Shopify, Pinterest, email, social media, and recipe publishing.

---

## 4. Immediate Objective

The immediate objective is to achieve a solid Capability Maturity Level 2.

At Level 2, the system and its development process must be:

- planned,
- documented,
- repeatable,
- version controlled,
- recoverable,
- secure,
- testable,
- and suitable for controlled use.

The current goal is not maximum functionality.

The current goal is dependable capability.

---

## 5. Phase 1 Scope

Phase 1 includes:

- scanning one approved Google Drive root folder,
- scanning its subfolders,
- identifying supported image assets,
- recording one spreadsheet row per asset,
- preventing duplicate inventory records using Drive File ID,
- marking new images as Pending Analysis,
- analyzing approved images through the OpenAI API,
- recording structured metadata,
- processing images in controlled batches,
- resuming work through Apps Script triggers,
- and preserving all original assets unchanged.

---

## 6. Phase 1 Out of Scope

Phase 1 does not include:

- renaming files,
- moving files,
- deleting files,
- duplicating files,
- changing ownership,
- changing permissions,
- automatically reorganizing Google Drive,
- publishing to Shopify,
- publishing to Pinterest,
- modifying original images,
- or making irreversible asset-management decisions.

These capabilities require separate approval and later project phases.

---

## 7. Primary Stakeholders

### Sandy Schirmer — Product Owner

Responsibilities:

- define business priorities,
- approve scope,
- approve risk,
- evaluate business value,
- and make final product decisions.

### ChatGPT — Software Engineer

Responsibilities:

- architecture,
- implementation,
- security review,
- testing strategy,
- documentation,
- maintainability,
- and technical recommendations.

### Tabatha — AI Coach and Strategic Advisor

Responsibilities may include:

- evaluating business alignment,
- challenging priorities,
- supporting accountability,
- and helping assess whether the work is producing meaningful business value.

---

## 8. Guiding Principles

The project will follow these principles:

1. Protect original assets.
2. Build capabilities, not isolated features.
3. Business value drives engineering.
4. Small, reversible changes are preferred.
5. Secrets never belong in source code.
6. Every material change is version controlled.
7. AI output must be reviewed when uncertain.
8. Cost controls must exist before large-scale processing.
9. Documentation is part of the product.
10. Capability maturity is more important than speed alone.

---

## 9. Success Measures

Phase 1 will be considered successful when:

- every approved image can be inventoried,
- duplicate rows are prevented,
- original Drive files remain unchanged,
- AI analysis results are written into defined columns,
- interrupted processing can recover,
- failures are visible and limited,
- API usage is controlled,
- source code is version controlled,
- documentation is current,
- and the system can be operated through a repeatable procedure.

---

## 10. Level 2 Exit Criteria

Capability Maturity Level 2 will be achieved when:

- Engineering Standards are approved,
- this Project Charter is approved,
- current architecture is documented,
- security risks are documented,
- a decision log exists,
- a risk register exists,
- source code is stored in GitHub,
- the development environment is operational,
- configuration is centralized,
- retry and recovery behavior is defined,
- testing is repeatable,
- deployment is documented,
- and one controlled end-to-end test succeeds.

---

## 11. Key Risks

Current risks include:

- API key exposure,
- excessive API usage,
- duplicate processing,
- interrupted Apps Script executions,
- rows becoming stuck in incomplete states,
- poor-quality AI output,
- accidental processing of sensitive images,
- excessive Google permissions,
- undocumented changes,
- and loss of project knowledge.

These risks must be reduced before large-scale processing begins.

---

## 12. Technical Constraints

The system currently depends on:

- Google Drive,
- Google Sheets,
- Google Apps Script,
- OpenAI API,
- Git,
- GitHub,
- Node.js,
- and Google clasp.

Important constraints include:

- Apps Script execution limits,
- Google service quotas,
- API request limits,
- image file size,
- external API cost,
- and the need to preserve original assets.

---

## 13. Deliverables

Current required deliverables include:

- working image inventory engine,
- working image analysis engine,
- centralized configuration,
- engineering standards,
- project charter,
- architecture documentation,
- security review,
- risk register,
- decision log,
- test procedure,
- deployment procedure,
- and Level 2 maturity assessment.

---

## 14. Future Phases

Potential later phases include:

### Phase 2 — Improved Metadata and Search

- richer image classification,
- quality scoring,
- hero-image identification,
- marketing suitability,
- review workflow,
- and search.

### Phase 3 — Asset Relationships

- image-to-recipe matching,
- PDF matching,
- Shopify linking,
- Pinterest linking,
- and duplicate detection.

### Phase 4 — Business Asset Platform

- dashboards,
- reporting,
- content planning,
- marketing workflows,
- automated recommendations,
- and additional media types.

---

## 15. Approval Standard

This charter is approved when the Product Owner agrees that it accurately defines:

- why the project exists,
- what Phase 1 includes,
- what Phase 1 excludes,
- how success will be measured,
- and what Level 2 maturity requires.

---

## 16. Closing Principle

> Build slowly enough to remain safe, clearly enough to remain understandable, and deliberately enough to create lasting business capability.