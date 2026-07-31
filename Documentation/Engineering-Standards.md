\# SB Schirmer Asset Management System



\## Engineering Standards



\*\*Version:\*\* 1.0  

\*\*Status:\*\* Approved  

\*\*Governing principle:\*\* Slow is smooth. Smooth is fast.



\---



\## 1. Purpose



These standards govern the design, development, testing, deployment, and maintenance of the SB Schirmer Asset Management System.



The system exists to safely inventory, classify, relate, and manage SB Schirmer Co. digital business assets while reducing manual work and supporting future marketing, recipe, Shopify, Pinterest, and content-management capabilities.



The standards are intended to prevent:



\- data loss,

\- uncontrolled changes,

\- security exposure,

\- unnecessary cost,

\- duplicated work,

\- undocumented decisions,

\- and dependence on individual memory.



\---



\## 2. Roles



\### Product Owner



The Product Owner defines:



\- business goals,

\- priorities,

\- acceptable risk,

\- required outcomes,

\- and final approval.



\### Software Engineer



The Software Engineer is responsible for:



\- architecture,

\- implementation,

\- security,

\- testing,

\- documentation,

\- code review,

\- maintainability,

\- and technical recommendations.



The Software Engineer must raise concerns when a requested change would create unnecessary risk, cost, complexity, or technical debt.



\---



\## 3. Capability-Maturity Objective



The project will not operate at Capability Maturity Level 0 or 1.



The immediate objective is a practical CM 2 foundation in which:



\- work is planned,

\- requirements are recorded,

\- source code is version controlled,

\- changes are reversible,

\- risks are documented,

\- testing is repeatable,

\- and releases are traceable.



The longer-term objective is CM 3, where engineering and operating processes are documented, standardized, and consistently followed.



\---



\## 4. Protection of Original Assets



During Phase 1, the system is inventory-only.



The system may:



\- read files,

\- inspect file metadata,

\- analyze approved images,

\- and write inventory information to the designated spreadsheet.



The system may not:



\- rename original files,

\- move original files,

\- delete original files,

\- replace original files,

\- duplicate original files,

\- alter file ownership,

\- or change sharing permissions.



Any future capability that changes original assets requires:



1\. a separate approved project phase,

2\. documented requirements,

3\. a review queue,

4\. explicit user approval,

5\. and a tested rollback procedure.



\---



\## 5. Configuration Management



Shared configuration belongs in:



```text

AppsScript/Config.js

