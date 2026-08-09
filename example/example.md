# FlexUp Founders Agreement – Special Conditions (Founders-SC)

::: Published
Based on a template first drafted by Claude on 18^th^ April 2026
:::

::: Comments
{? Guidelines on using this template: This template uses the following types of text: – **Standard text** (black) – Fixed legal wording that is part of the Agreement. Do not modify unless explicitly permitted. – **Data fields** (blue) – Parameters filled automatically by the FlexUp app ({ { ... } }) or manually (‹ ... ›). When editing manually, replace any placeholder in angle brackets ‹ ... › with the appropriate value. – **Options** (green) – Conditional sections processed by the FlexUp app ({ % ... % }) or manually ([ ... ]). When editing manually, read the condition and keep or delete the section as appropriate. – **Guidance** (grey) {? NOTE: ... ?} – Drafting instructions and illustrations for the drafter. Not part of the Agreement. Delete all guidance text before finalising the document. If you wish to modify any standard text, you must: – Explicitly state that changes have been made in the introduction. – Mark any added text in underlined italics. – Mark any deleted text with strikethrough. Failure to declare changes to the template or standard text as indicated above constitutes a violation of the FlexUp Licence and a breach of this Agreement. – After completing your edits, remove all guidance, option blocks, and brackets to ensure the document is clean and easy to read. ?}

{? NOTE: This Agreement may be signed by two or more Founders (multi-founder case) or by a single Founder (solo case). In the solo case, the Agreement constitutes a unilateral commitment by the Founder. Before incorporation into an Incubation Agreement or Grouping Agreement, it binds only the signing Founder or Founders and creates no rights or obligations for any third party. Once incorporated, third-party enforcement rights apply only as provided in the Founders-GC and the incorporating agreement. Adjust the opening sentence and signatures accordingly. ?}
:::

This FlexUp Founders Agreement ("**Agreement**") is entered into by and between the Founders identified in [Appendix 1](#appendix-1) below, collectively referred to as the "**Founders**" and individually as a "**Founder**".

This template is provided by FlexUp under the terms of the FlexUp Licence, which can be found on the FlexUp website ([www.flexup.org](http://www.flexup.org)).

### Article 1. Composition and Interpretation

1.1. The Agreement is composed of the following documents, in descending order of priority:

  a) **Founders-SC**, the present document;
  b) **Founders-GC**;
  c) **NDA-GC**; and
  d) **FlexUp-GC**.

### Article 2. Scope and Obligations

2.1. This Founders-SC defines the specific terms under which the Founders commit to pursue the Project together, and governs their inter-se relationship in accordance with the Founders-GC.

### Article 3. Definitions and specific terms and conditions

Capitalised terms within this Agreement are defined terms, whose definitions are given in the table below or, if not here, elsewhere in the documents composing the Agreement or in the FlexUp Glossary (available on [www.flexup.org/glossary](http://www.flexup.org/glossary)).

::: Definitions

**Defined Term / Key Item**
: **Definition / Specific terms and conditions**

**Contract Label**
: {{ contract.text_label or '‹ Project name › – Founders Agreement • ‹ Date ›' }}

**Founders**
: The persons listed in [Appendix 1](#appendix-1) – Register of Founders.

**Founders' Representative**

: ::: Comments
  {? NOTE: Designate one Founder as Founders' Representative. This person collects votes, issues notices, and co-signs the Charter-SC as attestingwitness. They may also be granted a power of attorney in Appendix 3. ?}
  :::

  ::: Comments
  {{ founders.representative.name or '‹ Full name of Founders' Representative ›' }}
  :::

**Project**

: ::: Comments
  {? NOTE: Describe the common initiative the Founders are pursuing.This description does not need to be exhaustive – it will be detailedfurther in the Charter-SC when the Holding Structure is implemented. ?}
  :::

  ::: Comments
  ‹ Description of the project: sector, activity, and any other relevant details. ›
  :::

**Intended Structure**

: ::: Comments
  {? NOTE: Select one of the options below. ?}
  :::

  [ Grouping ]

  The Founders intend to implement a FlexUp Grouping Agreement. The Grouping is expected to be set up under: {{ grouping.target_name or '‹ Target Grouping name ›' }}. Each Founder must be listed as a Constituent in the Grouping-SC, and each Constituent must be listed as a Founder in this Founders-SC. The Grouping requires a Charter-SC and use of the FlexUp App, even if the Founders do not use flexible remuneration or flexible equity.

  [ Subaccount ]

  The Founders intend for the Project to be held as a Subaccount under the Charter and in the FlexUp App. The Holder of the Subaccount will be:

  {{ incubator.name or '‹ Name of Holder / Incubator, either one of the Founders or an external Incubator ›' }}, acting as Incubator under an Incubation Agreement.

  The Subaccount requires a Charter-SC and use of the FlexUp App, even if the Founders do not use flexible remuneration or flexible equity.

  [ Postponed ]

  The decision between a Grouping and a Subaccount is deferred. It shall be taken by the Founders' Assembly by a Simple Majority when the Founders are ready to proceed, in accordance with the Founders-GC.

**Remuneration and Equity Framework**

: ::: Comments
  {? NOTE: Choose the remuneration and equity framework. A Grouping orSubaccount still requires a Charter-SC and use of the FlexUp App, evenwhere the Founders do not use flexible remuneration or flexible equity. ?}
  :::

  [ Flexible remuneration and equity ]

  The Project will use flexible remuneration and equity under the Charter.

  [ Classic remuneration and equity ]

  The Project will use the Charter and the FlexUp App where required for

  the selected Holding Structure, but will not use flexible remuneration or flexible equity. The Founders' economic rights, exit treatment, and governance rights shall be governed by this Founders-SC and, once applicable, by the Charter-SC, corporate, Subaccount, or holding documents adopted for the Project.

  [ Custom ]

  ‹ Describe which parts use flexible remuneration or flexible equity, and which parts use classic or corporate arrangements ›.

**Voting Weights**

: ::: Comments
  {? NOTE: Define voting weights for the Founders' Assembly during thePre-Charter Phase. The default is one vote per Founder. If customweights are agreed, list them here and in Appendix 1. ?}
  :::

  [ Default – one vote per Founder ]

  Each Founder holds one (1) vote in the Founders' Assembly during the Pre-Charter Phase.

  [ Custom ]

  The voting weights of the Founders during the Pre-Charter Phase are as specified in [Appendix 1](#appendix-1) – Register of Founders.

**Currency**
: {{ contract.currency or '‹ Currency ›' }}.

**Jurisdiction**
: {{ contract.jurisdiction or '‹ City, Country ›' }}.

**Governing Law**

: ::: Comments
  {? NOTE: Choose one governing law. French law applies by default underthe FlexUp-GC if no choice is made. ?}
  :::

  [ Default ]

  French law.

  [ Custom ]

  ‹ Applicable law ›.

**Effective Date**

: ::: Comments
  {? NOTE: Choose the effective date rule. ?}
  :::

  [ Date of last signature ]

  The Agreement enters into effect on the date of the last signature.

  [ Specific date ]

  The Agreement enters into effect on {{ contract.effective_date or '‹ Date ›' }}.

**Exit Notice Period**

: ::: Comments
  {? NOTE: Choose an exit notice period. The Founders-GC default isthirty (30) days if no period is specified here. ?}
  :::

  [ 30 days ]

  Thirty (30) days' prior written notice to the Founders' Representative.

  [ Custom ]

  ‹ Duration › prior written notice to the Founders' Representative.

**Treatment of Rights on Exit**

: Upon voluntary exit or exclusion, the exiting Founder's Credits, Tokens, shares, contractual rights, and other rights shall be:

  ::: Comments
  {? NOTE: Chooseone treatment or write a custom one. ?}
  :::

  [ Retained by exiting Founder ]

  retained by the exiting Founder, subject to the Charter.

  [ Buyback offer ]

  subject to a buyback offer by the remaining Founders at the Token Index price, within ninety (90) days of the exit date.

  [ Custom ]

  ‹ Describe treatment ›.

**Confidentiality**

: ::: Comments
  {? NOTE: Confidentiality applies by default under the NDA-GC. Select a custom option only if the Founders need a special scope, duration, or carve-out. ?}
  :::

  [ Standard confidentiality ]

  Each Founder must protect Confidential Information during the Agreement and after exit or Cessation.

  [ Custom confidentiality ]

  ‹ Describe any special confidentiality scope, authorised disclosures, duration, or carve-outs. ›

**Non-Compete Post-Exit Duration**

: ::: Comments
  {? NOTE: Choose one non-compete option. The default is "Not applicable". ?}
  :::

  [ Not applicable ]

  No non-compete applies.

  [ Narrow ]

  For six (6) months after exit or Cessation, Founders must not directly compete with the Project's core business in its active markets.

  [ Enhanced ]

  For twelve (12) months after exit or Cessation, Founders must not materially compete with the Project's products, services, market segment, or business opportunity.

  [ Custom ]

  ‹ Describe the restricted activities, territory, duration, and carve-outs. ›

**Non-Circumvention**

: ::: Comments
  {? NOTE: Choose one non-circumvention option. The default is "Standard". ?}
  :::

  [ Light ]

  For twelve (12) months after exit or Cessation, Founders must not bypass the Project to take a specific opportunity, relationship, or Project IP identified through it.

  [ Standard ]

  For twenty-four (24) months after exit or Cessation, Founders must not divert or exploit Project opportunities, relationships, or Project IP outside the agreed Project framework.

  [ Enhanced ]

  For thirty-six (36) months after exit or Cessation, Founders must not bypass, restructure, divert, or help a third party capture Project opportunities, relationships, fundraising opportunities, or Project IP.

  [ Custom ]

  ‹ Describe the restricted opportunities, relationships, duration, and carve-outs. ›

**Personnel Non-Solicitation**

: ::: Comments
  {? NOTE: Choose one personnel non-solicitation option. The default is "Standard". ?}
  :::

  [ Not applicable ]

  No personnel non-solicitation applies.

  [ Standard ]

  For twelve (12) months after exit or Cessation, Founders must not directly solicit or try to hire restricted Project or Founder personnel.

  [ Enhanced ]

  For twenty-four (24) months after exit or Cessation, Founders must not directly or indirectly solicit, hire, retain, or encourage restricted Project or Founder personnel to leave.

  [ Custom ]

  ‹ Describe the restricted persons, duration, and carve-outs. ›

**Client Non-Solicitation**

: ::: Comments
  {? NOTE: Choose one client non-solicitation option. The default is "Not applicable". ?}
  :::

  [ Not applicable ]

  No client non-solicitation applies.

  [ Standard ]For twelve (12) months after exit or Cessation, Founders must not directly solicit Project clients or prospects outside the Project or in direct competition with it.

  [ Enhanced ]

  For twenty-four (24) months after exit or Cessation, Founders must not directly or indirectly capture or help others capture Project clients, prospects, partners, investors, or commercial leads.

  [ Custom ]

  ‹ Describe the restricted relationships, duration, and carve-outs. ›

**Confidentiality Post-Exit Duration**

: ::: Comments
  {? NOTE: Choose the confidentiality survival period. The Founders-GC default is five (5) years if no period is specified here. ?}
  :::

  [ 5 years ]

  Confidentiality survives for five (5) years after exit or Cessation.

  [ Custom ]

  Confidentiality survives for ‹ Duration › after exit or Cessation.

**Initial Contributions**

: ::: Comments
  {? NOTE: Choose whether past contributions made before this Agreementare formally recognised. ?}
  :::

  [ No initial contributions ]

  No initial contributions are recognised.

  [ Initial contributions recognised ]

  Past contributions are recognised as specified in [Appendix2](#appendix-2) – Initial Contributions.
:::

### Article 4. Exceptions

4.1. The Exceptions set out in this Article ("**Exceptions**") derogate from certain provisions of the Founders-GC. Each Exception shall clearly specify the corresponding Article of the Founders-GC from which it deviates.{% if not contract.exceptions %}

4.2. There are no Exceptions to the provisions of the Founders-GC in this Founders-SC.{% else %}This Agreement contains the following Exceptions to the Founders-GC:


  {{ contract.exceptions or '‹ Contract exceptions ›' }}{% endif %}

### Article 5. Extensions

5.1. The Extensions set out in this Article ("**Extensions**") complete the stipulations of the Founders-GC, by providing additional specific conditions.{% if not contract.extensions %}

5.2. There are no Extensions to the provisions of the Founders-GC in this Founders-SC.{% else %}This Agreement contains the following Extensions to the Founders-GC:


  {{ contract.extensions or '‹ Contract extensions ›' }}{% endif %}

------------------------------------------------------------------------

By signing this document, the Founders confirm they have received, reviewed, and understood all the documents that compose the Agreement, as defined above, which together form an inseparable whole, and agree without reservation to all terms and conditions described therein.

This Founders-SC is based on the template published by FlexUp on the date specified at the top of this document.

#### List of Appendices

- Appendix 1. Register of Founders
- Appendix 2. Initial Contributions (if applicable)
- Appendix 3. Founders' Representative Mandate (if applicable)

### Signatures

::: Comments
{? NOTE: Include one signature block per Founder. In the solo-Founder case, include only one block. ?}
:::

+-----------------------------+-----------------------------+
| #### Founder 1              | #### Founder 2              |
+:============================+:============================+
| **‹ Full name ›**           | **‹ Full name ›**           |
+-----------------------------+-----------------------------+
|                             |                             |
+-----------------------------+-----------------------------+
| Date: ‹ Date of signature › | Date: ‹ Date of signature › |
+-----------------------------+-----------------------------+
|                             |                             |
+-----------------------------+-----------------------------+

+-----------------------------+-----------------------------+
| #### Founder 3              | #### Founder 4              |
+:============================+:============================+
| **‹ Full name ›**           | **‹ Full name ›**           |
+-----------------------------+-----------------------------+
|                             |                             |
+-----------------------------+-----------------------------+
| Date: ‹ Date of signature › | Date: ‹ Date of signature › |
+-----------------------------+-----------------------------+
|                             |                             |
+-----------------------------+-----------------------------+

# Appendix 1. Register of Founders

::: Comments
{ NOTE: List each Founder with their identification details. This appendix is the authoritative register of all Founders. If custom voting weights apply, add a "Voting Weight" column. }
:::

+--------+---------------+---------------------+-------------------+
| **\#** | **Full Name** | **Role / Function** | **Voting Weight** |
+:=======+:==============+:====================+===================+
| 1      | ‹ Full name › | ‹ Role ›            | ::: Comments      |
|        |               |                     | ‹ Nr of votes ›   |
|        |               |                     | :::               |
+--------+---------------+---------------------+-------------------+
| 2      | ‹ Full name › | ‹ Role ›            | ::: Comments      |
|        |               |                     | ‹ Nr of votes ›   |
|        |               |                     | :::               |
+--------+---------------+---------------------+-------------------+
| 3      | ‹ Full name › | ‹ Role ›            | ::: Comments      |
|        |               |                     | ‹ Nr of votes ›   |
|        |               |                     | :::               |
+--------+---------------+---------------------+-------------------+
| 4      | ‹ Full name › | ‹ Role ›            | ::: Comments      |
|        |               |                     | ‹ Nr of votes ›   |
|        |               |                     | :::               |
+--------+---------------+---------------------+-------------------+

# Appendix 2. Initial Contributions

::: Comments
{ NOTE: This appendix is optional. Include it only if the Founders wish to formally recognise contributions made before the signing of this Agreement. Delete this appendix if no initial contributions are to be recognised. }
:::

The following past contributions are recognised by the Founders and shall be converted into initial Credits once the Charter-SC is signed for the Project:

+--------+-------------+---------------------------------+------------+---------------------------+
| **\#** | **Founder** | **Description**                 | **Value**  | **Credit Type**           |
+:=======+:============+:================================+:===========+===========================+
| 1      | ‹ Name ›    | ‹ Description of contribution › | ‹ Amount › | ::: Comments              |
|        |             |                                 |            | ‹ Standard / Redeemable › |
|        |             |                                 |            | :::                       |
+--------+-------------+---------------------------------+------------+---------------------------+
| 2      | ‹ Name ›    | ‹ Description of contribution › | ‹ Amount › | ::: Comments              |
|        |             |                                 |            | ‹ Standard / Redeemable › |
|        |             |                                 |            | :::                       |
+--------+-------------+---------------------------------+------------+---------------------------+
| 3      | ‹ Name ›    | ‹ Description of contribution › | ‹ Amount › | ::: Comments              |
|        |             |                                 |            | ‹ Standard / Redeemable › |
|        |             |                                 |            | :::                       |
+--------+-------------+---------------------------------+------------+---------------------------+
| 4      | ‹ Name ›    | ‹ Description of contribution › | ‹ Amount › | ::: Comments              |
|        |             |                                 |            | ‹ Standard / Redeemable › |
|        |             |                                 |            | :::                       |
+--------+-------------+---------------------------------+------------+---------------------------+

[Note:]{.underline}

- "Standard" refers to "Credit (Standard)" as defined in the FlexUp Charter-GC
- "Redeemable" refers to "Credit (Redeemable Tokens)" as defined in the FlexUp Charter-GC

# Appendix 3. Founders' Representative Mandate

::: Comments
{ NOTE: This appendix is optional. Include it only if the Founders wish to grant the Founders' Representative a formal power of attorney or mandate to act on their behalf beyond the default role described in the Founders-GC – for example, to co-sign the Grouping-SC or Incubation-SC on behalf of all Founders, or to sign other documents in connection with the implementation of the Intended Structure. Delete this appendix if no mandate is required. }

{? NOTE: Describe the scope and limits of the mandate granted to the Founders' Representative. ?}
:::

The Founders grant the Founders' Representative, {{ founders.representative.name or '‹ Full name ›' }}, the authority to:

‹ Describe scope of mandate – e.g.: "sign the FlexUp Incubation Agreement on behalf of all Founders with ‹ Incubator name ›, on the terms of the draft appended hereto as Exhibit A" ›

This mandate:

  a) is limited to the specific acts described above;
  b) does not authorise the Founders' Representative to deviate from the agreed terms without the prior written consent of a Simple Majority of the Founders; and
  c) expires on ‹ date or event ›, or upon revocation by a Simple Majority of the Founders' Assembly.
