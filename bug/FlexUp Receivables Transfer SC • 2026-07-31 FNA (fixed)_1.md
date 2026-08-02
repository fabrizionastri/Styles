# FlexUp Receivables Transfer Agreement – Special Conditions ("Receivables Transfer-SC")

::: Published
Based on a template published by FlexUp on ‹ date ›
:::

::: Comments
{? START NOTE ?}

Guidelines on using this template:

This template uses the following types of text:

– **Standard text** (black) – Fixed legal wording that is part of the Contract. Do not modify unless explicitly permitted.

– **Data fields** (blue) – Parameters filled automatically by the FlexUp app ({ { ... } }) or manually (‹ ... ›). When editing manually, replace any placeholder in angle brackets ‹ ... › with the appropriate value.

– **Options** (green) – Conditional sections processed by the FlexUp app ({ % ... % }) or manually ([ ... ]). When editing manually, read the condition and keep or delete the section as appropriate.

– **Guidance** (grey) {? NOTE: ... ?} – Drafting instructions and illustrations for the drafter. Not part of the Contract. Delete all guidance text before finalising the document.

If you wish to modify any standard text, you must:

– Explicitly state that changes have been made in the introduction.

– Mark any added text in underlined italics.

– Mark any deleted text with strikethrough.

Failure to declare changes to the template or standard text as indicated above constitutes a violation of the FlexUp Licence and a breach of the Charter.

– After completing your edits, remove all guidance, option blocks, and brackets to ensure the document is clean and easy to read.

{? END NOTE ?}
:::

This Receivables Transfer Agreement is entered into by and between Seller and Buyer, as defined in Article 3 below, collectively referred to as the "**Parties**" and individually as a "**Party**".

This contract is based on a template published by FlexUp on the date specified at the top of this document, and provided under the terms of the FlexUp Licence, which can be found on the FlexUp website ([www.flexup.org](http://www.flexup.org)).

### Article 1. Composition of the Contract

1.1. This Contract is composed of the following documents, listed in descending order of priority:

  a) the **Receivables Transfer-OSC**, if applicable;
  b) the **Receivables Transfer-SC**, the present document;{% if contract.charter %}
  c) the **Charter-SC**;{% endif %}
  d) the **Receivables Transfer-AC**, if applicable;
  e) the **Receivables Transfer-GC**;{% if contract.charter %}
  f) the **Charter-GC**;{% endif %}
  g) the FlexUp General Conditions ("**FlexUp-GC**").

1.2. The documents listed in Article 1.1 form an inseparable contractual whole and are collectively referred to as the "**Receivables Transfer Agreement**" or, in this document, the "**Contract**". In the event of any inconsistency between these documents, the order of priority set out in Article 1.1 shall apply.

1.3. The documents composing the General Conditions are incorporated by reference and are not required to be appended to the Contract.

1.4. The applicable versions of the FlexUp-GC and of the other documents composing the General Conditions are the latest versions published on the FlexUp website as of the date of signature of the Receivables Transfer-SC, subject to the update mechanisms described in the FlexUp-GC.

### Article 2. Scope and Obligations

2.1. Seller agrees to sell, assign, and transfer the Transferred Portion of each Receivable listed in the Receivables Schedule, and Buyer agrees to purchase it and to pay the Price, subject to the terms and conditions of this Contract.

2.2. The Receivables are transferred without recourse against Seller, unless this Contract states otherwise. Buyer acknowledges that payment of a Conditional Receivable is uncertain in amount and in timing, and that it may occur in several instalments over time or not at all.

2.3. The Debtor of each Receivable, and the Originator of each Receivable where it is not Seller, are not Parties to this Contract and do not sign it. They are identified in the Receivables Schedule for the purposes of Notification and of the Tax Collection Mandate.

### Article 3. Definitions and specific terms and conditions

3.1. Capitalised terms within this Contract are defined terms, whose definitions are given in the table below or, if not here, elsewhere in the documents composing the Contract or in the FlexUp Glossary (available on [www.flexup.org/glossary](http://www.flexup.org/glossary)).

::: Definitions

**Defined Term / Key Item**
:   **Definition / Specific terms and conditions**

**Contract Label**
:   {{ contract.text_label or '‹ Seller name › → ‹ Buyer name › – Receivables Transfer Agreement • ‹ Date ›' }}

**Seller**

:   {{ supplier.legal_identification_with_representative or

    '[ Individual ]

    ‹ full name ›, a ‹ citizenship › citizen, born on ‹ date › in ‹ city, country ›, with ‹ ID card/Passport › number ‹ identification number › and residing at ‹ address ›.

    [ Registered entity ]

    ‹ entity name ›, a ‹ legal form › entity, with a capital of ‹ capital › ‹ currency ›, registered at ‹ address › under the Trade and Companies Registry of ‹ city › with number ‹ registration number ›, represented by ‹ name › in the role of ‹ capacity ›' }}.

**Buyer**

:   {{ client.legal_identification_with_representative or

    '[ Individual ]

    ‹ full name ›, a ‹ citizenship › citizen, born on ‹ date › in ‹ city, country ›, with ‹ ID card/Passport › number ‹ identification number › and residing at ‹ address ›.

    [ Registered entity ]

    ‹ entity name ›, a ‹ legal form › entity, with a capital of ‹ capital › ‹ currency ›, registered at ‹ address › under the Trade and Companies Registry of ‹ city › with number ‹ registration number ›, represented by ‹ name › in the role of ‹ capacity ›' }}.

**Currency**
:   {{ contract.currency or '‹ Currency ›' }}.

**Jurisdiction**
:   {{ contract.jurisdiction or '‹ City, Country ›' }}.

**Effective Date**
:   The Contract enters into effect {% if contract.signature_date and contract.effective_date < contract.signature_date %}retroactively {% endif %}on {{ contract.effective_date or '‹ Effective Date ›' }}.

**The Receivables**

:   The Receivables forming the subject of the Transfer are listed individually in the Receivables Schedule set out in Appendix 1, which forms an integral part of this Contract.

    ::: Comments
    {? NOTE: State the aggregate figures below. The detail of each Receivable goes in Appendix 1. Delete any line that does not apply. ?}
    :::

    Aggregate Nominal Value of the Transferred Portions, by Priority:

    - Firm: ‹ amount › ‹ currency ›
    - Preferred, Flex, or Superflex: ‹ amount › ‹ currency ›
    - Credit: ‹ amount › ‹ currency ›
    - Token: ‹ number › Token Units

    Number of Receivables: ‹ number ›, of which ‹ number › are Payable Receivables and ‹ number › are Conditional Receivables.

**Debtors**

:   The Debtor of each Receivable is identified in the Receivables Schedule. A Receivable may have a Debtor different from that of any other Receivable.

    ::: Comments
    {? NOTE: Give the full legal identification of each Debtor here, or in Appendix 1 where there are several. This identification supports the Notification required under the Receivables Transfer-GC. ?}
    :::

    ‹ Full legal identification of each Debtor ›

**Originators**

:   ::: Comments
    {? NOTE: Choose the first option on a first transfer, where Seller originated every Receivable. Choose the second where any Receivable was acquired from someone else, so that its Originator is a third party. ?}
    :::

    [ Seller is the Originator of every Receivable ]

    Seller is the Originator of every Receivable, and accepts the obligations that the article "Tax on the Underlying Order – Tax Collection Mandate" of the Receivables Transfer-GC places on the Originator, for the benefit of Buyer and of every successive Payee.

    [ One or more Receivables have a third-party Originator ]

    The Originator of each Receivable is identified in the Receivables Schedule. Where the Originator is not Seller, Seller confirms that the Originator is bound by the obligations set out in the article "Tax on the Underlying Order – Tax Collection Mandate" of the Receivables Transfer-GC, and that Buyer may enforce them directly against it.

**Receivables under a Charter**

:   ::: Comments
    {? NOTE: This concerns the Charter of each Debtor's Project, which continues to govern the Receivable after the Transfer. It is a separate question from whether this Contract is itself an Associate Contract – see the next row. ?}
    :::

    [ None ]

    No Receivable arises under a Charter. Every Receivable is of Firm Priority.

    [ One or more Charters apply ]

    Each Receivable identified in the Receivables Schedule as arising under a Charter remains subject to the Charter of the ‹ Project name › Project after the Transfer, even though that Project is not a Party to this Contract. Buyer becomes an Associate of that Project in respect of that Receivable, and is bound by that Charter to that extent.

**Status of this Contract**
:   {% if contract.charter %}The Price is payable under a Flexible Priority subject to the Charter of the {{ contract.charter.account.name or '‹ Project Name ›' }} Project. This Contract is therefore an Associate Contract, in which Parties are Associates, and payment of the Price is conditional upon the available cash of that Project.{% else %}The Price is payable under a Firm Priority. This Contract is not an Associate Contract, notwithstanding that a Receivable may itself remain subject to a Charter.{% endif %}

**Price**

:   ‹ amount › ‹ currency ›, being the total Price for all the Receivables.

    The Receivables Schedule states the Transfer Price allocated to each Receivable. Where it does not, the Price is allocated between the Receivables in proportion to their respective Nominal Values.

**Payment Terms of the Price**

:   ::: Comments
    {? NOTE: Choose one option. Where delivery is tracked separately for each Receivable, state whether the Price falls due in full on the first Delivery Acceptance or in proportion to the Receivables transferred. ?}
    :::

    [ In full on a fixed date ]

    Buyer shall pay the Price in full to Seller on ‹ date ›.

    [ In full on Delivery Acceptance ]

    Buyer shall pay the Price in full to Seller within ‹ 15 › days of the Delivery Acceptance of the Transfer Order.

    [ Per Receivable, as each Transfer takes effect ]

    Buyer shall pay Seller the Transfer Price allocated to each Receivable within ‹ 15 › days of the Delivery Acceptance of that Receivable.

    [ Custom ]

    ‹ Describe the agreed Payment Structure of the Price ›

    Payment shall be made by bank transfer to the following account:

    - **Account name / beneficiary:** ‹ Full legal name of the account holder ›
    - **Account number / IBAN:** ‹ IBAN or account number ›
    - **Bank name:** ‹ Name of the bank ›
    - **Bank address:** ‹ Full street address of the bank branch ›
    - **BIC / SWIFT code:** ‹ 8 or 11-character code ›

**Earn-Out**

:   ::: Comments
    {? NOTE: The Earn-Out is the commercial risk-sharing supplement owed by Buyer to Seller. It is separate from, and additional to, the Tax Amount payable to the Originator under the Tax Collection Mandate – do not merge the two. ?}
    :::

    [ No Earn-Out ]

    No Earn-Out is payable. The Price is the entire consideration for the Transfer.

    [ Percentage of Collected Amounts ]

    Buyer shall pay Seller an Earn-Out equal to ‹ 50 ›% of the Net Amount of each Collected Amount, within seven (7) days of that collection, in respect of every Receivable.

    [ Per Receivable ]

    Buyer shall pay Seller the Earn-Out stated for each Receivable in the Receivables Schedule, within seven (7) days of each Collected Amount for that Receivable.

    { EXAMPLE: On a Receivable of 1 200.00 EUR nominal, comprising 1 000.00 EUR Net Amount and 200.00 EUR Tax Amount, transferred for a Transfer Price of 100.00 EUR with an Earn-Out of 50% of the Net Amount collected: if the Debtor pays 600.00 EUR, Buyer remits 100.00 EUR to the Originator under the Tax Collection Mandate, pays 250.00 EUR of Earn-Out to Seller, and retains 250.00 EUR. }

**Outstanding Earn-Outs**

:   ::: Comments
    {? NOTE: Keep this row only where a Receivable was acquired by Seller in an earlier transfer and still carries an Earn-Out owed to an earlier seller. Buyer must know about it in order to price the Receivable. ?}
    :::

    The Receivables Schedule states every Earn-Out that remains outstanding in favour of an earlier seller of a Receivable. Buyer accepts those Earn-Outs, shall pay each of them directly upon each collection, and shall disclose them to any further buyer.

**Transfer Date and Delivery**

:   ::: Comments
    {? NOTE: Choose one option. The Transfer of each Receivable takes effect on its Transfer Date, subject to Delivery Acceptance. ?}
    :::

    [ Single Transfer Date ]

    Delivery is tracked at the level of the Transfer Order. All the Receivables share the same Transfer Date, being ‹ date ›, and a single Delivery Acceptance applies to all of them.

    [ One Transfer Date per Receivable ]

    Delivery is tracked separately for each Receivable. Each Receivable has the Transfer Date stated for it in the Receivables Schedule and its own Delivery Acceptance, and the Transfer of one Receivable is independent of the Transfer of any other.

**Notification of the Debtors**

:   ::: Comments
    {? NOTE: Choose one option. Notification is what makes the Transfer enforceable against a Debtor. It is made separately for each Debtor. ?}
    :::

    [ By registration in the FlexUp App ]

    Each Debtor has access to the record of the Transfer in the FlexUp App, and registration of the Transfer constitutes Notification of that Debtor.

    [ By separate written notice ]

    Buyer shall notify each Debtor in writing within ‹ 15 › days of the Transfer Date, and Seller shall provide reasonable assistance. Seller shall countersign the notice where the applicable law so requires.

    [ By acknowledgment of the Debtor ]

    Each Debtor has acknowledged and accepted the Transfer in writing before the date of signature of this Contract, and a copy of that acknowledgment is appended to the Receivables Schedule.

**Recourse**

:   ::: Comments
    {? NOTE: Non-recourse is the default and should be kept in the great majority of cases. Recourse turns the Transfer into a guarantee by Seller and materially changes its price. ?}
    :::

    [ Without recourse (default) ]

    The Transfer is made without recourse. Seller warrants the existence and validity of each Receivable, but does not guarantee that any Debtor will pay it. The risk of non-payment passes to Buyer on the Transfer Date.

    [ With recourse ]

    The Transfer is made with recourse. Seller guarantees payment by the Debtor of ‹ identify the Receivables concerned ›, up to ‹ amount › ‹ currency ›, for a period of ‹ 12 › months from the Transfer Date. Buyer shall first demand payment from the Debtor before calling on that guarantee.

**Tax Collection Mandate**

:   ::: Comments
    {? NOTE: The default below is the mechanism set out in the Receivables Transfer-GC and should be kept unless a Receivables Transfer-AC provides an alternative for the relevant jurisdiction. Only select the second option on specific tax advice. ?}
    :::

    [ Default mandate (recommended) ]

    The mechanism set out in the article "Tax on the Underlying Order – Tax Collection Mandate" of the Receivables Transfer-GC applies without modification. Buyer acquires the Net Amount as assignee and collects the Tax Amount as mandatary of the Originator, remitting it to the Originator within seven (7) days of each collection.

    [ Alternative provided by the Receivables Transfer-AC ]

    Parties select the following mechanism, provided by the applicable Receivables Transfer-AC, in place of the default: ‹ identify the mechanism and the Receivables Transfer-AC that provides it ›.

    [ No Tax Amount ]

    No Receivable includes a Tax Amount, and that article does not apply.

**Allocation Rule**

:   Each Collected Amount is apportioned between the Net Amount and the Tax Amount in proportion to the amounts that each of them represents in the corresponding invoice issued by the Originator or, where no invoice has yet been issued, in the Receivable itself.

    ::: Comments
    {? NOTE: Keep the standard rule above unless a different apportionment is agreed, in which case state it per Receivable in the Receivables Schedule. ?}
    :::

**Compliance with Transfer Restrictions**

:   ::: Comments
    {? NOTE: Keep this row only where Receivables of Token Priority are included in the Transfer, since Tokens carry governance rights under the Charter. Choose as applicable. ?}
    :::

    [ Unrestricted – 2% ]

    As per the Charter, this Transfer, considered with all Related Transfers, is not subject to restrictions, since it represents less than 2% of the total number of Token Units in all Active Tokens.

    [ Unrestricted – 5% + Council approval ]

    As per the Charter, this Transfer, considered with all Related Transfers, is not subject to restrictions, since it represents less than 5% of the total number of Token Units in all Active Tokens, and the Council approved it.

    [ Restricted ]

    This Transfer has complied with the procedure set out in the Charter for Restricted Transfers.
:::

### Article 1. Exceptions

1.1. The Exceptions set out in this Article ("**Exceptions**") derogate from certain provisions of the Receivables Transfer-GC. Each Exception shall clearly specify the corresponding Article of the Receivables Transfer-GC from which it deviates. Parties may not derogate from the obligations that the Receivables Transfer-GC attaches to a Receivable in respect of Earn-Outs and of the Tax Collection Mandate.{% if not contract.exceptions %}

1.2. There are no Exceptions to the provisions of the Receivables Transfer-GC in this Receivables Transfer-SC.{% else %}This Contract contains the following Exceptions to the Receivables Transfer-GC:

  {{ contract.exceptions or '‹ Contract exceptions ›' }}{% endif %}

### Article 2. Extensions

2.1. The Extensions set out in this Article ("**Extensions**") complete the stipulations of the Receivables Transfer-GC, by providing additional specific conditions.{% if not contract.extensions %}

2.2. There are no Extensions to the provisions of the Receivables Transfer-GC in this Receivables Transfer-SC.{% else %}This Contract contains the following Extensions to the Receivables Transfer-GC:

  {{ contract.extensions or '‹ Contract extensions ›' }}{% endif %}

### Article 3. Parameters

3.1. The Parameters set out in this Article ("**Parameters**") modify specific default values established in the General Conditions applicable to this Contract.{% if not contract.parameters %}

3.2. There are no Parameters to the General Conditions in this Receivables Transfer-SC.{% else %}This Contract modifies the following default values of the General Conditions:

  {{ contract.parameters or '‹ Parameter name › ( ‹ article number › ) : ‹ applicable value › (instead of the ‹ default value ›)' }}{% endif %}

By signing this document, the Parties confirm they have received, reviewed, and understood all the documents that compose the Contract, as defined above, which together form an inseparable whole, and agree without reservation to all terms and conditions described therein.

::: Comments
{? NOTE: Confirm whether you modified the standard text. ?}
:::

[ No modification to standard text ]

The template has not been modified in any other way.

[ Modifications made ]

Further modifications have been made: additions are shown in underlined italics and deletions with a strikethrough.

::: Comments
{? END NOTE ?}
:::

[List of Appendices:]{.underline}

- Appendix 1. Receivables Schedule{% if contract.charter %}
- Appendix 2. FlexUp Charter Special Conditions (Charter-SC){% endif %}

### Signatures

+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| [Signature of Seller]{.underline}                                                                                               | [Signature of Buyer]{.underline}                                                                                            |
+=================================================================================================================================+=============================================================================================================================+
|                                                                                                                                 |                                                                                                                             |
+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| **{{ supplier.name }}**                                                                                                         | **{{ client.name }}**                                                                                                       |
+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| Date: {{ supplier_signature_date or '‹ Date of signature ›' }}                                                                  | Date: {{ client_signature_date or '‹ Date of signature ›' }}                                                                |
+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| Name: {{ supplier.main_representative.name or '‹ Seller name ›' }}                                                              | Name: {{ client.main_representative.name or '‹ Buyer name ›' }}                                                             |
+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+
| {% if supplier.main_representative_capacity %}Role: {{ supplier.main_representative_capacity or '‹ Role as representative ›' }} | {% if client.main_representative_capacity %}Role: {{ client.main_representative_capacity or '‹ Role as representative ›' }} |
|                                                                                                                                 |                                                                                                                             |
| {% endif %}                                                                                                                     | {% endif %}                                                                                                                 |
+---------------------------------------------------------------------------------------------------------------------------------+-----------------------------------------------------------------------------------------------------------------------------+

# Appendix 1. Receivables Schedule

::: Comments
{? NOTE: One row per Receivable. Each row identifies one Commitment, which is transferred together with all its subsequent Iterations and Residues. Where a Receivable arises under a Charter, state the Project in the "Underlying Order" column. The FlexUp App populates this table from the transferred tranches of the Transfer Order. ?}
:::

| No. | Underlying Order                         | Originator | Debtor   | Priority                    | Nominal Value           | Transfer % | Transfer Price          | Earn-Out basis                           |
|-----|------------------------------------------|------------|----------|-----------------------------|-------------------------|------------|-------------------------|------------------------------------------|
| 1   | ‹ reference, date, and Project, if any › | ‹ name ›   | ‹ name › | ‹ Firm / Credit / Token … › | ‹ amount › ‹ currency › | ‹ 100 ›%   | ‹ amount › ‹ currency › | ‹ % of Net Amount collected, or "none" › |
| 2   |                                          |            |          |                             |                         |            |                         |                                          |
| 3   |                                          |            |          |                             |                         |            |                         |                                          |

::: Comments
{? NOTE: Add the columns below only where they apply. ?}
:::

- **Transfer Date** – required only where delivery is tracked separately for each Receivable.
- **Tax Amount** – the part of the Nominal Value representing value added tax or an equivalent turnover tax, where any Receivable includes one.
- **Outstanding Earn-Outs** – for each Receivable acquired by Seller in an earlier transfer, the Earn-Out still owed to each earlier seller, and that seller's identity.
- **Status** – whether the Receivable is a Payable Receivable or a Conditional Receivable at the Transfer Date.

For each Receivable, Parties shall keep with this Contract the documents supporting it, including the Underlying Order, the Delivery Declarations, the Statements, and any invoice already issued.
