# SEBON IT Policy 2083 Readiness Self-Assessment

A free, tier-aware self-assessment for SEBON-licensed entities to check readiness against the **Nepal Securities Board (SEBON) Information Technology Policy, 2083** (effective 2083-07-18).

Maintained by **Topology Defense Pvt. Ltd.**, Kathmandu.

## What it does

* 57 requirements across 10 domains, each mapped to the policy clause it comes from
* Tier-aware scoring as per Clause 10.1.1 (Tier 1, Tier 2, Tier 3). Requirements that do not apply to your tier are excluded automatically
* Maturity scoring from 0 (Not in place) to 4 (Managed), plus N/A
* Domain summary with average maturity, readiness percentage and gap count
* Quick reference of reporting deadlines (incident reporting, VAPT, IT audit, BCP/DRP drills)

## Domains

|#|Domain|
|-|-|
|1|Governance and Policy Implementation|
|2|Information Security Management|
|3|Access, Authentication and Digital Signature|
|4|API Security|
|5|Third-Party and Outsourcing Risk|
|6|Business Continuity, DR and Backup|
|7|Security Testing and IT Audit|
|8|Security Monitoring and Incident Response|
|9|Data Management and Localization|
|10|AI/ML and Algorithmic Trading|

## How to use

1. Download the workbook from [Releases](../../releases) or the `tool/` folder.
2. Fill the **Profile** sheet and select your tier.
3. Score each applicable requirement on the **Assessment** sheet and record evidence.
4. Review the **Summary** sheet.

Works in Microsoft Excel and LibreOffice Calc. No macros.

## Repository layout

```
tool/      Excel workbook (download this)
source/    question-bank.csv (requirements in plain text) and the build script
```

The CSV is the readable source of the requirements, so changes are easy to review in pull requests.

## Pending SEBON directives

Some requirements depend on standards SEBON has said it will issue, such as the minimum cyber security baseline, API security standard, Data Classification Framework, System Audit Framework, SEDFAR and the integrated incident reporting portal. These are flagged in the workbook and will be updated when published.

## Disclaimer

This is an independent readiness aid. It is **not issued, reviewed or endorsed by SEBON**. A score from this tool does not represent compliance, certification or audit assurance. Requirements are paraphrased from the policy; the official Nepali text prevails in case of any difference.

## Feedback

Found a mapping error or have a suggestion? Open an [Issue](../../issues).

For a guided gap assessment, implementation roadmap or IT audit support, contact Topology Defense Pvt. Ltd. *(info@topologydefense.com / topologydefense.com)*

## License

Content licensed under [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). Share and adapt with attribution to Topology Defense Pvt. Ltd.

