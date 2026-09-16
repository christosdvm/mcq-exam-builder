# Security policy

## Supported version

Security fixes are applied to the latest production release.

| Version | Supported |
| --- | --- |
| 15.4.x | Yes |
| Earlier versions | No |

## Report a vulnerability

Do not disclose a suspected vulnerability in a public issue and do not attach real examinations, answers, personal data, credentials or weaponised files.

Use GitHub private vulnerability reporting when it is enabled for this repository. Include:

- the affected version;
- the security boundary that is affected;
- a minimal reproduction using synthetic data;
- the likely impact;
- a proposed mitigation, if known;
- whether the issue has been disclosed elsewhere.

Acknowledgement and remediation timing depend on severity, reproducibility and affected deployments.

## Security boundaries

The application treats uploaded documents as untrusted input.

- Modern Office containers are checked for malformed structure, member count and expanded size before parsing.
- Raster images are decoded defensively and limited to 50 megapixels.
- LibreOffice conversion uses a disposable profile, disables macros and automatic link updates, and has a timeout.
- Formula-like user text is neutralised before spreadsheet export.
- Missing or invalid images block generation instead of producing an incomplete exam.
- Generated ZIP members are created by the application; user-controlled paths are not written directly to the archive.
- Declared Python dependencies are audited for known vulnerabilities in continuous integration.

These controls reduce risk but do not turn the public deployment into an approved processor for confidential material. Institutions must decide whether hosted processing meets their own policies.

## Deployment responsibilities

- Never commit secrets, private question banks or generated examination packages.
- Review dependency changes and CI results before merging.
- Protect the production branch and require passing checks when repository settings permit it.
- Deploy a reviewed commit and retain the previous known-good commit for rollback.
- Keep the [privacy notice](PRIVACY.md) aligned with the actual hosting model.

