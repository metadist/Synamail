# Security Policy

We take the security of Synamail seriously. Because Synamail handles email content and
holds a per-mailbox API key to a Synaplan workspace, we ask you to report suspected
vulnerabilities **privately** so we can fix them before they are disclosed.

## Reporting a vulnerability

**Do not open a public GitHub issue for security problems.**

Instead, please use one of these private channels:

- **Preferred:** GitHub's private vulnerability reporting —
  **Security → Report a vulnerability** on this repository.
- **Email:** `security@synaplan.com` (PGP key available on request).

Please include:

- A description of the issue and its impact.
- Steps to reproduce (a proof-of-concept if possible).
- The affected version / commit and the Outlook client + platform you tested on.

## What to expect

- **Acknowledgement:** within **3 business days**.
- **Assessment & triage:** within **10 business days**, with a severity rating.
- **Fix & disclosure:** we aim to ship a fix within **90 days** of triage and will
  coordinate a disclosure timeline with you. We're happy to credit you in the release
  notes unless you prefer to remain anonymous.

## Supported versions

Security fixes are provided for the **latest released `1.x`** version. Pre-release and
unreleased branches are not covered.

| Version        | Supported |
| -------------- | --------- |
| `1.x` (latest) | ✅        |
| `< 1.0` (beta) | ❌        |

## Scope & handling notes

- **In scope:** the Synamail add-in source in this repository (taskpane, commands,
  auth/dialog flow, manifest) and its build/release pipeline.
- **Out of scope:** the Synaplan backend/API (report those to the
  [`synaplan`](../) project), Microsoft Office.js itself, and issues that require a
  malicious Outlook admin or a compromised end-user device.
- **Never** include real API keys, passwords, or another person's email content in a
  report. Redact sensitive data in screenshots and logs.

Thank you for helping keep Synamail and its users safe.
