# Herpstat Monitor Instructions

Before work, read `README.md`, `..\CODEX-WORKFLOW.md`, `..\TO-DO.md`, and the Herpstat rows in `..\CREDENTIAL-LOCATIONS.md`.

- This is the shareable GitHub project. Keep examples portable and free of personal addresses, phone numbers, device addresses, passwords, tokens, API keys, and Healthchecks URLs.
- Inspect Git status before edits and preserve unrelated changes.
- Prefer dry-run alert tests before sending live email or SMS.
- The tracked script currently accepts blank inline secret parameters. Never populate those parameters with real values; protected secret storage is the desired future design.
- Keep runtime logs, state, and generated reports out of Git.

After a completed change task, update `README.md`, document verification and unresolved work, update `..\CREDENTIAL-LOCATIONS.md` if credential storage changes, and reconcile every outstanding or completed action in `..\TO-DO.md`.
