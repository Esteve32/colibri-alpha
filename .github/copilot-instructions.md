# GitHub Copilot: Repository Custom Instructions (Colibri Style)

Project scope
- Repository: Esteve32/colibri-alpha
- Purpose: Alpha versions of browser and software elements for testing and iterative design.
- Primary languages: HTML, CSS, JavaScript

Colibri style principles
- Clarity first: prioritize simple, explicit code and comments; reduce cleverness that harms readability.
- Minimal dependencies: prefer vanilla HTML/CSS/JS; only use dependencies already present or clearly justified.
- Rapid iteration: keep examples small, isolated, and easy to try; optimize for quick feedback loops.
- Accessible and semantic: use semantic HTML and accessible ARIA patterns where appropriate.
- Consistent formatting: respect existing configs (.editorconfig, Prettier, ESLint) if present.
- Security and privacy: never include secrets; use placeholders and environment variables.
- Documentation-first: briefly document purpose and how to preview or run an example.

Authoring guidance for Copilot
- Keep changes narrowly scoped and self-contained.
- For new examples/components, include a short "How to try it" note in comments or README.
- Favor clear naming and small, composable functions; avoid over-engineering.

HTML/CSS/JS guidance
- HTML: semantic structure; prioritize a11y and readability.
- CSS: small, modular styles; avoid global overrides; prefer utility-like patterns if they exist.
- JS: vanilla JS or the minimal framework already used in this repo; avoid adding new frameworks or build steps.
- Avoid network calls or analytics unless the demo requires them; document any external dependency.

Pull requests and code review
- Keep PRs focused with imperative commit messages (e.g., "Add alpha prototype for X").
- If behavior or usage changes, update README or any example index pages.

Non-goals / guardrails
- Do not introduce heavy build systems or bundlers unless already part of the repo.
- Do not break existing demos without a clear migration note and rationale.