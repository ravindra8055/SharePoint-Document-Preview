Umbraco 17 component regression suite — Spec Kit blocks (v3)
Self-contained and copy-paste ready. Supersedes v1 and v2. Changes from v2 are listed at
the end.
Context: Umbraco 17; components are a mix of backoffice-configured blocks and code-based
editors in `App_Plugins`; component source is readable via Bitbucket but not
modifiable; validation covers backoffice authoring, persisted data and rendered output;
tests run against a shared dev/staging instance that cannot be reset; Umbraco MCP and the
Playwright agent skill are available as development aids. Enterprise delivery standards
apply.
---
1. `/speckit.constitution`
```
/speckit.constitution

Create the governing principles for an enterprise component regression suite for an
Umbraco 17 backoffice. The suite validates reusable content components — some configured
as blocks in the backoffice, some built as code extensions by a separate development team.
We can read their source but cannot modify it. The suite runs against each component
release, on a shared development environment used concurrently by other people, and its
results gate releases.

ENVIRONMENT STEWARDSHIP

Principle 1 — Confined blast radius.
The target instance cannot be reset and is used by others. All test-created content lives
under a single designated fixture root. Everything outside it is read-only. The suite never
creates, modifies or deletes a document type, element type, data type, block configuration,
language, user or any other globally shared setting.

Principle 2 — The harness type is governed, not owned.
Testing components requires a container that allows them: a single permanent harness
document type, established once by agreement with the environment owner and read-only to
the suite. Because new components will need adding to it, the suite must detect when it no
longer permits every component in the inventory and report that as a blocked-coverage
condition requiring a governed change — never by modifying the type, and never by silently
skipping the component.

Principle 3 — Cleanup must survive abnormal termination, safely.
Per-test teardown is necessary but insufficient: runs get cancelled and agents die. The
suite also sweeps orphaned fixtures before starting. Sweeping must never delete fixtures
belonging to a concurrently executing run: sweep only by run identifier confirmed inactive,
with an age threshold safely above the maximum run duration.

Principle 4 — Concurrency safety.
Runs may overlap and workers run in parallel. Every fixture identifier is unique per run,
per worker and per test. No test depends on being the only thing executing, asserts on
global counts, or asserts on the absence of content it did not create.

Principle 5 — No real or personal data in fixtures.
Fixture content uses clearly synthetic values. No personal data, no customer data, no
copied production content. Fixture content may become publicly reachable and must be safe
if it does.

WHAT WE ASSERT

Principle 6 — Assert at the appropriate layer, and say which layer failed.
Each component is validated at independent layers that fail for different reasons:
authoring, persisted data, rendered output, and — once stable — appearance. Every failure
identifies its layer.

Principle 7 — Structure is the component team's to change; meaning is not.
For authoring and persistence, the assertion is the stored data contract, never the markup.
For rendering, the assertion is semantic output — visible text, alternative text, link
targets, heading levels, structured data — never CSS classes, DOM nesting or element order.
The component team will refactor markup between releases; that is legitimate and must not
fail the suite.

Principle 8 — Selector discipline.
We cannot add test attributes to components we do not own. Locators are chosen in this
order: official helper methods for core functionality; accessible labels derived from
backoffice configuration; registered custom element tag names from extension manifests;
ARIA role plus accessible name; locators scoped inside an element already found by one of
the above. Where component source is readable, locators are derived from it — element tag
names, declared ARIA roles, resolved localization keys, existing data attributes — and
recorded with the commit they came from, rather than discovered by trial. Reading the
source does not license coupling to internal DOM structure. Never XPath: it does not cross
shadow boundaries. Never generated class names or positional chains from the page root.

Principle 9 — Test what we are responsible for.
Core CMS behaviour is the vendor's responsibility and is covered by their suite. We test
the components under our remit and our configuration of core editors — the toolbars,
allowed types and constraints chosen for this build — not the core editors themselves.

Principle 10 — Source-derived expectations are characterization, not validation.
Specifications generated from component source describe current behaviour. They are valid
as regression baselines but are never evidence that behaviour is correct. Every generated
expectation is labelled with its tier: characterization, derived from implementation; or
intent, derived from a source of truth outside the implementation such as design
documentation, acceptance criteria or the content model. Where no intent source exists,
that is stated rather than inferred from code. Behaviour that appears defective is raised
as a suspected defect, never silently baselined.

GOVERNANCE

Principle 11 — Change is approved, not absorbed.
A changed data contract, rendered output or appearance is a failure until someone
accountable records it as intended. Baselines are version-controlled and updated only
through a reviewed change with a stated reason. Regenerating a baseline to make a build
green is forbidden.

Principle 12 — Every result is traceable.
A run records what it tested — component source revision, deployed build identifier,
environment, run identifier, timestamp — and results are retained. A regression report that
cannot name the build it came from is not actionable.

Principle 13 — Coverage cannot silently drift.
Coverage is measured against a generated inventory, not a maintained list. A component
present in the build with no corresponding test fails the run. A component that disappears
between releases is reported as a regression, not as a passing absence.

Principle 14 — Distinguish inconclusive from failed.
A deployment mid-run, an environment restart, an expired credential or a network fault is
reported as an inconclusive run, never as component regressions. A suite that cries wolf is
ignored, and an ignored suite gates nothing.

Principle 15 — Segregated credentials, least privilege, managed secrets.
Two distinct service accounts exist and are never conflated. A read-only account is used
for inspection, exploration and anything an AI agent can reach. A separate write-capable
account, scoped to the fixture area only, is used by the suite to create and remove
fixtures. Neither holds permissions to alter shared schema. Credentials come from the
pipeline secret store, never from the repository or a developer machine, and are rotatable
without code change. Certificate validation is never disabled against a client environment,
whatever a tool's sample configuration suggests.

Principle 16 — Agent tooling is development-time only.
AI agent tooling — browser-driving skills, CMS protocol servers, code assistants — helps us
write and diagnose tests. It is never part of a delivered artifact. The suite must execute
headless in continuous integration with no agent present, no agent protocol server running,
and no non-deterministic step. Any generated test that invokes agent tooling at runtime is
a defect. Agent tooling is additionally restricted to read-only credentials and to the
minimum tool surface needed, because an agent with write access to shared configuration
defeats Principle 1.

Principle 17 — Use the vendor's official test tooling.
Authentication, navigation and core entity setup use the official acceptance test helpers
rather than hand-rolled automation. Raw browser automation is reserved for the custom
components themselves, which the helpers cannot know about.
```
---
2. `/speckit.specify`
```
/speckit.specify

Build an enterprise regression suite for the reusable content components of an Umbraco 17
website.

CONTEXT
The site is built from reusable components — a hero banner, a configured rich text area,
and others. Some are blocks assembled from standard editors in the backoffice; some are
code extensions built by a separate development team that releases on its own schedule. We
can read their source repositories but cannot modify them.

Components are authored by content editors in the backoffice and rendered to the public
website. A regression in either place is a defect, and the two fail independently: a
component can author perfectly and render wrongly, or render correctly from data saved in
the wrong shape.

Tests run against a shared development environment used concurrently by other people. It
cannot be reset, wiped or reconfigured. Results are expected to gate component releases, so
the suite must be trustworthy enough that a red run stops a release.

USER STORIES, in priority order. Each must be independently valuable and testable.

US1 (P1) — Know what components exist, and what changed.
As a QA engineer, I obtain an authoritative inventory of every component in this build,
generated from the component source and manifests together with the deployed configuration,
rather than hand-maintained.
Acceptance: records each component's identity, type and exposed properties, and the source
revision it was derived from; diffing two builds shows components added, removed or
changed; a component present with no corresponding test fails the run; a component that
vanishes between releases is reported as a regression; where the source-declared definition
and the deployed configuration disagree, that divergence is itself reported, because each
alone is self-consistent and only the comparison reveals it.

US2 (P1) — Operate safely in a shared environment.
As the owner of the test environment, I am confident the suite cannot damage shared state or
leave debris, even when runs overlap or are cancelled.
Acceptance: all created content is confined to one designated fixture area; no globally
shared configuration is created or modified; identifiers are unique per run, worker and
test; orphaned fixtures from interrupted runs are removed, and sweeping provably cannot
delete fixtures of a concurrently executing run; fixture content is synthetic and contains
no personal or production data.

US3 (P1) — Know when coverage is blocked rather than passing.
As a QA lead, I am told when a component cannot be tested because the shared harness
container does not permit it, rather than seeing a green run that silently skipped it.
Acceptance: detected automatically; reported distinctly from both pass and fail; names what
needs changing and by whom.

US4 (P1) — Validate that each component can be authored.
As a QA engineer, I verify a content editor can add each component, complete its fields,
configure its settings and save.
Acceptance: the component is offered where it should be and not where it should not; fields
accept input of the expected kind; required fields block saving and explain why; optional
fields save when empty; the settings configuration is reachable and saves alongside the
content; the component can be removed again.

US5 (P1) — Validate the data contract each component persists.
As a QA engineer, I confirm authoring a component produces the same stored data structure
as the approved baseline, and that this matches what its source declares it should persist.
Acceptance: the stored structure per component is a reviewable, version-controlled
baseline; differences are reported field by field; content values and settings values are
covered separately; a structure that differs from the baseline fails and names the change;
a structure that differs from the source-declared contract is reported as a defect even
when it matches the baseline; this assertion does not depend on the component's markup.

US6 (P1) — Validate the rendered output.
As a QA engineer, I confirm content authored in the backoffice renders correctly, without
publishing test content to a live public site unless explicitly sanctioned.
Acceptance: authored values appear in the output; heading levels, alternative text and
responsive image sources are correct; links resolve to intended targets; emitted structured
data is well formed; components render correctly with optional fields empty and with
multiple instances on one page; assertions are on semantic output, not CSS classes or DOM
nesting.

US7 (P1) — Approve intended changes explicitly.
As a QA lead, when a component legitimately changes, I record that approval against the
specific component and the specific difference, so the baseline updates with an auditable
reason rather than being silently regenerated.
Acceptance: an unapproved difference fails; an approved one passes and is reported as an
accepted change; approvals are version-controlled, attributable and reviewable; a blanket
"accept everything" action does not exist.

US8 (P1) — Triage without reading code.
As a QA lead, I see for each component its result at each layer, and for each failure what
specifically differs, which build and source revision it was observed in, and whether the
run was conclusive.
Acceptance: one summary a non-specialist can read; failures state the layer and the observed
difference; blocked, inconclusive and failed are visually distinct from each other and from
passed.

US9 (P1) — Gate releases reliably.
As a release manager, I can treat a red run as a reason to stop a release.
Acceptance: the suite runs automatically on each component release candidate and executes
headless with no AI agent, agent protocol server or interactive tooling present; it exits
with a status the pipeline can act on; inconclusive runs do not fail the gate but are
surfaced and retried once; flaky tests are quarantined explicitly rather than retried into
silence, and the quarantine list is visible, owned and time-bound.

US10 (P2) — Detect contract drift before testing.
As a QA lead, when a component's source changes between releases, I see the difference in
its declared contract before the suite runs, so an intended change can be approved in
advance and an unintended one questioned with the component team rather than appearing as an
unexplained failure.
Acceptance: source revisions are recorded per component; a changed declared contract is
reported as a readable diff; drift is routed through the approval workflow; the report
distinguishes a changed contract from an unchanged contract with changed behaviour.

US11 (P2) — Check component accessibility.
As a QA engineer, each component is checked against automated accessibility rules in its
rendered state, because these components are reused across the site and a defect in one
propagates everywhere.
Acceptance: violations reported per component with severity; a baseline of known accepted
violations is version-controlled and governed as in US7. Automated checks only; this is not
a conformance claim.

US12 (P2) — Survive environment disruption.
As a QA engineer, a deployment, restart, credential expiry or network fault during a run is
reported as inconclusive rather than as a wave of regressions.
Acceptance: detected rather than inferred from failure patterns; the report states which
components completed before disruption.

US13 (P3) — Detect visual regressions.
As a QA engineer, once the other layers are stable, I am alerted when a component's
appearance changes unintentionally.
Acceptance: components captured on controlled fixture pages at fixed widths under
deterministic rendering conditions, with varying regions excluded; baselines governed as in
US7. Deliberately last: this layer carries the highest maintenance cost and is only worth
enabling once the layers above it are quiet.

SUCCESS CRITERIA
- Coverage is reported as a percentage of the generated inventory, and a component with no
  coverage fails the run.
- The suite leaves no content outside its fixture area, and none behind after a completed or
  interrupted run.
- Two concurrent runs complete without interfering with each other, demonstrated.
- A release in which the component team refactors markup without changing behaviour produces
  zero failures.
- A release that changes a component's stored data structure fails and names the field.
- The suite runs to completion in continuous integration with no agent tooling installed.
- Across a release cycle, fewer than 5% of reported failures prove environmental, and fewer
  than 2% of tests are quarantined for flakiness.
- Every result traces to a component source revision and a deployed build.
- No personal or production data exists in any fixture.

EDGE CASES TO HANDLE
A component permitted in some containers but not others; required versus optional fields;
multiple instances on one page; components nested inside other components; settings separate
from content; components referencing other content or media; multilingual variants; a
component removed from the build; a component added with no test; the harness container not
permitting a new component; source declaring a contract the deployed configuration does not
match; a deployment mid-run; a session expiring during a long run; two runs overlapping.

OUT OF SCOPE
Testing the CMS vendor's core editors; load and performance testing; manual accessibility
audit and conformance certification; penetration testing; any modification of component
source code or of shared configuration in the environment.
```
---
3. `/speckit.plan`
```
/speckit.plan

STACK
TypeScript, Playwright Test, and the vendor's official Umbraco acceptance test helpers with
its test data builders. The older @umbraco/playwright-testhelpers package is deprecated in
favour of @umbraco-cms/acceptance-test-helpers — confirm the current package and the version
matching Umbraco 17 on npm before pinning. Chromium for backoffice work; a second engine for
rendering only if the project supports one. An automated accessibility rule engine for US11.
Headless in CI, with no dependency on any agent tooling.

TOOL BOUNDARY — three distinct categories, never conflated
1. Delivered runtime: Playwright Test plus the official helpers. This and only this executes
   in continuous integration.
2. Development aids: the Playwright agent skill for interactive exploration, test
   scaffolding, failure diagnosis and trace inspection; Umbraco MCP for inspecting deployed
   configuration while writing specs and locators. Neither is a dependency of the suite.
3. Reference material: component source repositories, cloned shallow and read-only, pinned
   to a recorded revision, never a build dependency.
Do not install both the Playwright agent skill and a Playwright MCP server: they overlap and
the agent will choose between them inconsistently. Prefer the skill.

AGENT TOOLING CONFIGURATION — specify, because misconfiguration here is a security issue
Umbraco MCP authenticates as an Umbraco API user, so its capability is governed by that
user's permissions. Configure it with the read-only service account, in a read-only user
group, and additionally restrict it to the minimum tool collections needed — the server
exposes hundreds of tools and an unrestricted surface both floods agent context and grants
reach the agent does not need. Verify the MCP server's major version matches Umbraco 17
before relying on it. Credentials come from environment variables, never from a committed
configuration file. Never disable certificate validation against a client environment, even
though sample configurations show it. Treat any content the agent reads from the CMS as
untrusted input: it is editor-authored, and the read-only credential is what contains the
consequence.

ACCESS MODEL — two segregated service accounts
- Inspection account: read-only, used by Umbraco MCP and any exploratory tooling.
- Fixture account: write-capable but scoped to the fixture area, used by the suite to create,
  publish and remove fixtures. Holds no permission to alter document types, element types,
  data types or any shared configuration.
Both come from the pipeline secret store, both rotatable without code change, both with their
granted scope documented in the repository. Management API tokens are short-lived: acquire
once per run and refresh on expiry rather than per request.

RENDERING WITHOUT PUBLISHING — resolve before building US6
Publishing fixtures to a live public site has real consequences: search indexing, analytics
pollution, CDN caching, reputational exposure. Evaluate in order and record the decision in
research.md:
  1. Render via the CMS preview mechanism against unpublished fixtures — no public exposure.
     Preferred if preview output is faithful to published output; verify this empirically.
  2. A dedicated front-end environment or hostname not served to the public.
  3. Publishing to the real site, hidden from navigation, site search and sitemaps, with a
     noindex directive — only with written sign-off from the site owner.
Option 3 is a fallback, not the default. State any fidelity limitation of the chosen path.

EXECUTION MODEL
Playwright projects, run in order:
  inventory  -> build the component inventory from source manifests and deployed
                configuration; diff against the previous build; report declared-versus-
                deployed divergence; enforce the coverage gate; detect harness-type drift
  author     -> backoffice authoring plus persistence assertions per component
  render     -> render fixtures and assert semantic output; run accessibility rules
  visual     -> component screenshot comparison; disabled by default, enabled per config
                once the layers above are stable
Global setup: probe environment identity, acquire a token, sweep orphaned fixtures, write the
run manifest. Global teardown: remove this run's fixtures, write the summary, classify the run
outcome. Data-driven: per-component contract artifacts are read synchronously at module scope
so one test is generated per component per layer.

PER-COMPONENT ARTIFACTS
Per-component specifications are data the suite consumes, not Spec Kit features. One Spec Kit
feature governs the suite; running the specify flow per component would produce ceremony
without shared architecture. Each component has a human-readable specification with tier
labels and source citations, a machine-readable contract that drives test generation, and a
suspected-defect register. A component appearing in the inventory with no contract artifact
fails the coverage gate.

EXTRACTION
Component source is read at a pinned revision to produce the contract artifact: manifest
identity and element tag name; the declared value contract; whether every mutation path
dispatches a property-value-change event; readonly, mandatory and variant handling; validation
implementation; localization keys and their resolved labels; rendered structure, ARIA roles
and any existing data attributes. Everything extracted is labelled characterization or intent
per Principle 10. Apparent defects go to the register, never into a baseline. On each release,
diff the source revision range, re-extract what changed, and route contract differences
through the approval workflow before the suite runs.

FIXTURE LIFECYCLE
Fixtures are created under a configured root node, named with run identifier, worker index and
test identifier. Creation, publication and deletion go through the management interface, never
the UI — UI setup is slow and couples one test's arrangement to the behaviour under test.
Per-test teardown removes what that test made. Global setup sweeps leftovers, but only those
whose run identifier is recorded finished or whose age exceeds a threshold configured safely
above maximum run duration. The run manifest written at setup and cleared at teardown is what
makes safe sweeping possible under concurrency.

INCONCLUSIVE-RUN DETECTION — specify concretely, do not infer from failures
Capture an environment fingerprint at setup and again at teardown: application version or
build identifier, process start time, deployed component version. A change mid-run classifies
the run inconclusive. Also classify as inconclusive on authentication failure, repeated 5xx
from the backoffice or the rendered site, and loss of connectivity. Record which components
completed before the disruption so a retry can be scoped.

BASELINE AND APPROVAL GOVERNANCE
Baselines for data contracts, accessibility exceptions and screenshots live in the repository
alongside the tests. Updating one is a reviewed change carrying a reason and an approver.
Provide a tool to produce a proposed baseline update for a specific component, never a blanket
update switch. An unreviewed baseline change must be visible in code review as a diff a human
can read, which argues for normalised, stably ordered, human-readable serialisation rather than
opaque blobs.

MODULE BOUNDARIES
- extraction: read component repositories at a pinned revision; produce contract artifacts;
  diff against committed ones; label expectations by tier; maintain the defect register.
- inventory: merge source-derived definitions with deployed configuration; report divergence;
  diff builds; enforce the coverage gate; check harness-type drift.
- fixtures: create, publish, unpublish, delete; identifier generation; run manifest; sweep.
- authoring: per-component UI interaction built on the documented selector hierarchy.
- contract: retrieve persisted values, normalise, compare to baseline and to the declared
  contract, report field-level differences.
- rendering: obtain rendered output by the chosen mechanism and assert semantic output.
- accessibility: run rule engine per component, compare to accepted-violation baseline.
- visual: capture, mask, compare, propose baseline updates.
- reporting: per-component results across layers, run status separate from component status,
  build and source-revision traceability on every record.

SELECTOR HIERARCHY — implement as a documented shared helper, not ad hoc per test
1. Official helper methods for core functionality.
2. Accessible labels on property names, which derive from backoffice configuration and survive
   component refactors.
3. Registered custom element tag names from extension manifests — a published contract rather
   than an implementation detail.
4. ARIA role plus accessible name, preferably read from the component source or from an
   accessibility-tree snapshot rather than guessed.
5. Locators scoped inside an element already located by one of the above.
Forbidden: XPath, generated class names, positional chains from the page root.

KEY TECHNICAL DECISIONS, with rationale — carry these into research.md
1. Persistence assertions are primary for authoring; semantic assertions for rendering. The
   component team owns markup and will change it.
2. The inventory is generated, never hand-written, so coverage cannot drift and a vanished
   component is detected.
3. Source-declared contract and deployed configuration are compared against each other.
   Either alone is self-consistent by construction; only the comparison reveals the
   divergence, and that divergence is a defect class black-box testing cannot surface.
4. Fixture setup through the API, not the UI, so one component's authoring failure does not
   cascade.
5. Safe sweeping requires a run manifest. Age alone is unsafe under concurrency.
6. The harness type is read-only to the suite; drift is reported as blocked coverage. A suite
   that edits shared schema to unblock itself is a hazard in a shared environment.
7. Changes are approved per component with a recorded reason. Blanket baseline regeneration is
   the fastest way to render a regression suite meaningless.
8. Environment faults are detected by fingerprint comparison, not inferred from failure counts.
9. Agent tooling accelerates authoring tests and never executes them. The delivered suite must
   be fully deterministic and runnable with no agent installed.
10. Credentials are segregated by capability rather than shared by convenience, so that an
    agent cannot reach anything the agent should not reach.
11. Visual comparison is admissible — the same component across builds is expected to be pixel
    stable, unlike a redesign — but it is the highest-maintenance layer, ships last and off by
    default, and requires deterministic rendering: animations disabled, fonts loaded before
    capture, images settled, pinned browser version, fixed widths. Without these it generates
    noise and gets switched off, which is worse than not having it.
12. Content and settings values are asserted separately. Settings are the half that most often
    breaks silently.
13. Accessibility runs at component level because these components are reused site-wide, so one
    defect propagates. Automated rules only; not a conformance claim.

CI/CD INTEGRATION
Triggered on each component release candidate and on a schedule. The CI image contains the
suite and its runtime only — no agent tooling, no MCP server, no interactive helpers — and a
build that cannot complete without them is a defect. Exit status drives the release gate:
failed blocks, inconclusive does not block but is surfaced and retried once, blocked coverage
is surfaced to a named owner. Publish a machine-readable result artifact and a human-readable
summary per run, retained with build and source-revision identifiers for trend reporting.
Quarantined tests are declared in a version-controlled list with an owner and expiry date, and
still run and report — they simply do not gate.

CONFIGURATION
Environment variables for: backoffice base URL, rendering base URL, both service accounts'
credentials, fixture root identifier, harness type identifier, run identifier, sweep age
threshold, maximum run duration, viewport widths, baseline directories, layer enable flags,
component source checkout path, and a scoped baseline-update flag.

QUALITY GATES
Extract contracts for a handful of representative components and review the suspected-defect
register before building any test. Pilot on two components — one block, one code extension —
through authoring, persistence and rendering before extending coverage. Prove safe concurrency
by running the suite twice simultaneously. Prove agent independence by running the full suite
in a container with no agent tooling installed. Resolve the rendering-without-publishing
decision and the harness-type change process with their owners before building US6 and US3.
```
---
Changes from v2
Constitution. Principle 8 extended for source-derived locators. New Principle 10 folds in
the characterization-versus-intent rule. Principle 15 rewritten around two segregated service
accounts and an explicit ban on disabling certificate validation. New Principle 16 makes agent
tooling development-time only — previously this lived only in the plan, where Copilot forgets
it. Principle 12 now traces source revision as well as build.
Spec. US1 covers source-derived inventory and declared-versus-deployed divergence. US5
asserts observed structure against the declared contract, not only against the baseline. US9
requires the suite to run with no agent present. US10 added for contract drift. Success
criteria gained an agent-independence check and source-revision traceability.
Plan. Tool boundary split into three explicit categories. A new agent-tooling section
covering MCP scoping, version matching, secrets and the certificate-validation trap. Access
model split into inspection and fixture accounts. Extraction module added. Selector hierarchy
now prefers roles read from source or an accessibility snapshot over guessed ones. CI section
requires an image with no agent tooling, and quality gates require proving that.
Still to settle before `/speckit.plan`
Current acceptance-test-helpers package name and the version matching Umbraco 17.
Whether an Umbraco MCP release matching Umbraco 17 exists.
Whether preview rendering is faithful enough to published output for US6.
The harness document type, and who maintains it.
Security sign-off for pointing an MCP server at the client environment.
