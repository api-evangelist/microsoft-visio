# Microsoft Visio (microsoft-visio)

<!-- API-EVANGELIST-PROVENANCE:BEGIN -->
> ### About this repository
>
> **This is not our API.** This repository is an independent, third-party profile of a company's
> **publicly available** API surface, maintained by [API Evangelist](https://apievangelist.com).
> API Evangelist does not operate, host, resell, or support this company's APIs, and is not
> affiliated with or endorsed by the company unless stated on the profile.
>
> **Where the information came from.** Everything here is assembled from material a member of the
> public can reach with a browser and no credentials — the company's own website, developer portal
> and documentation, the specifications it publishes for public use (OpenAPI, AsyncAPI, JSON Schema,
> `apis.json`, `llms.txt` and similar), its public repositories, and its public status, pricing and
> changelog pages. **Nothing here is obtained by breaching a system, defeating an access control, or
> using credentials of any kind.**
>
> **The rating is an independent assessment.** The Kin Score and Agent Readiness rating are
> independently calculated scores of a company's *public* API artifacts, produced by API Evangelist
> against a published rubric. They are not certifications, endorsements, security assessments, or
> audits, and they score published artifacts — not the quality, safety, or security of the software.
>
> **Corrections, re-scores, and removal are free.** No partnership, contract, or purchase is
> required, and you do not need to justify the request.
>
> - **Something wrong?** Open an issue on this repository, or email
>   [info@apievangelist.com](mailto:info@apievangelist.com).
> - **Published something new?** Ask for a re-score and we will re-run the rating.
> - **Want the listing taken down?** Say so and we will honor it. The profile is reduced to your
>   company name, a factual description, and a link to your own site, and the company is recorded as
>   **unrated** — never scored zero for having asked.
>
> **Response times.** Acknowledgement within **one business day**; removal or restriction within
> **two business days**; corrections and re-scores within **five business days**.
>
> **Not from the company, and here with a question?** You are welcome here — we would rather be the
> front line and point you the right way than have a good report go nowhere. What this repository
> can answer is narrow, though, so it is worth knowing who you are actually looking for:
>
> - **A question about how the API works, an account, billing, or a bug in the service** — that is
>   the company's own support, not us. We profile this API; we do not operate it and cannot see
>   your account.
> - **A bug in an open-source project we only catalog** — file it on that project's own repository.
>   This has happened with a real and correct bug report that reached us instead of the people who
>   could fix it, which helped nobody.
> - **Anything about this listing itself** — the description, the tags, the rating, a missing or
>   wrong artifact — is ours. Open an issue here.
> - **Not sure, or something general about API Evangelist or APIs.io** — open an issue on the
>   [APIs.io Inbox](https://github.com/api-search/inbox) and we will route it.
>
> **This repository contains no software, and we will never ask you to download anything.** There is
> no build, release, installer, or binary here — only text and machine-readable API descriptions, so
> there is nothing here that can be "corrupt" or need "repairing". Any issue, comment, or email
> claiming otherwise and offering a download link is not from us and is hostile. Do not follow the
> link; it is a lure. Report it to GitHub and, if you like, tell us at
> [info@apievangelist.com](mailto:info@apievangelist.com) so we can take it down.
>
> **On a security or compliance team?** Email
> [info@apievangelist.com](mailto:info@apievangelist.com) with *security* in the subject line and
> you will get a person, not a form. We will tell you exactly which public URLs this profile was
> built from so your team can see the same surface we did, and we will take the listing down on
> request while you work through it.
>
> Full detail: **[Where this data comes from](https://apievangelist.com/about/where-our-data-comes-from)**
<!-- API-EVANGELIST-PROVENANCE:END -->

APIs and resources for Microsoft Visio, a diagramming and vector graphics application that helps visualize data-connected business process flows. Provides programmatic access to diagrams, pages, shapes, data items, comments, and hyperlinks through Microsoft Graph and JavaScript APIs.

**URL:** [Visit APIs.json URL](https://raw.githubusercontent.com/api-evangelist/microsoft-visio/refs/heads/main/apis.yml)

**Run:** [Capabilities Using Naftiko](https://github.com/naftiko/fleet?utm_source=api-evangelist&utm_medium=readme&utm_campaign=company-api-evangelist&utm_content=repo)

## Tags:

 - Business Process, Diagramming, Flowcharts, Microsoft 365, Visualization

## Timestamps

- **Created:** 2024
- **Modified:** 2026-04-18

## APIs

### Microsoft Graph Visio API
REST API for accessing and interacting with Visio files stored in SharePoint Online and OneDrive for Business through Microsoft Graph. Supports reading pages, shapes, shape data, comments, and hyperlinks.

**Human URL:** [https://learn.microsoft.com/en-us/graph/api/resources/visio](https://learn.microsoft.com/en-us/graph/api/resources/visio)

#### Tags:

 - Microsoft Graph, OneDrive, REST API, SharePoint, Visio Files

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/graph/api/resources/visio)
- [OpenAPI](openapi/microsoft-visio-graph-api.yaml)
- [Authentication](https://learn.microsoft.com/en-us/graph/auth/)
- [GettingStarted](https://learn.microsoft.com/en-us/graph/api/resources/visio)

### Visio JavaScript API
JavaScript API for building add-ins and extending Visio functionality in the browser with access to documents, pages, shapes, and comments.

**Human URL:** [https://learn.microsoft.com/en-us/javascript/api/visio](https://learn.microsoft.com/en-us/javascript/api/visio)

#### Tags:

 - Add-Ins, Browser, JavaScript, Office Add-Ins

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/javascript/api/visio)
- [APIReference](https://learn.microsoft.com/en-us/javascript/api/visio?view=visio-js-1.1)
- [GettingStarted](https://learn.microsoft.com/en-us/office/dev/add-ins/visio/visio-add-ins-overview)
- [CodeExamples](https://github.com/OfficeDev/Office-Add-in-samples)

## Common Properties

- [Support](https://support.microsoft.com/visio)
- [Blog](https://techcommunity.microsoft.com/t5/visio-blog/bg-p/VisioBlog)
- [PrivacyPolicy](https://privacy.microsoft.com/en-us/privacystatement)
- [TermsOfService](https://www.microsoft.com/en-us/servicesagreement)
- [StatusPage](https://status.microsoft365.com/)
- [Pricing](https://www.microsoft.com/en-us/microsoft-365/visio/microsoft-visio-plans-and-pricing-compare-visio-options)
- [GitHubOrganization](https://github.com/OfficeDev)

## Features

| Name | Description |
|------|-------------|
| Diagram Rendering | Render Visio diagrams in the browser via JavaScript API. |
| Shape Data Access | Read data items attached to diagram shapes. |
| Page Navigation | Navigate and list pages within Visio documents. |
| Comment Support | Read and manage comments on shapes. |
| Hyperlink Management | Access hyperlinks associated with diagram shapes. |

## Use Cases

| Name | Description |
|------|-------------|
| Network Topology Analysis | Programmatically analyze network diagrams for infrastructure review. |
| Business Process Review | Extract and analyze business process flow data from diagrams. |
| Compliance Auditing | Inspect diagram shapes and data for compliance validation. |

## Integrations

| Name | Description |
|------|-------------|
| SharePoint | Access Visio files stored in SharePoint document libraries. |
| OneDrive | Work with Visio diagrams in OneDrive for Business. |
| Power Automate | Trigger workflows based on Visio diagram changes. |

## Artifacts

Machine-readable API specifications organized by format.

### OpenAPI

- [Microsoft Graph Visio API](openapi/microsoft-visio-graph-api.yaml)

### JSON Schema

- [VisioPage](json-schema/visio-graph-api-visio-page-schema.json)
- [VisioShape](json-schema/visio-graph-api-visio-shape-schema.json)
- [ShapeDataItem](json-schema/visio-graph-api-shape-data-item-schema.json)
- [VisioComment](json-schema/visio-graph-api-visio-comment-schema.json)
- [VisioHyperlink](json-schema/visio-graph-api-visio-hyperlink-schema.json)

### JSON Structure

- [VisioPage](json-structure/visio-graph-api-visio-page-structure.json)
- [VisioShape](json-structure/visio-graph-api-visio-shape-structure.json)
- [ShapeDataItem](json-structure/visio-graph-api-shape-data-item-structure.json)
- [VisioComment](json-structure/visio-graph-api-visio-comment-structure.json)
- [VisioHyperlink](json-structure/visio-graph-api-visio-hyperlink-structure.json)

### JSON-LD

- [Microsoft Visio Graph API Context](json-ld/microsoft-visio-graph-api-context.jsonld)

### Examples

- [VisioPage](examples/visio-graph-api-visio-page-example.json)
- [VisioShape](examples/visio-graph-api-visio-shape-example.json)
- [ShapeDataItem](examples/visio-graph-api-shape-data-item-example.json)
- [VisioComment](examples/visio-graph-api-visio-comment-example.json)
- [VisioHyperlink](examples/visio-graph-api-visio-hyperlink-example.json)

## Capabilities

Naftiko capabilities organized as shared per-API definitions composed into customer-facing workflows.

### Shared Per-API Definitions

- [Microsoft Graph Visio API](capabilities/shared/visio-graph-api.yaml) -- 4 operations for diagram data access

### Workflow Capabilities

| Workflow | APIs Combined | Tools | Persona |
|----------|--------------|-------|---------|
| [Diagram Analysis](capabilities/diagram-analysis.yaml) | Visio Graph API | 4 | IT Architect, Business Analyst |

## Vocabulary

- [Microsoft Visio Vocabulary](vocabulary/microsoft-visio-vocabulary.yaml) -- Unified taxonomy mapping 5 resources, 2 actions, 1 workflow, and 2 personas across operational (OpenAPI) and capability (Naftiko) dimensions

## Rules

- [Microsoft Visio Spectral Rules](rules/microsoft-visio-spectral-rules.yml) -- 14 rules across 6 categories enforcing Microsoft Visio API conventions

## Maintainers

**FN:** Kin Lane

**Email:** kin@apievangelist.com
