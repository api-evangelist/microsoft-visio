# Microsoft Visio (microsoft-visio)
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
