# Changelog

This file tracks changes specific to the `ms365-mcp-server-compact` fork.

## Unreleased

### Added

- **Contacts + People compaction (`/me/contacts`, `/me/people`).** New `projectContact`
  and `projectPerson` helpers compact Graph responses by dropping ~25 of ~40 metadata
  fields per record (changeKey, parentFolderId, @odata.etag, yomi\* fields, fax
  numbers, imAddresses, profession, manager, assistantName, spouseName, generation,
  initials, title, children for contacts; selectionLikelihood, userPrincipalName,
  imAddress, websites, postalAddresses, isFavorite, yomiCompany for people).
  Measured byte reduction: ~45–55% per record.

  Tools affected: `list-outlook-contacts`, `get-outlook-contact`, `list-relevant-people`.

  **`personType` projection.** /me/people responses preserve `personType.class` +
  `personType.subclass` for disambiguating Person vs Group entries (e.g.,
  `Person/OrganizationUser`, `Person/PersonalContact`, `Group/UnifiedGroup`).

  **`userPrincipalName` intentionally dropped from /me/people.** For
  organization users, UPN is redundant with `scoredEmailAddresses[0].address`
  in nearly all cases. For external contacts and guests, Graph returns UPN as
  an empty string and identity must come from the scored email address anyway.
  Dropping UPN keeps the projector consistent across both populations.

### Changed

- **Personal mode now derives `People.Read` scope.** The `/me/people` endpoint
  previously required `--org-mode` because its scope was tagged as a work scope
  in `endpoints.json`. It is now a personal scope and is included in the personal
  scope derivation.

  **⚠️ MIGRATION — re-auth required.** Existing personal-mode token caches do
  NOT carry `People.Read`. After updating, log out and log back in to pick up
  the new scope, otherwise `list-relevant-people` calls will fail with an
  authentication / insufficient-privileges error.

  ```bash
  # Personal mode (default)
  npx @softeria/ms-365-mcp-server --logout
  npx @softeria/ms-365-mcp-server --login
  ```

  Work/school accounts (`--org-mode`) are unaffected — `People.Read` was already
  in the org scope set.

- `tool-categories.ts`: new `people` category for filtering; the `personal`
  category pattern now also matches `people` tools so they are enabled by
  default in personal mode.
